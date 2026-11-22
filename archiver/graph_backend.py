"""Microsoft 365 online mode: reads the mailbox through Microsoft Graph.

Works whatever Outlook app is used (the new Outlook, classic Outlook, or
none), because it talks to the mailbox in Microsoft 365 directly. It needs
a sign-in with the work account. It cannot create PST files (only classic
Outlook can); it saves the original items instead: emails as .eml,
calendar items as .ics, contacts as .vcf.
"""
from __future__ import annotations

import datetime as dt
import html
import os
import re
import tempfile
import time
from html.parser import HTMLParser
from typing import Iterator

from .base import Backend, BackendError, Cancelled, FolderNode, IsCancelled, Progress
from .core import Criteria, ItemInfo, PathAllocator, matches_deep, matches_quick, safe_name

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["User.Read", "Mail.ReadWrite", "Calendars.Read", "Contacts.Read", "Tasks.Read"]
SIZE_PROP = "Integer 0x0E08"  # PR_MESSAGE_SIZE: the item's size in bytes, attachments included
EXPAND_SIZE = f"singleValueExtendedProperties($filter=id eq '{SIZE_PROP}')"
MSG_FIELDS = ("id,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,"
              "createdDateTime,lastModifiedDateTime,categories,hasAttachments,isDraft")
EVENT_FIELDS = "id,subject,organizer,start,end,createdDateTime,lastModifiedDateTime,categories,hasAttachments,type"
UTC_HEADER = {"Prefer": 'outlook.timezone="UTC"'}


# ------------------------------------------------------------------ sign-in

class MsalAuth:
    """Signs in with the work account (Microsoft's own sign-in page) and keeps the session."""

    def __init__(self, client_id: str, tenant: str, cache_path: str, t) -> None:
        import msal

        self.t = t
        self.cache = _token_cache(cache_path)
        self.app = msal.PublicClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant or 'organizations'}",
            token_cache=self.cache)

    def token(self, interactive: bool = True) -> str:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        if not result and interactive:
            result = self.app.acquire_token_interactive(SCOPES, prompt="select_account", timeout=300)
        if not result or "access_token" not in result:
            detail = (result or {}).get("error_description") or (result or {}).get("error") or ""
            raise BackendError(self.t("e_signin", err=detail.split("\n")[0][:300]))
        _save_cache(self.cache)
        return result["access_token"]

    def sign_out(self) -> None:
        for acc in self.app.get_accounts():
            self.app.remove_account(acc)
        _save_cache(self.cache)


def _token_cache(path: str):
    """On Windows the session is stored encrypted for this Windows user only (DPAPI)."""
    try:
        from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache
        if os.name == "nt":
            return PersistedTokenCache(FilePersistenceWithDataProtection(path))
    except ImportError:
        pass
    import msal
    return msal.SerializableTokenCache()  # kept in memory only


def _save_cache(cache) -> None:
    pass  # PersistedTokenCache saves itself; the in-memory cache is not written to disk


class GraphHttp:
    """Small Microsoft Graph client: retries when Microsoft asks us to slow down."""

    def __init__(self, auth: MsalAuth, t) -> None:
        import requests

        self.auth = auth
        self.t = t
        self.s = requests.Session()

    def _call(self, method: str, path: str, params=None, headers=None, json=None, raw=False):
        url = path if path.startswith("http") else GRAPH + path
        for attempt in range(6):
            h = {"Authorization": "Bearer " + self.auth.token()}
            h.update(headers or {})
            try:
                r = self.s.request(method, url, params=params if not path.startswith("http") else None,
                                   headers=h, json=json, timeout=120)
            except Exception as e:  # noqa: BLE001 - network problems
                if attempt == 5:
                    raise BackendError(self.t("e_network", err=str(e)[:200])) from e
                time.sleep(2 * (attempt + 1))
                continue
            if r.status_code in (429, 503, 504):
                time.sleep(min(60, int(r.headers.get("Retry-After") or 5)))
                continue
            if r.status_code in (401, 403):
                raise BackendError(self.t("e_denied", err=_graph_error(r)))
            if r.status_code >= 400:
                raise BackendError(self.t("e_graph", code=r.status_code, err=_graph_error(r)))
            if raw:
                return r.content
            return r.json() if r.content else {}
        raise BackendError(self.t("e_graph", code=429, err="busy"))

    def get(self, path, params=None, headers=None):
        return self._call("GET", path, params, headers)

    def get_bytes(self, path):
        return self._call("GET", path, raw=True)

    def post(self, path, body):
        return self._call("POST", path, json=body)


def _graph_error(r) -> str:
    try:
        return r.json()["error"]["message"][:300]
    except Exception:  # noqa: BLE001
        return r.text[:300]


# --------------------------------------------------------------- data source

class GraphBackend(Backend):
    label = "Microsoft 365"
    supports_pst = False
    kinds = ("mail", "meeting", "appointment", "contact", "task", "other")
    formats = ("eml", "mht", "pdf", "txt")

    def __init__(self, http=None, client_id: str = "", tenant: str = "", cache_path: str = "") -> None:
        super().__init__()
        self.http = http
        self._auth_args = (client_id, tenant, cache_path)
        self.account = ""
        self._folders: dict[str, tuple[str, str, tuple[str, ...]]] = {}  # key -> (type, id, parts)
        self._protected: set[str] = set()
        self._tmp: str | None = None
        self._word = None

    # ------------------------------------------------------------ connect
    def connect(self) -> str:
        if self.http is None:
            client_id, tenant, cache_path = self._auth_args
            if not client_id:
                raise BackendError(self.t("e_no_client_id"))
            try:
                auth = MsalAuth(client_id, tenant, cache_path, self.t)
            except ImportError as e:
                raise BackendError(self.t("e_signin", err=str(e))) from e
            self.http = GraphHttp(auth, self.t)
        me = self.http.get("/me", {"$select": "displayName,mail,userPrincipalName"})
        self.account = me.get("mail") or me.get("userPrincipalName") or me.get("displayName") or "Microsoft 365"
        return self.t("connected_online", who=self.account)

    def sign_out(self) -> None:
        auth = getattr(self.http, "auth", None)
        if auth:
            auth.sign_out()
        self.http = None

    def _pages(self, path: str, params=None, headers=None) -> Iterator[dict]:
        data = self.http.get(path, params, headers)
        while True:
            yield from data.get("value", [])
            nxt = data.get("@odata.nextLink")
            if not nxt:
                return
            data = self.http.get(nxt, None, headers)

    # ------------------------------------------------------------- folders
    def list_tree(self) -> list[FolderNode]:
        self._folders.clear()
        self._protected.clear()
        name = self.account
        root = FolderNode(key="store1", name=name, label=name, is_store=True)
        try:
            deleted_id = self.http.get("/me/mailFolders/deleteditems", {"$select": "id"}).get("id")
        except BackendError:
            deleted_id = None

        def mail_nodes(path: str, parts: tuple, protected: bool) -> list[FolderNode]:
            out = []
            for f in self._pages(path, {"$top": "250", "$select": "id,displayName,totalItemCount,childFolderCount,sizeInBytes"}):
                p = parts + (f.get("displayName") or "?",)
                is_prot = protected or f["id"] == deleted_id
                key = self._remember("mail", f["id"], p)
                if is_prot:
                    self._protected.add(key)
                node = FolderNode(key=key, name=p[-1], label="/".join((name,) + p),
                                  count=int(f.get("totalItemCount") or 0), size=f.get("sizeInBytes"))
                if f.get("childFolderCount"):
                    node.children = mail_nodes(f"/me/mailFolders/{f['id']}/childFolders", p, is_prot)
                out.append(node)
            return out

        root.children = mail_nodes("/me/mailFolders", (), False)
        for cal in self._pages("/me/calendars", {"$select": "id,name"}):
            p = (cal.get("name") or "Calendar",)
            root.children.append(FolderNode(key=self._remember("calendar", cal["id"], p), name=p[0],
                                            label="/".join((name,) + p), count=-1))
        root.children.append(FolderNode(key=self._remember("contacts", "", ("Contacts",)), name="Contacts",
                                        label=f"{name}/Contacts", count=-1))
        for cf in self._pages("/me/contactFolders", {"$select": "id,displayName"}):
            p = ("Contacts", cf.get("displayName") or "?")
            root.children[-1].children.append(FolderNode(
                key=self._remember("contacts", cf["id"], p), name=p[-1], label="/".join((name,) + p), count=-1))
        try:
            lists = list(self._pages("/me/todo/lists", {"$select": "id,displayName"}))
        except BackendError:
            lists = []  # To Do is not available for every account
        if lists:
            tasks = FolderNode(key=self._remember("group", "", ("Tasks",)), name="Tasks", label=f"{name}/Tasks", count=-1)
            for lst in lists:
                p = ("Tasks", lst.get("displayName") or "?")
                tasks.children.append(FolderNode(key=self._remember("todo", lst["id"], p), name=p[-1],
                                                 label="/".join((name,) + p), count=-1))
            root.children.append(tasks)
        sizes = [c.size for c in root.children if c.size is not None]
        root.size = _total_size(root.children) if sizes else None
        return [root]

    def _remember(self, kind: str, fid: str, parts: tuple) -> str:
        key = f"g{len(self._folders)}"
        self._folders[key] = (kind, fid, parts)
        return key

    def folder_label(self, key: str) -> str:
        return "/".join((self.account,) + self._folders[key][2])

    # ---------------------------------------------------------------- scan
    def iter_folder(self, key: str, criteria: Criteria, progress: Progress,
                    cancelled: IsCancelled) -> Iterator[ItemInfo]:
        kind, fid, parts = self._folders[key]
        label = self.folder_label(key)
        if kind == "mail":
            params = {"$select": MSG_FIELDS, "$top": "200", "$expand": EXPAND_SIZE}
            flt = _date_filter("receivedDateTime", criteria) if criteria.date_field in ("auto", "received") else ""
            if flt:
                params["$filter"] = flt
            rows = self._pages(f"/me/mailFolders/{fid}/messages", params)
            make = self._message
        elif kind == "calendar":
            rows = self._pages(f"/me/calendars/{fid}/events",
                               {"$select": EVENT_FIELDS, "$top": "200", "$expand": EXPAND_SIZE}, UTC_HEADER)
            make = self._event
        elif kind == "contacts":
            path = f"/me/contactFolders/{fid}/contacts" if fid else "/me/contacts"
            rows = self._pages(path, {"$select": "id,displayName,createdDateTime,lastModifiedDateTime,categories", "$top": "200"})
            make = self._contact
        elif kind == "todo":
            rows = self._pages(f"/me/todo/lists/{fid}/tasks", {"$top": "200"})
            make = self._task
        else:
            return
        n = 0
        for row in rows:
            if cancelled():
                raise Cancelled()
            n += 1
            if n % 200 == 0:
                progress(self.t("p_looking_n", folder=label, n=n, total=n), n, 0)
            item = make(row, key, parts)
            if not matches_quick(item, criteria):
                continue
            if criteria.needs_item_open:
                body = self._body_text(item) if criteria.body_words.strip() else ""
                if not matches_deep(item.subject, body, item.attachments or 0, criteria):
                    continue
            yield item

    def _base(self, row: dict, key: str, parts: tuple, mclass: str, **kw) -> ItemInfo:
        return ItemInfo(entry_id=row["id"], store_id="graph", folder_key=key, folder_parts=parts,
                        store_name=self.account, message_class=mclass,
                        categories=", ".join(row.get("categories") or []),
                        created=_when(row.get("createdDateTime")), modified=_when(row.get("lastModifiedDateTime")),
                        size=_size(row), **kw)

    def _message(self, row: dict, key: str, parts: tuple) -> ItemInfo:
        meeting = "eventMessage" in (row.get("@odata.type") or "")
        who = (row.get("from") or row.get("sender") or {}).get("emailAddress") or {}
        recips = [(r.get("emailAddress") or {}).get("name") or (r.get("emailAddress") or {}).get("address", "")
                  for r in (row.get("toRecipients") or []) + (row.get("ccRecipients") or [])]
        return self._base(row, key, parts, "IPM.Schedule.Meeting.Request" if meeting else "IPM.Note",
                          subject=row.get("subject") or "", sender=who.get("name") or who.get("address") or "",
                          recipients="; ".join(r for r in recips if r),
                          received=None if row.get("isDraft") else _when(row.get("receivedDateTime")),
                          sent=None if row.get("isDraft") else _when(row.get("sentDateTime")),
                          attachments=1 if row.get("hasAttachments") else 0)

    def _event(self, row: dict, key: str, parts: tuple) -> ItemInfo:
        org = ((row.get("organizer") or {}).get("emailAddress") or {})
        return self._base(row, key, parts, "IPM.Appointment", subject=row.get("subject") or "",
                          sender=org.get("name") or org.get("address") or "",
                          start=_when((row.get("start") or {}).get("dateTime"), utc=True),
                          end=_when((row.get("end") or {}).get("dateTime"), utc=True),
                          recurring=row.get("type") == "seriesMaster",
                          attachments=1 if row.get("hasAttachments") else 0)

    def _contact(self, row: dict, key: str, parts: tuple) -> ItemInfo:
        return self._base(row, key, parts, "IPM.Contact", subject=row.get("displayName") or "", attachments=0)

    def _task(self, row: dict, key: str, parts: tuple) -> ItemInfo:
        return self._base(row, key, parts, "IPM.Task", subject=row.get("title") or "", attachments=0)

    # --------------------------------------------------------------- items
    def _item_path(self, item: ItemInfo) -> str:
        kind, fid, _ = self._folders[item.folder_key]
        if kind == "calendar":
            return f"/me/events/{item.entry_id}"
        if kind == "contacts":
            return f"/me/contacts/{item.entry_id}"
        if kind == "todo":
            return f"/me/todo/lists/{fid}/tasks/{item.entry_id}"
        return f"/me/messages/{item.entry_id}"

    def open_item(self, item: ItemInfo):
        headers = UTC_HEADER if self._folders[item.folder_key][0] == "calendar" else None
        return self.http.get(self._item_path(item), None, headers)  # the full item, as Graph sends it

    def attachment_count(self, handle) -> int:
        return 1 if handle.get("hasAttachments") else 0

    def _body_text(self, item: ItemInfo) -> str:
        try:
            data = self.http.get(self._item_path(item), {"$select": "body"} if item.kind in ("mail", "meeting") else None)
        except BackendError:
            return ""
        return _text_of(data.get("body") or {})

    # ----------------------------------------------------------- no PST here
    def pst_begin(self, path: str, display_name: str) -> str:
        raise BackendError(self.t("e_no_pst_online"))

    def pst_copy(self, handle, item: ItemInfo, include_store: bool) -> None:
        raise BackendError(self.t("e_no_pst_online"))

    def pst_inventory(self) -> list[ItemInfo]:
        return []

    def pst_end(self, keep_open: bool) -> None:
        pass

    # --------------------------------------------------------------- files
    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        if fmt == "eml":
            return self._save_original(handle, item, directory, stem, alloc)
        if fmt == "mht":
            path = alloc.allocate(directory, stem, ".html")
            _write(path, _html_page(handle, item).encode("utf-8"))
            return path
        if fmt == "txt":
            path = alloc.allocate(directory, stem, ".txt")
            _write(path, _text_page(handle, item).encode("utf-8-sig"))
            return path
        if fmt == "pdf":
            return self._save_pdf(handle, item, directory, stem, alloc)
        raise BackendError(self.t("e_format_online", fmt=fmt.upper()))

    def _save_original(self, handle, item: ItemInfo, directory: str, stem: str, alloc: PathAllocator) -> str:
        if item.kind in ("mail", "meeting", "other"):
            path = alloc.allocate(directory, stem, ".eml")
            _write(path, self.http.get_bytes(f"/me/messages/{item.entry_id}/$value"))
        elif item.kind == "appointment":
            path = alloc.allocate(directory, stem, ".ics")
            _write(path, _ics(handle).encode("utf-8"))
        elif item.kind == "contact":
            path = alloc.allocate(directory, stem, ".vcf")
            _write(path, _vcf(handle).encode("utf-8"))
        else:  # tasks: the standard calendar-task format
            path = alloc.allocate(directory, stem, ".ics")
            _write(path, _ics_todo(handle).encode("utf-8"))
        return path

    def _save_pdf(self, handle, item: ItemInfo, directory: str, stem: str, alloc: PathAllocator) -> str:
        if os.name != "nt":
            raise BackendError(self.t("e_no_word"))
        if self._word is None:
            try:
                import win32com.client
                self._word = win32com.client.DispatchEx("Word.Application")
                self._word.Visible = False
                self._word.DisplayAlerts = 0
            except Exception as e:  # noqa: BLE001
                raise BackendError(self.t("e_no_word")) from e
        if self._tmp is None:
            self._tmp = tempfile.mkdtemp(prefix="outlook-archiver-")
        tmp = os.path.join(self._tmp, "item.html")
        _write(tmp, _html_page(handle, item).encode("utf-8"))
        path = alloc.allocate(directory, stem, ".pdf")
        doc = self._word.Documents.Open(tmp, False, True, False)
        try:
            doc.ExportAsFixedFormat(path, 17)
        finally:
            doc.Close(0)
        return path

    def save_attachments(self, handle, item: ItemInfo, directory: str, alloc: PathAllocator) -> list[str]:
        base = self._item_path(item)
        if item.kind not in ("mail", "meeting", "appointment", "other"):
            return []
        saved = []
        att_dir = os.path.join(directory, self.t("attachments_dir", name=safe_name(item.subject, 50)))
        for a in self._pages(f"{base}/attachments", {"$select": "id,name,isInline,size"}):
            if a.get("isInline"):
                continue  # pictures inside the email body
            name = a.get("name") or "attachment"
            stem, ext = os.path.splitext(name)
            if "itemAttachment" in (a.get("@odata.type") or "") and not ext:
                ext = ".eml"
            os.makedirs(att_dir, exist_ok=True)
            path = alloc.allocate(att_dir, safe_name(stem, 80, "attachment"), ext[:10])
            _write(path, self.http.get_bytes(f"{base}/attachments/{a['id']}/$value"))
            saved.append(path)
        return saved

    # -------------------------------------------------------------- removal
    def can_remove(self, item: ItemInfo) -> bool:
        # Online, only emails and meeting messages are removed (moved to Deleted Items).
        return item.kind in ("mail", "meeting", "other") and self._folders[item.folder_key][0] == "mail"

    def in_protected_folder(self, item: ItemInfo) -> bool:
        return item.folder_key in self._protected

    def remove_item(self, item: ItemInfo) -> None:
        self.http.post(f"/me/messages/{item.entry_id}/move", {"destinationId": "deleteditems"})

    def finish_export(self) -> None:
        if self._word is not None:
            try:
                self._word.Quit(0)
            except Exception:  # noqa: BLE001
                pass
            self._word = None
        if self._tmp:
            for name in os.listdir(self._tmp):
                try:
                    os.remove(os.path.join(self._tmp, name))
                except OSError:
                    pass
            try:
                os.rmdir(self._tmp)
            except OSError:
                pass
            self._tmp = None


# ------------------------------------------------------------------ helpers

def _total_size(nodes) -> int:
    return sum((n.size or 0) + _total_size(n.children) for n in nodes)


def _size(row: dict) -> int:
    for p in row.get("singleValueExtendedProperties") or []:
        if p.get("id", "").lower() == SIZE_PROP.lower():
            try:
                return max(0, int(p.get("value") or 0))
            except ValueError:
                return 0
    return 0


def _when(value, utc: bool = True) -> dt.datetime | None:
    """Graph times are UTC; show them in this laptop's local time, like Outlook does."""
    if not value:
        return None
    try:
        d = dt.datetime.strptime(str(value)[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None
    if d.year < 1950 or d.year > 4000:
        return None
    return d.replace(tzinfo=dt.timezone.utc).astimezone().replace(tzinfo=None)


def _date_filter(field: str, c: Criteria) -> str:
    """A slightly wider date range for Microsoft to filter on; the exact check happens here."""
    parts = []
    if c.date_from:
        a = dt.datetime.combine(c.date_from, dt.time()) - dt.timedelta(days=1)
        parts.append(f"{field} ge {a:%Y-%m-%dT%H:%M:%SZ}")
    if c.date_to:
        b = dt.datetime.combine(c.date_to, dt.time()) + dt.timedelta(days=2)
        parts.append(f"{field} lt {b:%Y-%m-%dT%H:%M:%SZ}")
    return " and ".join(parts)


class _Text(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.out: list[str] = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head"):
            self.skip += 1
        if tag in ("br", "p", "div", "tr", "li", "h1", "h2", "h3"):
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head") and self.skip:
            self.skip -= 1

    def handle_data(self, data):
        if not self.skip:
            self.out.append(data)


def _text_of(body: dict) -> str:
    content = body.get("content") or ""
    if (body.get("contentType") or "").lower() != "html":
        return content
    p = _Text()
    p.feed(content)
    return re.sub(r"\n{3,}", "\n\n", "".join(p.out)).strip()


def _people(lst) -> str:
    out = []
    for r in lst or []:
        e = r.get("emailAddress") or {}
        name, addr = e.get("name") or "", e.get("address") or ""
        out.append(f"{name} <{addr}>" if name and addr and name != addr else (name or addr))
    return "; ".join(x for x in out if x)


def _header_rows(h: dict, item: ItemInfo) -> list[tuple[str, str]]:
    rows = []
    who = h.get("from") or h.get("organizer")
    if who:
        rows.append(("From", _people([who])))
    if h.get("toRecipients"):
        rows.append(("To", _people(h["toRecipients"])))
    if h.get("ccRecipients"):
        rows.append(("Cc", _people(h["ccRecipients"])))
    if item.kind == "appointment":
        s, e = item.start, item.end
        rows.append(("When", f"{s:%Y-%m-%d %H:%M} - {e:%Y-%m-%d %H:%M}" if s and e else ""))
        loc = (h.get("location") or {}).get("displayName")
        if loc:
            rows.append(("Where", loc))
    else:
        d = item.display_date
        rows.append(("Date", f"{d:%Y-%m-%d %H:%M}" if d else ""))
    return rows


def _html_page(h: dict, item: ItemInfo) -> str:
    body = h.get("body") or {}
    if (body.get("contentType") or "").lower() == "html":
        content = body.get("content") or ""
    else:
        content = "<pre style='white-space:pre-wrap;font:inherit'>" + html.escape(_item_text(h, item)) + "</pre>"
    head = "".join(f"<tr><th>{html.escape(k)}</th><td>{html.escape(v)}</td></tr>" for k, v in _header_rows(h, item))
    return ("<!doctype html><html><head><meta charset='utf-8'><title>" + html.escape(item.subject) + "</title>"
            "<style>body{font-family:Segoe UI,Tahoma,sans-serif;margin:24px}table.h{border-collapse:collapse;"
            "margin-bottom:16px}table.h th{text-align:start;padding:2px 12px 2px 0;color:#555}"
            "h1{font-size:20px}</style></head><body><h1>" + html.escape(item.subject or "(no subject)") +
            "</h1><table class='h'>" + head + "</table><hr>" + content + "</body></html>")


def _item_text(h: dict, item: ItemInfo) -> str:
    if item.kind == "contact":
        lines = [h.get("displayName") or ""]
        lines += [e.get("address", "") for e in h.get("emailAddresses") or []]
        lines += list(h.get("businessPhones") or []) + ([h["mobilePhone"]] if h.get("mobilePhone") else [])
        lines += [h.get("companyName") or "", h.get("jobTitle") or ""]
        return "\n".join(x for x in lines if x)
    if item.kind == "task":
        lines = [h.get("title") or "", f"Status: {h.get('status', '')}"]
        due = (h.get("dueDateTime") or {}).get("dateTime")
        if due:
            lines.append(f"Due: {due[:10]}")
        return "\n".join(lines + ["", _text_of(h.get("body") or {})])
    return _text_of(h.get("body") or {})


def _text_page(h: dict, item: ItemInfo) -> str:
    head = "\n".join(f"{k}: {v}" for k, v in [("Subject", item.subject)] + _header_rows(h, item))
    return head + "\n\n" + _item_text(h, item) + "\n"


def _ics_text(s: str) -> str:
    return (s or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\r", "").replace("\n", "\\n")


def _ics_time(value) -> str:
    try:
        d = dt.datetime.strptime(str(value)[:19], "%Y-%m-%dT%H:%M:%S")
    except (TypeError, ValueError):
        return ""
    return d.strftime("%Y%m%dT%H%M%SZ")


def _ics(h: dict) -> str:
    org = (h.get("organizer") or {}).get("emailAddress") or {}
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Outlook Archiver//EN", "BEGIN:VEVENT",
             f"UID:{h.get('iCalUId') or h.get('id')}", f"DTSTAMP:{dt.datetime.now(dt.timezone.utc):%Y%m%dT%H%M%SZ}",
             f"DTSTART:{_ics_time((h.get('start') or {}).get('dateTime'))}",
             f"DTEND:{_ics_time((h.get('end') or {}).get('dateTime'))}",
             f"SUMMARY:{_ics_text(h.get('subject') or '')}"]
    if org.get("address"):
        lines.append(f"ORGANIZER;CN={_ics_text(org.get('name') or '')}:mailto:{org['address']}")
    for a in h.get("attendees") or []:
        e = a.get("emailAddress") or {}
        if e.get("address"):
            lines.append(f"ATTENDEE;CN={_ics_text(e.get('name') or '')}:mailto:{e['address']}")
    loc = (h.get("location") or {}).get("displayName")
    if loc:
        lines.append(f"LOCATION:{_ics_text(loc)}")
    lines.append(f"DESCRIPTION:{_ics_text(_text_of(h.get('body') or {}))}")
    lines += ["END:VEVENT", "END:VCALENDAR"]
    return "\r\n".join(lines) + "\r\n"


def _ics_todo(h: dict) -> str:
    status = {"completed": "COMPLETED", "inProgress": "IN-PROCESS"}.get(h.get("status") or "", "NEEDS-ACTION")
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Outlook Archiver//EN", "BEGIN:VTODO",
             f"UID:{h.get('id')}", f"DTSTAMP:{dt.datetime.now(dt.timezone.utc):%Y%m%dT%H%M%SZ}",
             f"SUMMARY:{_ics_text(h.get('title') or '')}", f"STATUS:{status}"]
    due = (h.get("dueDateTime") or {}).get("dateTime")
    if due:
        lines.append(f"DUE:{_ics_time(due)}")
    lines.append(f"DESCRIPTION:{_ics_text(_text_of(h.get('body') or {}))}")
    lines += ["END:VTODO", "END:VCALENDAR"]
    return "\r\n".join(lines) + "\r\n"


def _vcf(h: dict) -> str:
    def esc(s):
        return (s or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
    lines = ["BEGIN:VCARD", "VERSION:3.0", f"FN:{esc(h.get('displayName'))}",
             f"N:{esc(h.get('surname'))};{esc(h.get('givenName'))};;;"]
    for e in h.get("emailAddresses") or []:
        if e.get("address"):
            lines.append(f"EMAIL;TYPE=INTERNET:{e['address']}")
    for p in h.get("businessPhones") or []:
        lines.append(f"TEL;TYPE=WORK:{p}")
    if h.get("mobilePhone"):
        lines.append(f"TEL;TYPE=CELL:{h['mobilePhone']}")
    if h.get("companyName"):
        lines.append(f"ORG:{esc(h['companyName'])}")
    if h.get("jobTitle"):
        lines.append(f"TITLE:{esc(h['jobTitle'])}")
    lines.append("END:VCARD")
    return "\r\n".join(lines) + "\r\n"


def _write(path: str, data: bytes) -> None:
    with open(path, "wb") as f:
        f.write(data)
