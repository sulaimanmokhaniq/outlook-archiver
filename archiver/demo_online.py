"""A pretend Microsoft 365 mailbox, answering like Microsoft Graph does.

Used by `--demo` (online mode) and the tests, so the real online code runs
end to end without an account or internet. Built from the same made-up
data as the classic demo.
"""
from __future__ import annotations

import datetime as dt
import re
from urllib.parse import parse_qs, urlencode, urlparse

from .base import BackendError
from .demo_backend import DEMO_FOLDERS, DemoBackend
from .graph_backend import GraphBackend, SIZE_PROP

PAGE = 50


def _utc(d: dt.datetime | None) -> str | None:
    if d is None:
        return None
    return d.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


class FakeGraph:
    def __init__(self) -> None:
        demo = DemoBackend()
        self.account = "manager@example.com"
        self.folders: dict[str, dict] = {}  # id -> {name, parent, items(kind=mail)}
        self.messages: dict[str, dict] = {}
        self.events: dict[str, dict] = {}
        self.contacts: dict[str, dict] = {}
        self.tasks: dict[str, dict] = {}
        self.calls = 0
        ids = {}
        for idx, (parts, _mc, _n) in enumerate(DEMO_FOLDERS):
            entries = demo._items[f"f{idx}"]
            top = parts[0]
            if top in ("Calendar",):
                for info, body, atts in entries:
                    self.events[f"ev{len(self.events)}"] = self._event(info, body, atts)
                continue
            if top == "Contacts":
                for info, _b, _a in entries:
                    self.contacts[f"ct{len(self.contacts)}"] = self._contact(info)
                continue
            if top == "Tasks":
                for info, body, _a in entries:
                    self.tasks[f"tk{len(self.tasks)}"] = self._task(info, body)
                continue
            fid = "mf-" + "-".join(p.lower().replace(" ", "") for p in parts)
            ids[parts] = fid
            parent = ids.get(parts[:-1]) if len(parts) > 1 else None
            self.folders[fid] = {"name": parts[-1], "parent": parent}
            for info, body, atts in entries:
                mid = f"m{len(self.messages)}"
                self.messages[mid] = self._message(mid, fid, info, body, atts)

    # ---- data in Graph's shape
    @staticmethod
    def _size(n):
        return [{"id": SIZE_PROP, "value": str(n)}]

    def _message(self, mid, fid, info, body, atts):
        meeting = info.kind == "meeting"
        draft = info.received is None
        return {
            "@odata.type": "#microsoft.graph.eventMessage" if meeting else "#microsoft.graph.message",
            "id": mid, "parentFolderId": fid, "subject": info.subject,
            "from": {"emailAddress": {"name": info.sender, "address": info.sender.lower().replace(" ", ".") + "@example.com"}},
            "toRecipients": [{"emailAddress": {"name": info.recipients, "address": "x@example.com"}}],
            "ccRecipients": [], "receivedDateTime": _utc(info.received or info.created),
            "sentDateTime": _utc(info.sent or info.created), "createdDateTime": _utc(info.created),
            "lastModifiedDateTime": _utc(info.modified), "categories": [info.categories] if info.categories else [],
            "hasAttachments": atts > 0, "isDraft": draft,
            "singleValueExtendedProperties": self._size(info.size),
            "body": {"contentType": "html", "content": "<p>" + body.replace("\n", "<br>") + "</p>"},
            "_atts": atts,
        }

    def _event(self, info, body, atts):
        return {
            "id": f"ev{len(self.events)}", "subject": info.subject,
            "organizer": {"emailAddress": {"name": info.sender, "address": "org@example.com"}},
            "start": {"dateTime": (_utc(info.start) or "")[:19] + ".0000000", "timeZone": "UTC"},
            "end": {"dateTime": (_utc(info.end) or "")[:19] + ".0000000", "timeZone": "UTC"},
            "createdDateTime": _utc(info.created), "lastModifiedDateTime": _utc(info.modified),
            "categories": [], "hasAttachments": atts > 0, "type": "seriesMaster" if info.recurring else "singleInstance",
            "singleValueExtendedProperties": self._size(info.size),
            "body": {"contentType": "text", "content": body}, "location": {"displayName": "Meeting room 2"},
            "_atts": atts,
        }

    def _contact(self, info):
        first, _, last = info.subject.partition(" ")
        return {"id": f"ct{len(self.contacts)}", "displayName": info.subject, "givenName": first, "surname": last,
                "emailAddresses": [{"address": info.subject.lower().replace(" ", ".") + "@example.com"}],
                "mobilePhone": "+966 50 000 0000", "createdDateTime": _utc(info.created),
                "lastModifiedDateTime": _utc(info.modified), "categories": []}

    def _task(self, info, body):
        return {"id": f"tk{len(self.tasks)}", "title": info.subject, "status": "notStarted",
                "createdDateTime": _utc(info.created), "lastModifiedDateTime": _utc(info.modified),
                "body": {"contentType": "text", "content": body}}

    # ---- the "network"
    def _page(self, rows: list[dict], path: str, params: dict | None) -> dict:
        params = dict(params or {})
        skip = int(params.pop("$skip", 0) or 0)
        out = {"value": [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows[skip:skip + PAGE]]}
        if skip + PAGE < len(rows):
            params["$skip"] = str(skip + PAGE)
            out["@odata.nextLink"] = "https://fake.graph" + path + "?" + urlencode(params)
        return out

    def get(self, path, params=None, headers=None):
        self.calls += 1
        if path.startswith("http"):
            u = urlparse(path)
            path = u.path
            params = {k: v[0] for k, v in parse_qs(u.query).items()}
        m = re.fullmatch
        if path == "/me":
            return {"displayName": "Manager", "mail": self.account}
        if path == "/me/mailFolders/deleteditems":
            return {"id": "mf-deleteditems"}
        if path == "/me/mailFolders":
            return self._page([self._folder_json(f) for f, d in self.folders.items() if d["parent"] is None], path, params)
        if g := m(r"/me/mailFolders/([^/]+)/childFolders", path):
            return self._page([self._folder_json(f) for f, d in self.folders.items() if d["parent"] == g[1]], path, params)
        if g := m(r"/me/mailFolders/([^/]+)/messages", path):
            return self._page([r for r in self.messages.values() if r["parentFolderId"] == g[1]], path, params)
        if g := m(r"/me/messages/([^/]+)", path):
            return self._strip(self._get(self.messages, g[1]))
        if g := m(r"/me/(messages|events)/([^/]+)/attachments", path):
            src = self._get(self.messages if g[1] == "messages" else self.events, g[2])
            return {"value": [{"@odata.type": "#microsoft.graph.fileAttachment", "id": f"a{i}",
                               "name": f"document {i + 1}.pdf", "isInline": False, "size": 1000}
                              for i in range(src["_atts"])]}
        if path == "/me/calendars":
            return {"value": [{"id": "cal-main", "name": "Calendar"}]}
        if path == "/me/calendars/cal-main/events":
            return self._page(list(self.events.values()), path, params)
        if g := m(r"/me/events/([^/]+)", path):
            return self._strip(self._get(self.events, g[1]))
        if path == "/me/contactFolders":
            return {"value": []}
        if path == "/me/contacts":
            return self._page(list(self.contacts.values()), path, params)
        if g := m(r"/me/contacts/([^/]+)", path):
            return self._get(self.contacts, g[1])
        if path == "/me/todo/lists":
            return {"value": [{"id": "todo-main", "displayName": "Tasks"}]}
        if path == "/me/todo/lists/todo-main/tasks":
            return self._page(list(self.tasks.values()), path, params)
        if g := m(r"/me/todo/lists/todo-main/tasks/([^/]+)", path):
            return self._get(self.tasks, g[1])
        raise BackendError(f"fake graph: no route for GET {path}")

    def get_bytes(self, path):
        self.calls += 1
        if g := re.fullmatch(r"/me/messages/([^/]+)/\$value", path):
            msg = self._get(self.messages, g[1])
            text = re.sub("<[^>]+>", "\n", msg["body"]["content"])
            return ("From: {f}\r\nSubject: {s}\r\nDate: {d}\r\nMIME-Version: 1.0\r\n"
                    "Content-Type: text/plain; charset=utf-8\r\n\r\n{t}\r\n").format(
                f=msg["from"]["emailAddress"]["address"], s=msg["subject"], d=msg["sentDateTime"], t=text).encode()
        if re.fullmatch(r"/me/(messages|events)/[^/]+/attachments/[^/]+/\$value", path):
            return b"%PDF-1.4 demo attachment\n"
        raise BackendError(f"fake graph: no route for bytes {path}")

    def post(self, path, body):
        self.calls += 1
        if g := re.fullmatch(r"/me/messages/([^/]+)/move", path):
            msg = self._get(self.messages, g[1])
            msg["parentFolderId"] = "mf-" + body["destinationId"]
            return {"id": msg["id"]}
        raise BackendError(f"fake graph: no route for POST {path}")

    def _folder_json(self, fid):
        msgs = [r for r in self.messages.values() if r["parentFolderId"] == fid]
        return {"id": fid, "displayName": self.folders[fid]["name"], "totalItemCount": len(msgs),
                "childFolderCount": sum(1 for d in self.folders.values() if d["parent"] == fid),
                "sizeInBytes": sum(int(r["singleValueExtendedProperties"][0]["value"]) for r in msgs)}

    @staticmethod
    def _get(table, key):
        if key not in table:
            raise BackendError(f"fake graph: {key} not found")
        return table[key]

    @staticmethod
    def _strip(row):
        return {k: v for k, v in row.items() if not k.startswith("_")}


class DemoOnlineBackend(GraphBackend):
    label = "Demo Microsoft 365"
    is_demo = True

    def __init__(self) -> None:
        super().__init__(http=FakeGraph())
