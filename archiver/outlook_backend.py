"""Reads and copies items through classic Outlook's automation (COM) interface.

Windows only. Needs classic Outlook (the one with a File menu) installed;
it does not need to be open - Windows starts it in the background.
"""
from __future__ import annotations

import os
import tempfile
from typing import Iterator

import pythoncom
import pywintypes
import win32com.client

from .base import Backend, BackendError, Cancelled, FolderNode, IsCancelled, Progress
from .core import Criteria, ItemInfo, PathAllocator, clean_date, matches_deep, matches_quick, safe_name

# Outlook constants (numbers so we do not depend on generated type files)
OL_STORE_UNICODE = 3
OL_CREATE_APPOINTMENT = 1
OL_TXT, OL_HTML, OL_VCARD, OL_ICAL, OL_MSG_UNICODE, OL_MHTML = 0, 5, 6, 8, 9, 10
OL_BY_VALUE, OL_EMBEDDED_ITEM = 1, 5
WD_EXPORT_PDF = 17

# Folder.DefaultItemType -> folder type to create in the PST
PST_FOLDER_TYPE = {0: 6, 1: 9, 2: 10, 3: 13, 4: 11, 5: 12, 6: 6, 7: 10}

TABLE_COLUMNS = [
    "EntryID", "MessageClass", "Subject", "SenderName", "Organizer", "To", "CC",
    "Categories", "ReceivedTime", "SentOn", "CreationTime", "LastModificationTime", "Start",
]

PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}

NO_CLASSIC_OUTLOOK = (
    "This tool could not reach classic Outlook on this laptop.\n\n"
    "It works with classic Outlook (the version with a 'File' menu at the top left).\n\n"
    "* If you see a 'New Outlook' switch at the top right of Outlook, turn it OFF to go back "
    "to classic Outlook, then run this tool again.\n"
    "* If classic Outlook is not installed, install the Microsoft 365 desktop apps "
    "(Outlook is included) from https://portal.office.com, open it once to sign in, "
    "then run this tool again."
)


def _err(e: Exception) -> str:
    if isinstance(e, pywintypes.com_error):
        try:
            detail = e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else e.strerror
        except Exception:  # noqa: BLE001
            detail = None
        return str(detail or e)
    return str(e)


class OutlookBackend(Backend):
    label = "Outlook"

    def __init__(self) -> None:
        self.app = None
        self.ns = None
        # key -> (entry_id, store_id, store_name, folder_parts, default_item_type)
        self._folders: dict[str, tuple[str, str, str, tuple[str, ...], int]] = {}
        self._pst_root = None
        self._pst_folders: dict[tuple, object] = {}
        self._word = None
        self._word_failed = False
        self._tmp = None

    # ------------------------------------------------------------ connect
    def connect(self) -> str:
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.ns = self.app.GetNamespace("MAPI")
            count = self.ns.Stores.Count
        except pywintypes.com_error as e:
            raise BackendError(NO_CLASSIC_OUTLOOK + f"\n\n(Technical details: {_err(e)})") from e
        return f"Connected to Outlook - {count} mailbox(es) / data file(s) found."

    # -------------------------------------------------------- folder tree
    def list_tree(self) -> list[FolderNode]:
        self._folders.clear()
        nodes: list[FolderNode] = []
        stores = self.ns.Stores
        for i in range(1, stores.Count + 1):
            try:
                store = stores.Item(i)
                root = store.GetRootFolder()
                name = store.DisplayName
                sid = store.StoreID
            except pywintypes.com_error:
                continue
            node = FolderNode(key=f"store{i}", name=name, is_store=True)
            node.children = self._children(root, sid, name, (), 0)
            nodes.append(node)
        return nodes

    def _children(self, folder, sid: str, sname: str, parts: tuple, depth: int) -> list[FolderNode]:
        out: list[FolderNode] = []
        if depth > 15:
            return out
        try:
            folders = folder.Folders
            n = folders.Count
        except pywintypes.com_error:
            return out
        for i in range(1, n + 1):
            try:
                f = folders.Item(i)
                name = f.Name
                eid = f.EntryID
                try:
                    count = f.Items.Count
                except pywintypes.com_error:
                    count = 0
                try:
                    item_type = int(f.DefaultItemType)
                except pywintypes.com_error:
                    item_type = 0
            except pywintypes.com_error:
                continue
            key = f"f{len(self._folders)}"
            p = parts + (name,)
            self._folders[key] = (eid, sid, sname, p, item_type)
            node = FolderNode(key=key, name=name, count=count)
            node.children = self._children(f, sid, sname, p, depth + 1)
            out.append(node)
        return out

    def folder_label(self, key: str) -> str:
        _, _, sname, parts, _ = self._folders[key]
        return "/".join((sname,) + parts)

    # --------------------------------------------------------------- scan
    def iter_folder(self, key: str, criteria: Criteria, progress: Progress,
                    cancelled: IsCancelled) -> Iterator[ItemInfo]:
        eid, sid, sname, parts, _ = self._folders[key]
        try:
            folder = self.ns.GetFolderFromID(eid, sid)
            table = folder.GetTable("", 0)  # 0 = normal (not hidden) items
            cols = table.Columns
            cols.RemoveAll()
        except pywintypes.com_error as e:
            progress(f"Skipped {self.folder_label(key)}: {_err(e)}", 0, 0)
            return
        added: list[str] = []
        for name in TABLE_COLUMNS:
            try:
                cols.Add(name)
                added.append(name)
            except pywintypes.com_error:
                pass
        if "EntryID" not in added:
            return
        need_recips_from_item = bool(criteria.recipients.strip()) and "To" not in added
        total = table.GetRowCount()
        label = self.folder_label(key)
        n = 0
        while not table.EndOfTable:
            if cancelled():
                raise Cancelled()
            row = table.GetNextRow()
            if row is None:
                break
            n += 1
            if n % 250 == 0:
                progress(f"Looking in {label} ... {n:,} of {total:,}", n, total)
            v = dict(zip(added, row.GetValues()))
            item = ItemInfo(
                entry_id=str(v.get("EntryID") or ""),
                store_id=sid,
                folder_key=key,
                folder_parts=parts,
                store_name=sname,
                message_class=str(v.get("MessageClass") or ""),
                subject=str(v.get("Subject") or ""),
                sender=str(v.get("SenderName") or v.get("Organizer") or ""),
                recipients="; ".join(str(x) for x in (v.get("To"), v.get("CC")) if x),
                categories=str(v.get("Categories") or ""),
                received=clean_date(v.get("ReceivedTime")),
                sent=clean_date(v.get("SentOn")),
                created=clean_date(v.get("CreationTime")),
                modified=clean_date(v.get("LastModificationTime")),
                start=clean_date(v.get("Start")),
            )
            if not item.entry_id:
                continue
            opened = None
            if need_recips_from_item:
                try:
                    opened = self.open_item(item)
                    item.recipients = "; ".join(
                        str(x) for x in (getattr(opened, "To", ""), getattr(opened, "CC", "")) if x)
                except pywintypes.com_error:
                    pass
            if not matches_quick(item, criteria):
                continue
            if criteria.needs_item_open:
                try:
                    if opened is None:
                        opened = self.open_item(item)
                    body = _get(opened, "Body", "")
                    item.attachments = self.attachment_count(opened)
                except pywintypes.com_error:
                    continue
                if not matches_deep(item.subject, body, item.attachments, criteria):
                    continue
            yield item

    # ------------------------------------------------------------- items
    def open_item(self, item: ItemInfo):
        return self.ns.GetItemFromID(item.entry_id, item.store_id)

    def attachment_count(self, handle) -> int:
        try:
            return int(handle.Attachments.Count)
        except (pywintypes.com_error, AttributeError):
            return 0

    # --------------------------------------------------------------- PST
    def pst_begin(self, path: str, display_name: str) -> str:
        self._pst_folders.clear()
        try:
            self.ns.AddStoreEx(path, OL_STORE_UNICODE)
        except pywintypes.com_error as e:
            raise BackendError(f"Outlook could not create the PST file {path}: {_err(e)}") from e
        target = os.path.normcase(os.path.abspath(path))
        stores = self.ns.Stores
        for i in range(1, stores.Count + 1):
            s = stores.Item(i)
            try:
                fp = s.FilePath
            except pywintypes.com_error:
                continue
            if fp and os.path.normcase(os.path.abspath(fp)) == target:
                self._pst_root = s.GetRootFolder()
                try:
                    self._pst_root.Name = display_name
                except pywintypes.com_error:
                    pass
                return path
        raise BackendError("Outlook created the PST file but could not open it.")

    def _pst_folder(self, item: ItemInfo, include_store: bool):
        _, _, sname, parts, item_type = self._folders[item.folder_key]
        path = ((sname,) if include_store else ()) + parts
        if path in self._pst_folders:
            return self._pst_folders[path]
        folder = self._pst_root
        for depth, name in enumerate(path):
            sub = path[: depth + 1]
            if sub in self._pst_folders:
                folder = self._pst_folders[sub]
                continue
            is_last = depth == len(path) - 1
            ftype = PST_FOLDER_TYPE.get(item_type, 6) if is_last else 6
            folder = self._get_or_add(folder, name, ftype)
            self._pst_folders[sub] = folder
        return folder

    @staticmethod
    def _get_or_add(parent, name: str, ftype: int):
        folders = parent.Folders
        for i in range(1, folders.Count + 1):
            f = folders.Item(i)
            if f.Name.lower() == name.lower():
                return f
        return folders.Add(name, ftype)

    def pst_copy(self, handle, item: ItemInfo, include_store: bool) -> None:
        dest = self._pst_folder(item, include_store)
        if item.kind == "appointment":
            try:
                handle.CopyTo(dest, OL_CREATE_APPOINTMENT)  # keeps the subject without "Copy:"
                return
            except (pywintypes.com_error, AttributeError):
                pass
        copy = handle.Copy()
        try:
            copy.Move(dest)
        except pywintypes.com_error:
            try:
                copy.Delete()
            except pywintypes.com_error:
                pass
            raise

    def pst_end(self, keep_open: bool) -> None:
        if self._pst_root is not None and not keep_open:
            self.ns.RemoveStore(self._pst_root)
        self._pst_root = None
        self._pst_folders.clear()

    # ------------------------------------------------------------- files
    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        if fmt == "msg":
            path = alloc.allocate(directory, stem, ".msg")
            handle.SaveAs(path, OL_MSG_UNICODE)
            return path
        if fmt == "txt":
            path = alloc.allocate(directory, stem, ".txt")
            handle.SaveAs(path, OL_TXT)
            return path
        if fmt == "mht":
            return self._save_readable(handle, item, directory, stem, alloc)
        if fmt == "pdf":
            return self._save_pdf(handle, item, directory, stem, alloc)
        raise ValueError(f"Unknown format {fmt}")

    def _save_readable(self, handle, item: ItemInfo, directory: str, stem: str,
                       alloc: PathAllocator) -> str:
        """A copy that opens without Outlook (web page, calendar or contact card)."""
        tries = [(".mht", OL_MHTML)]
        if item.kind == "appointment":
            tries = [(".ics", OL_ICAL)] + tries
        elif item.kind == "contact":
            tries = [(".vcf", OL_VCARD)] + tries
        tries.append((".txt", OL_TXT))
        last: Exception | None = None
        for ext, kind in tries:
            path = alloc.allocate(directory, stem, ext)
            try:
                handle.SaveAs(path, kind)
                return path
            except pywintypes.com_error as e:
                last = e
        raise last  # type: ignore[misc]

    def _word_app(self):
        if self._word is None and not self._word_failed:
            try:
                self._word = win32com.client.DispatchEx("Word.Application")
                self._word.Visible = False
                self._word.DisplayAlerts = 0
            except pywintypes.com_error:
                self._word_failed = True
        return self._word

    def _save_pdf(self, handle, item: ItemInfo, directory: str, stem: str,
                  alloc: PathAllocator) -> str:
        word = self._word_app()
        if word is None:
            raise BackendError("PDF files were skipped: Microsoft Word is not installed on this laptop.")
        if self._tmp is None:
            self._tmp = tempfile.mkdtemp(prefix="outlook-archiver-")
        tmp = os.path.join(self._tmp, "item.mht")
        if os.path.exists(tmp):
            os.remove(tmp)
        try:
            handle.SaveAs(tmp, OL_MHTML)
        except pywintypes.com_error:
            tmp = os.path.join(self._tmp, "item.txt")
            if os.path.exists(tmp):
                os.remove(tmp)
            handle.SaveAs(tmp, OL_TXT)
        path = alloc.allocate(directory, stem, ".pdf")
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        doc = word.Documents.Open(tmp, False, True, False)
        try:
            doc.ExportAsFixedFormat(path, WD_EXPORT_PDF)
        finally:
            doc.Close(0)
        return path

    def save_attachments(self, handle, item: ItemInfo, directory: str,
                         alloc: PathAllocator) -> list[str]:
        saved: list[str] = []
        atts = handle.Attachments
        att_dir = os.path.join(directory, safe_name(item.subject, 50) + " - attachments")
        for i in range(1, atts.Count + 1):
            a = atts.Item(i)
            try:
                if int(a.Type) not in (OL_BY_VALUE, OL_EMBEDDED_ITEM):
                    continue
                filename = str(a.FileName or a.DisplayName or f"attachment {i}")
            except pywintypes.com_error:
                continue
            stem, ext = os.path.splitext(filename)
            if int(a.Type) == OL_EMBEDDED_ITEM and not ext:
                ext = ".msg"
            if _is_inline_image(a, ext):
                continue  # logos and pictures inside the email body
            os.makedirs(att_dir, exist_ok=True)
            path = alloc.allocate(att_dir, safe_name(stem, 80, f"attachment {i}"), ext[:10])
            a.SaveAsFile(path)
            saved.append(path)
        return saved

    def finish_export(self) -> None:
        if self._word is not None:
            try:
                self._word.Quit(0)
            except pywintypes.com_error:
                pass
        self._word = None
        self._word_failed = False
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


def _get(obj, attr: str, default=None):
    try:
        return getattr(obj, attr)
    except (pywintypes.com_error, AttributeError):
        return default


def _is_inline_image(att, ext: str) -> bool:
    try:
        pa = att.PropertyAccessor
    except pywintypes.com_error:
        return False
    try:
        if pa.GetProperty(PR_ATTACHMENT_HIDDEN):
            return True
    except pywintypes.com_error:
        pass
    if ext.lower() in IMAGE_EXTS:
        try:
            return bool(pa.GetProperty(PR_ATTACH_CONTENT_ID))
        except pywintypes.com_error:
            return False
    return False
