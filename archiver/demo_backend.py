"""Made-up mailbox so the app can be tried (and tested) without Outlook.

Start the app with --demo. PST files and .msg/PDF files need real Outlook,
so in demo mode those are written as small text placeholders.
"""
from __future__ import annotations

import datetime as dt
import os
import random
import time
from typing import Iterator

from .base import Backend, Cancelled, FolderNode, IsCancelled, Progress
from .core import Criteria, ItemInfo, PathAllocator, matches_deep, matches_quick, safe_name

PEOPLE = ["Ahmed Ali", "Sara Hassan", "Omar Khalid", "Mona Youssef", "Finance Team", "IT Support"]
TOPICS = ["Budget", "Quarterly report", "Board meeting", "Contract renewal",
          "Travel plan", "Project Alpha", "Invoice", "Weekly update", "Hiring",
          "تقرير المشروع", "اجتماع الإدارة"]
DEMO_FOLDERS = [
    (("Inbox",), "IPM.Note", 700),
    (("Inbox", "Projects"), "IPM.Note", 150),
    (("Sent Items",), "IPM.Note", 400),
    (("Drafts",), "IPM.Note", 25),
    (("Calendar",), "IPM.Appointment", 300),
    (("Contacts",), "IPM.Contact", 60),
    (("Tasks",), "IPM.Task", 40),
    (("Deleted Items",), "IPM.Note", 30),
]


class DemoBackend(Backend):
    label = "Demo data"
    is_demo = True

    def __init__(self) -> None:
        super().__init__()
        self.store = "Demo Mailbox"
        # folder key -> list of (info, body, attachment count)
        self._items: dict[str, list[tuple[ItemInfo, str, int]]] = {}
        self._copied = 0
        rnd = random.Random(7)
        start = dt.datetime(2017, 1, 1)
        for idx, (parts, mclass, count) in enumerate(DEMO_FOLDERS):
            key = f"f{idx}"
            items = []
            for n in range(count):
                when = start + dt.timedelta(minutes=rnd.randint(0, 60 * 24 * 365 * 10))
                topic = rnd.choice(TOPICS)
                mc = mclass
                if parts == ("Inbox",) and n % 15 == 0:
                    mc = "IPM.Schedule.Meeting.Request"
                person = rnd.choice(PEOPLE)
                outgoing = parts[0] in ("Sent Items", "Drafts")
                undated_received = parts[0] in ("Drafts", "Calendar", "Contacts", "Tasks")
                info = ItemInfo(
                    entry_id=f"{key}-{n}", store_id="demo", folder_key=key,
                    folder_parts=parts, store_name=self.store, message_class=mc,
                    subject=person if mclass == "IPM.Contact" else f"{topic} #{n}",
                    sender="Me" if outgoing else person,
                    recipients=rnd.choice(PEOPLE),
                    categories=rnd.choice(["", "", "Important", "Finance"]),
                    received=None if undated_received else when,
                    sent=None if undated_received else when,
                    created=when, modified=when,
                    start=when if mclass == "IPM.Appointment" else None,
                    end=when + dt.timedelta(hours=1) if mclass == "IPM.Appointment" else None,
                    recurring=mclass == "IPM.Appointment" and n % 10 == 0,
                )
                body = f"Hello,\n\nThis is about {topic}.\n\nRegards,\n{info.sender}"
                items.append((info, body, rnd.choice([0, 0, 0, 1, 2])))
            self._items[key] = items

    def connect(self) -> str:
        return self.t("demo")

    def list_tree(self) -> list[FolderNode]:
        root = FolderNode(key="store1", name=self.store, label=self.store, is_store=True)
        by_parts: dict[tuple, FolderNode] = {}
        for idx, (parts, _, _) in enumerate(DEMO_FOLDERS):
            key = f"f{idx}"
            node = FolderNode(key=key, name=parts[-1], label="/".join((self.store,) + parts),
                              count=len(self._items[key]))
            by_parts[parts] = node
            parent = by_parts.get(parts[:-1]) if len(parts) > 1 else None
            (parent.children if parent else root.children).append(node)
        return [root]

    def folder_label(self, key: str) -> str:
        return "/".join((self.store,) + DEMO_FOLDERS[int(key[1:])][0])

    def iter_folder(self, key: str, criteria: Criteria, progress: Progress,
                    cancelled: IsCancelled) -> Iterator[ItemInfo]:
        for info, body, atts in self._items[key]:
            if cancelled():
                raise Cancelled()
            if not matches_quick(info, criteria):
                continue
            if criteria.needs_item_open:
                info.attachments = atts
                if not matches_deep(info.subject, body, atts, criteria):
                    continue
            yield info

    def open_item(self, item: ItemInfo):
        for entry in self._items[item.folder_key]:
            if entry[0].entry_id == item.entry_id:
                return entry
        raise KeyError(item.entry_id)

    def attachment_count(self, handle) -> int:
        return handle[2]

    def pst_begin(self, path: str, display_name: str) -> str:
        path += ".demo.txt"
        self._copied = 0
        self._pst_log = open(path, "w", encoding="utf-8")
        self._pst_log.write("Demo mode: a real PST file is created only when the program runs with Outlook.\n")
        return path

    def pst_copy(self, handle, item: ItemInfo, include_store: bool) -> None:
        time.sleep(float(os.environ.get("OUTLOOK_ARCHIVER_DEMO_DELAY") or 0))  # to watch the progress screen
        self._pst_log.write(f"{item.folder_path}\t{item.subject}\n")
        self._copied += 1

    def pst_verify(self) -> str | None:
        return None

    def pst_end(self, keep_open: bool) -> None:
        self._pst_log.close()

    def in_protected_folder(self, item: ItemInfo) -> bool:
        return item.folder_parts[:1] == ("Deleted Items",)

    def remove_item(self, item: ItemInfo) -> None:
        entries = self._items[item.folder_key]
        self._items[item.folder_key] = [e for e in entries if e[0].entry_id != item.entry_id]

    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        info, body, _ = handle
        ext = {"msg": ".msg.demo.txt", "pdf": ".pdf.demo.txt", "mht": ".mht", "txt": ".txt"}[fmt]
        path = alloc.allocate(directory, stem, ext)
        d = info.display_date
        text = (f"Subject: {info.subject}\nFrom: {info.sender}\nTo: {info.recipients}\n"
                f"Date: {d:%Y-%m-%d %H:%M}\n\n{body}\n" if d else body)
        if fmt == "mht":
            text = "MIME-Version: 1.0\nContent-Type: text/plain; charset=utf-8\n\n" + text
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def save_attachments(self, handle, item: ItemInfo, directory: str,
                         alloc: PathAllocator) -> list[str]:
        att_dir = os.path.join(directory, self.t("attachments_dir", name=safe_name(item.subject, 50)))
        os.makedirs(att_dir, exist_ok=True)
        paths = []
        for i in range(handle[2]):
            p = alloc.allocate(att_dir, f"document {i + 1}", ".txt")
            with open(p, "w", encoding="utf-8") as f:
                f.write("demo attachment\n")
            paths.append(p)
        return paths
