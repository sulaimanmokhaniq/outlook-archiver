"""Made-up mailbox so the app can be tried (and tested) without Outlook.

Start the app with --demo. PST files and .msg/PDF files need real Outlook,
so in demo mode those are written as small text placeholders.
"""
from __future__ import annotations

import datetime as dt
import random
from typing import Iterator

from .base import Backend, Cancelled, FolderNode, IsCancelled, Progress
from .core import Criteria, ItemInfo, PathAllocator, matches_deep, matches_quick, safe_name

PEOPLE = ["Ahmed Ali", "Sara Hassan", "Omar Khalid", "Mona Youssef", "Finance Team", "IT Support"]
TOPICS = ["Budget 2024", "Quarterly report", "Board meeting", "Contract renewal",
          "Travel plan", "Project Alpha", "Invoice", "Weekly update", "Hiring"]
DEMO_FOLDERS = [
    (("Inbox",), "IPM.Note", 700),
    (("Inbox", "Projects"), "IPM.Note", 150),
    (("Sent Items",), "IPM.Note", 400),
    (("Drafts",), "IPM.Note", 25),
    (("Calendar",), "IPM.Appointment", 300),
    (("Contacts",), "IPM.Contact", 60),
    (("Tasks",), "IPM.Task", 40),
]


class DemoBackend(Backend):
    label = "Demo data"

    def __init__(self) -> None:
        self.store = "Demo Mailbox"
        self._items: dict[str, list[tuple[ItemInfo, str]]] = {}
        rnd = random.Random(7)
        start = dt.datetime(2017, 1, 1)
        for idx, (parts, mclass, count) in enumerate(DEMO_FOLDERS):
            key = f"f{idx}"
            items = []
            for n in range(count):
                when = start + dt.timedelta(minutes=rnd.randint(0, 60 * 24 * 365 * 9))
                topic = rnd.choice(TOPICS)
                mc = mclass
                if parts == ("Inbox",) and n % 15 == 0:
                    mc = "IPM.Schedule.Meeting.Request"
                person = rnd.choice(PEOPLE)
                subject = person if mclass == "IPM.Contact" else f"{topic} #{n}"
                info = ItemInfo(
                    entry_id=f"{key}-{n}", store_id="demo", folder_key=key,
                    folder_parts=parts, store_name=self.store, message_class=mc,
                    subject=subject,
                    sender="Me" if parts[0] in ("Sent Items", "Drafts") else person,
                    recipients=rnd.choice(PEOPLE),
                    categories=rnd.choice(["", "", "Important", "Finance"]),
                    received=None if parts[0] in ("Drafts", "Calendar", "Contacts", "Tasks") else when,
                    sent=None if parts[0] in ("Drafts", "Contacts", "Tasks", "Calendar") else when,
                    created=when, modified=when,
                    start=when if mclass == "IPM.Appointment" else None,
                    attachments=None,
                )
                body = f"Hello,\n\nThis is about {topic.lower()}.\n\nRegards,\n{info.sender}"
                info_atts = rnd.choice([0, 0, 0, 1, 2])
                items.append((info, body, info_atts))
            self._items[key] = items  # type: ignore[assignment]

    def connect(self) -> str:
        return "DEMO MODE - showing made-up data, not your real Outlook."

    def list_tree(self) -> list[FolderNode]:
        root = FolderNode(key="store1", name=self.store, is_store=True)
        by_parts = {}
        for idx, (parts, _, _) in enumerate(DEMO_FOLDERS):
            node = FolderNode(key=f"f{idx}", name=parts[-1], count=len(self._items[f"f{idx}"]))
            by_parts[parts] = node
            parent = by_parts.get(parts[:-1]) if len(parts) > 1 else None
            (parent.children if parent else root.children).append(node)
        return [root]

    def folder_label(self, key: str) -> str:
        idx = int(key[1:])
        return "/".join((self.store,) + DEMO_FOLDERS[idx][0])

    def iter_folder(self, key: str, criteria: Criteria, progress: Progress,
                    cancelled: IsCancelled) -> Iterator[ItemInfo]:
        for info, body, atts in self._items[key]:  # type: ignore[misc]
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
        for info, body, atts in self._items[item.folder_key]:  # type: ignore[misc]
            if info.entry_id == item.entry_id:
                return (info, body, atts)
        raise KeyError(item.entry_id)

    def attachment_count(self, handle) -> int:
        return handle[2]

    def pst_begin(self, path: str, display_name: str) -> str:
        path += ".demo.txt"
        self._pst_log = open(path, "w", encoding="utf-8")
        self._pst_log.write("Demo mode: a real PST file is created only when the tool runs with Outlook.\n")
        return path

    def pst_copy(self, handle, item: ItemInfo, include_store: bool) -> None:
        self._pst_log.write(f"{item.folder_path}\t{item.subject}\n")

    def pst_end(self, keep_open: bool) -> None:
        self._pst_log.close()

    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        info, body, _ = handle
        ext = {"msg": ".msg.demo.txt", "pdf": ".pdf.demo.txt", "mht": ".mht", "txt": ".txt"}[fmt]
        path = alloc.allocate(directory, stem, ext)
        d = info.display_date
        text = (f"Subject: {info.subject}\nFrom: {info.sender}\nTo: {info.recipients}\n"
                f"Date: {d:%Y-%m-%d %H:%M}\n\n{body}\n" if d else body)
        if fmt == "mht":
            text = ("MIME-Version: 1.0\nContent-Type: text/plain; charset=utf-8\n\n" + text)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def save_attachments(self, handle, item: ItemInfo, directory: str,
                         alloc: PathAllocator) -> list[str]:
        import os
        att_dir = os.path.join(directory, safe_name(item.subject, 50) + " - attachments")
        os.makedirs(att_dir, exist_ok=True)
        paths = []
        for i in range(handle[2]):
            p = alloc.allocate(att_dir, f"document {i + 1}", ".txt")
            with open(p, "w", encoding="utf-8") as f:
                f.write("demo attachment\n")
            paths.append(p)
        return paths
