"""Pure logic shared by the app: item types, filters and file naming.

Nothing in here talks to Outlook, so it can be tested on any computer.
"""
from __future__ import annotations

import datetime as dt
import os
import re
from dataclasses import dataclass, field

# ---------------------------------------------------------------- item types

ITEM_KINDS: dict[str, str] = {
    "mail": "Emails (incl. drafts)",
    "meeting": "Meeting invitations & replies",
    "appointment": "Calendar (appointments & meetings)",
    "contact": "Contacts",
    "task": "Tasks",
    "note": "Notes",
    "other": "Other (reports, posts, ...)",
}

KIND_FOLDER_NAMES: dict[str, str] = {
    "mail": "Emails",
    "meeting": "Meeting invitations",
    "appointment": "Calendar",
    "contact": "Contacts",
    "task": "Tasks",
    "note": "Notes",
    "other": "Other",
}


def classify(message_class: str | None) -> str:
    """Map an Outlook MessageClass (e.g. 'IPM.Note') to one of ITEM_KINDS."""
    mc = (message_class or "").upper()
    if mc.startswith("IPM.SCHEDULE.MEETING"):
        return "meeting"
    if mc.startswith("IPM.APPOINTMENT"):
        return "appointment"
    if mc.startswith("IPM.CONTACT") or mc.startswith("IPM.DISTLIST"):
        return "contact"
    if mc.startswith("IPM.TASK"):  # tasks and task requests
        return "task"
    if mc.startswith("IPM.STICKYNOTE"):
        return "note"
    if mc == "IPM" or mc.startswith("IPM.NOTE"):
        return "mail"
    return "other"


# ------------------------------------------------------------------- dates

DATE_FIELDS: dict[str, str] = {
    "auto": "Best date for each item",
    "received": "Received date",
    "sent": "Sent date",
    "created": "Created date",
    "modified": "Last modified date",
}


def clean_date(value) -> dt.datetime | None:
    """Turn whatever Outlook returns into a plain datetime, or None.

    Outlook uses 4501-01-01 to mean "no date"; pywin32 adds a time zone tag
    that does not match the clock time it holds, so the tag is dropped.
    """
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        d = value
    elif isinstance(value, dt.date):
        d = dt.datetime(value.year, value.month, value.day)
    else:
        return None
    if d.year >= 4500 or d.year < 1950:
        return None
    return dt.datetime(d.year, d.month, d.day, d.hour, d.minute, d.second)


@dataclass
class ItemInfo:
    """What the app knows about one Outlook item before opening it."""

    entry_id: str
    store_id: str
    folder_key: str
    folder_parts: tuple[str, ...]  # e.g. ("Inbox", "Projects"); store name not included
    store_name: str
    message_class: str
    subject: str = ""
    sender: str = ""
    recipients: str = ""
    categories: str = ""
    received: dt.datetime | None = None
    sent: dt.datetime | None = None
    created: dt.datetime | None = None
    modified: dt.datetime | None = None
    start: dt.datetime | None = None
    attachments: int | None = None  # filled in only when the item is opened

    @property
    def kind(self) -> str:
        return classify(self.message_class)

    @property
    def folder_path(self) -> str:
        return "/".join(self.folder_parts)

    def date_for(self, date_field: str) -> dt.datetime | None:
        if date_field == "received":
            return self.received
        if date_field == "sent":
            return self.sent
        if date_field == "created":
            return self.created
        if date_field == "modified":
            return self.modified
        # auto: the date a person would think of for this kind of item
        if self.kind == "appointment":
            chain = (self.start, self.created, self.modified)
        elif self.kind in ("contact", "note", "task"):
            chain = (self.created, self.modified)
        else:
            chain = (self.received, self.sent, self.modified, self.created)
        return next((d for d in chain if d is not None), None)

    @property
    def display_date(self) -> dt.datetime | None:
        return self.date_for("auto")


# ----------------------------------------------------------------- filters


def split_terms(text: str | None) -> list[str]:
    """'ahmed; sara , ' -> ['ahmed', 'sara']"""
    return [t.strip().lower() for t in re.split(r"[;,\n]", text or "") if t.strip()]


def _contains_any(haystack: str, terms: list[str]) -> bool:
    h = (haystack or "").lower()
    return any(t in h for t in terms)


@dataclass
class Criteria:
    kinds: set[str] = field(default_factory=lambda: set(ITEM_KINDS))
    date_field: str = "auto"
    date_from: dt.date | None = None  # inclusive
    date_to: dt.date | None = None  # inclusive
    include_undated: bool = False
    sender: str = ""
    recipients: str = ""
    subject: str = ""
    body_words: str = ""
    category: str = ""
    attachments: str = "any"  # any | with | without

    @property
    def needs_item_open(self) -> bool:
        """True if some filters need the full item (slower)."""
        return bool(split_terms(self.body_words)) or self.attachments != "any"

    def describe(self) -> str:
        parts = []
        if set(self.kinds) != set(ITEM_KINDS):
            parts.append("types: " + ", ".join(ITEM_KINDS[k] for k in ITEM_KINDS if k in self.kinds))
        if self.date_from or self.date_to:
            a = self.date_from.isoformat() if self.date_from else "the beginning"
            b = self.date_to.isoformat() if self.date_to else "today"
            parts.append(f"{DATE_FIELDS[self.date_field].lower()} from {a} to {b}")
        for label, value in (
            ("from", self.sender),
            ("to/cc", self.recipients),
            ("subject", self.subject),
            ("words", self.body_words),
            ("category", self.category),
        ):
            if value.strip():
                parts.append(f"{label} contains '{value.strip()}'")
        if self.attachments != "any":
            parts.append(f"{self.attachments} attachments")
        return "; ".join(parts) if parts else "everything in the chosen folders"


def matches_quick(item: ItemInfo, c: Criteria) -> bool:
    """Every filter that can be checked without opening the item."""
    if item.kind not in c.kinds:
        return False

    if c.date_from or c.date_to:
        d = item.date_for(c.date_field)
        if d is None:
            if not c.include_undated:
                return False
        else:
            if c.date_from and d.date() < c.date_from:
                return False
            if c.date_to and d.date() > c.date_to:
                return False

    terms = split_terms(c.sender)
    if terms and not _contains_any(item.sender, terms):
        return False
    terms = split_terms(c.recipients)
    if terms and not _contains_any(item.recipients, terms):
        return False
    terms = split_terms(c.subject)
    if terms and not _contains_any(item.subject, terms):
        return False
    terms = split_terms(c.category)
    if terms and not _contains_any(item.categories, terms):
        return False
    return True


def matches_deep(subject: str, body: str, attachment_count: int, c: Criteria) -> bool:
    """Filters that need the opened item (body text, attachments)."""
    terms = split_terms(c.body_words)
    if terms and not _contains_any((subject or "") + "\n" + (body or ""), terms):
        return False
    if c.attachments == "with" and attachment_count <= 0:
        return False
    if c.attachments == "without" and attachment_count > 0:
        return False
    return True


def parse_date(text: str) -> dt.date | None:
    """Accepts 2024-03-05, 2024/3/5, 05/03/2024 (day/month/year) or ''."""
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%Y.%m.%d", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"'{text}' is not a date. Use YYYY-MM-DD, for example 2024-03-31.")


# ------------------------------------------------------------- file naming

_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}

# Windows refuses paths longer than 260 characters in many programs.
MAX_PATH = 250


def safe_name(text: str | None, max_len: int = 80, fallback: str = "(no subject)") -> str:
    s = _BAD_CHARS.sub("_", text or "")
    s = re.sub(r"\s+", " ", s).strip().rstrip(". ")
    if len(s) > max_len:
        s = s[:max_len].rstrip(". ")
    if not s:
        s = fallback
    if s.split(".")[0].upper() in _RESERVED:
        s = "_" + s
    return s


LAYOUTS: dict[str, str] = {
    "folder_year": "Outlook folder, then year",
    "type_year": "Item type, then year",
    "year_month": "Year, then month",
    "flat": "All in one folder",
}


def relative_dir(item: ItemInfo, layout: str, include_store: bool) -> list[str]:
    """Sub-folders (as a list of names) where an item's files should go."""
    d = item.display_date
    year = str(d.year) if d else "No date"
    if layout == "flat":
        return []
    if layout == "type_year":
        return [KIND_FOLDER_NAMES[item.kind], year]
    if layout == "year_month":
        return [year, f"{d.month:02d}" if d else "No date"]
    parts = [safe_name(p, 60, "_") for p in item.folder_parts if p]
    if include_store:
        parts.insert(0, safe_name(item.store_name, 60, "Mailbox"))
    return parts + [year]


def file_stem(item: ItemInfo) -> str:
    d = item.display_date
    prefix = d.strftime("%Y-%m-%d %H%M") if d else "0000-00-00"
    if item.kind == "contact":
        return safe_name(item.subject, 80, "(no name)")
    return f"{prefix} - {safe_name(item.subject)}"


class PathAllocator:
    """Hands out file paths that are unique and not too long for Windows."""

    def __init__(self) -> None:
        self._used: set[str] = set()

    def allocate(self, directory: str, stem: str, ext: str) -> str:
        room = MAX_PATH - len(directory) - len(ext) - 8
        if room < 12:
            raise ValueError(f"Folder path is too long for Windows: {directory}")
        stem = stem[:room].rstrip(". ")
        n = 1
        while True:
            name = stem + (f" ({n})" if n > 1 else "") + ext
            path = os.path.join(directory, name)
            key = path.lower()
            if key not in self._used and not os.path.exists(path):
                self._used.add(key)
                return path
            n += 1


def unique_dir(parent: str, name: str) -> str:
    """parent/name, or parent/name (2) ... if that already exists."""
    base = os.path.join(parent, name)
    path, n = base, 2
    while os.path.exists(path):
        path = f"{base} ({n})"
        n += 1
    return path


# ------------------------------------------------------- export settings


FILE_FORMATS: dict[str, str] = {
    "msg": "Outlook files (.msg)",
    "mht": "Web pages (.mht)",
    "pdf": "PDF files (needs Word)",
    "txt": "Plain text (.txt)",
}


@dataclass
class ExportOptions:
    output_dir: str
    archive_name: str
    make_pst: bool = True
    keep_pst_open: bool = True
    file_formats: set[str] = field(default_factory=set)
    save_attachments: bool = False
    layout: str = "folder_year"
    write_index: bool = True

    def validate(self) -> None:
        if not self.output_dir:
            raise ValueError("Choose a folder to save the archive in.")
        if not os.path.isdir(self.output_dir):
            raise ValueError(f"This folder does not exist: {self.output_dir}")
        if not safe_name(self.archive_name, 80, ""):
            raise ValueError("Give the archive a name.")
        if not (self.make_pst or self.file_formats or self.save_attachments or self.write_index):
            raise ValueError("Choose at least one way to save (PST, files or the Excel list).")


@dataclass
class ExportResult:
    folder: str = ""
    pst_path: str = ""
    copied: dict[str, int] = field(default_factory=dict)
    files_written: int = 0
    errors: list[str] = field(default_factory=list)
    cancelled: bool = False

    @property
    def total(self) -> int:
        return sum(self.copied.values())


def summarize(items: list[ItemInfo]) -> str:
    """Short text: how many items per type and per year."""
    if not items:
        return "No items match these choices."
    by_kind: dict[str, int] = {}
    by_year: dict[str, int] = {}
    for it in items:
        by_kind[it.kind] = by_kind.get(it.kind, 0) + 1
        d = it.display_date
        y = str(d.year) if d else "No date"
        by_year[y] = by_year.get(y, 0) + 1
    lines = [f"{len(items):,} items match:"]
    for k in ITEM_KINDS:
        if k in by_kind:
            lines.append(f"   {ITEM_KINDS[k]}: {by_kind[k]:,}")
    lines.append("   By year: " + ", ".join(f"{y}: {n:,}" for y, n in sorted(by_year.items())))
    return "\n".join(lines)


INDEX_COLUMNS = [
    "Type",
    "Outlook folder",
    "Date",
    "From",
    "To / CC",
    "Subject",
    "Categories",
    "Attachments",
    "Saved files",
]


def csv_safe(value: str) -> str:
    """Stop Excel from treating text such as '=HYPERLINK(...)' as a formula."""
    value = (value or "").replace("\r", " ").replace("\n", " ")
    return "'" + value if value[:1] in ("=", "+", "-", "@") else value


def index_row(item: ItemInfo, files: list[str], base_dir: str) -> list[str]:
    d = item.display_date
    rel = [os.path.relpath(f, base_dir) for f in files]
    row = [
        KIND_FOLDER_NAMES[item.kind],
        f"{item.store_name}/{item.folder_path}",
        d.strftime("%Y-%m-%d %H:%M") if d else "",
        item.sender,
        item.recipients,
        item.subject,
        item.categories,
        "" if item.attachments is None else str(item.attachments),
        " | ".join(rel),
    ]
    return [csv_safe(v) for v in row]
