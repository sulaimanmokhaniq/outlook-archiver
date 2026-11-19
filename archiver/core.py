"""Pure logic shared by the app: item types, filters and file naming.

Nothing in here talks to Outlook, so it can be tested on any computer.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import os
import re
from dataclasses import dataclass, field

from .i18n import Translator

# ---------------------------------------------------------------- item types

ITEM_KINDS = ("mail", "meeting", "appointment", "contact", "task", "note", "other")

# Kinds that may be removed from the mailbox after archiving. Contacts and
# tasks are "living" data; calendar items are checked one by one.
REMOVABLE_KINDS = {"mail", "meeting", "note", "other", "appointment"}


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

DATE_FIELDS = ("auto", "received", "sent", "created", "modified")


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
    end: dt.datetime | None = None
    recurring: bool = False
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

    def fingerprint(self) -> str:
        """Identifies the item across runs, even if Outlook gives it a new ID."""
        d = self.display_date or self.created or self.modified
        raw = "\x1f".join([
            self.store_name.lower(), self.kind, self.subject.strip(), self.sender.strip().lower(),
            d.isoformat() if d else "", self.created.isoformat() if self.created else "",
        ])
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()

    def removal_blocker(self, now: dt.datetime | None = None) -> str | None:
        """Why this item must stay in the mailbox, or None if it may go."""
        if self.kind not in REMOVABLE_KINDS:
            return "kind"
        if self.kind == "appointment":
            now = now or dt.datetime.now()
            if self.recurring:
                return "recurring"
            if (self.end or self.start or now) >= now:
                return "future"
        return None


# ----------------------------------------------------------------- filters


def split_terms(text: str | None) -> list[str]:
    """'ahmed; sara , ' -> ['ahmed', 'sara']"""
    return [t.strip().lower() for t in re.split(r"[;,،\n]", text or "") if t.strip()]


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

    @classmethod
    def from_dict(cls, d: dict) -> "Criteria":
        kinds = {k for k in d.get("kinds", ITEM_KINDS) if k in ITEM_KINDS}
        field_ = d.get("date_field", "auto")
        return cls(
            kinds=kinds,
            date_field=field_ if field_ in DATE_FIELDS else "auto",
            date_from=parse_date(d.get("date_from", "")),
            date_to=parse_date(d.get("date_to", "")),
            sender=str(d.get("sender", "")),
            recipients=str(d.get("recipients", "")),
            subject=str(d.get("subject", "")),
            body_words=str(d.get("body_words", "")),
            category=str(d.get("category", "")),
            attachments=d.get("attachments", "any") if d.get("attachments") in ("any", "with", "without") else "any",
        )


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

    for text, value in ((c.sender, item.sender), (c.recipients, item.recipients),
                        (c.subject, item.subject), (c.category, item.categories)):
        terms = split_terms(text)
        if terms and not _contains_any(value, terms):
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


def parse_date(text) -> dt.date | None:
    """Accepts 2024-03-05, 2024/3/5, 05/03/2024 (day/month/year) or ''."""
    text = str(text or "").strip()
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


LAYOUTS = ("folder_year", "type_year", "year_month", "flat")


def relative_dir(item: ItemInfo, layout: str, include_store: bool, t: Translator | None = None) -> list[str]:
    """Sub-folders (as a list of names) where an item's files should go."""
    t = t or Translator()
    d = item.display_date
    year = str(d.year) if d else t("no_date")
    if layout == "flat":
        return []
    if layout == "type_year":
        return [t("kind_" + item.kind), year]
    if layout == "year_month":
        return [year, f"{d.month:02d}" if d else t("no_date")]
    parts = [safe_name(p, 60, "_") for p in item.folder_parts if p]
    if include_store:
        parts.insert(0, safe_name(item.store_name, 60, "Mailbox"))
    return parts + [year]


def file_stem(item: ItemInfo, t: Translator | None = None) -> str:
    t = t or Translator()
    if item.kind == "contact":
        return safe_name(item.subject, 80, t("no_name"))
    d = item.display_date
    prefix = d.strftime("%Y-%m-%d %H%M") if d else "0000-00-00"
    return f"{prefix} - {safe_name(item.subject, 80, t('no_subject'))}"


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

FILE_FORMATS = ("msg", "mht", "pdf", "txt")


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
    move_after: bool = False  # remove from the mailbox once safely in the PST
    lang: str = "en"

    @classmethod
    def from_dict(cls, d: dict) -> "ExportOptions":
        layout = d.get("layout", "folder_year")
        return cls(
            output_dir=str(d.get("output_dir", "")).strip(),
            archive_name=str(d.get("archive_name", "")).strip(),
            make_pst=bool(d.get("make_pst", True)),
            keep_pst_open=bool(d.get("keep_pst_open", True)),
            file_formats={f for f in d.get("file_formats", []) if f in FILE_FORMATS},
            save_attachments=bool(d.get("save_attachments", False)),
            layout=layout if layout in LAYOUTS else "folder_year",
            write_index=bool(d.get("write_index", True)),
            move_after=bool(d.get("move_after", False)),
            lang=d.get("lang", "en"),
        )

    def validate(self) -> None:
        t = Translator(self.lang)
        if not self.output_dir:
            raise ValueError(t("v_no_dir"))
        if not os.path.isdir(self.output_dir):
            raise ValueError(t("v_bad_dir", path=self.output_dir))
        if not safe_name(self.archive_name, 80, ""):
            raise ValueError(t("v_no_name"))
        if not (self.make_pst or self.file_formats or self.save_attachments or self.write_index):
            raise ValueError(t("v_no_output"))
        if self.move_after and not self.make_pst:
            raise ValueError(t("v_move_needs_pst"))


@dataclass
class ExportResult:
    folder: str = ""
    pst_path: str = ""
    copied: dict[str, int] = field(default_factory=dict)
    files_written: int = 0
    removed: int = 0
    kept: int = 0
    errors: list[str] = field(default_factory=list)
    cancelled: bool = False

    @property
    def total(self) -> int:
        return sum(self.copied.values())

    def to_dict(self) -> dict:
        return {
            "folder": self.folder, "pst_path": self.pst_path, "copied": self.copied,
            "total": self.total, "files_written": self.files_written, "removed": self.removed,
            "kept": self.kept, "errors": self.errors, "cancelled": self.cancelled,
        }


def summarize(items: list[ItemInfo]) -> dict:
    """How many items per type and per year."""
    by_kind: dict[str, int] = {}
    by_year: dict[str, int] = {}
    for it in items:
        by_kind[it.kind] = by_kind.get(it.kind, 0) + 1
        d = it.display_date
        y = str(d.year) if d else "-"
        by_year[y] = by_year.get(y, 0) + 1
    return {"total": len(items), "by_kind": by_kind,
            "by_year": dict(sorted(by_year.items()))}


INDEX_KEYS = ["col_type", "col_folder", "col_date", "col_from", "col_to", "col_subject",
              "col_categories", "col_attachments", "col_files", "col_removed"]


def csv_safe(value: str) -> str:
    """Stop Excel from treating text such as '=HYPERLINK(...)' as a formula."""
    value = (value or "").replace("\r", " ").replace("\n", " ")
    return "'" + value if value[:1] in ("=", "+", "-", "@") else value


def index_row(item: ItemInfo, files: list[str], base_dir: str, t: Translator | None = None,
              removed: bool = False) -> list[str]:
    t = t or Translator()
    d = item.display_date
    rel = [os.path.relpath(f, base_dir) for f in files]
    row = [
        t("kind_" + item.kind),
        f"{item.store_name}/{item.folder_path}",
        d.strftime("%Y-%m-%d %H:%M") if d else "",
        item.sender,
        item.recipients,
        item.subject,
        item.categories,
        "" if item.attachments is None else str(item.attachments),
        " | ".join(rel),
        t("yes") if removed else "",
    ]
    return [csv_safe(v) for v in row]


def preview_row(item: ItemInfo) -> dict:
    d = item.display_date
    return {"kind": item.kind, "date": d.strftime("%Y-%m-%d %H:%M") if d else "",
            "from": item.sender, "subject": item.subject,
            "folder": f"{item.store_name}/{item.folder_path}"}
