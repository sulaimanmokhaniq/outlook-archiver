"""What the program remembers between runs: settings and archive history.

Kept in %APPDATA%\\OutlookArchiver on Windows (or ~/.config/outlook-archiver).
Set OUTLOOK_ARCHIVER_HOME to use another folder (the tests do).
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sqlite3
import threading
from contextlib import contextmanager
from typing import Iterable, Iterator

from .core import ItemInfo


def app_dir() -> str:
    base = os.environ.get("OUTLOOK_ARCHIVER_HOME")
    if not base:
        if os.name == "nt" and os.environ.get("APPDATA"):
            base = os.path.join(os.environ["APPDATA"], "OutlookArchiver")
        else:
            base = os.path.join(os.path.expanduser("~"), ".config", "outlook-archiver")
    os.makedirs(base, exist_ok=True)
    return base


# ------------------------------------------------------------------ settings

DEFAULT_SETTINGS: dict = {
    "lang": "",  # "" = not chosen yet
    "kinds": ["mail", "meeting", "appointment", "contact", "task", "note", "other"],
    "period": "all",
    "year": "",
    "date_from": "",
    "date_to": "",
    "date_field": "auto",
    "sender": "",
    "recipients": "",
    "subject": "",
    "body_words": "",
    "category": "",
    "attachments": "any",
    "make_pst": True,
    "keep_pst_open": True,
    "file_formats": [],
    "save_attachments": False,
    "write_index": True,
    "layout": "folder_year",
    "output_dir": "",
    "skip_archived": True,
    "folders": None,  # list of folder labels, None = first mailbox
    "engine": "auto",  # auto | classic | online
    "client_id": "",  # Microsoft 365 online mode: the app registration's Application (client) ID
    "tenant": "organizations",
}


class Settings:
    def __init__(self, path: str | None = None) -> None:
        self.path = path or os.path.join(app_dir(), "settings.json")
        self._lock = threading.Lock()

    def load(self) -> dict:
        data = dict(DEFAULT_SETTINGS)
        if not data["output_dir"]:
            data["output_dir"] = os.path.join(os.path.expanduser("~"), "Outlook Archives")
        try:
            with open(self.path, encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                data.update({k: v for k, v in saved.items() if k in DEFAULT_SETTINGS})
        except (OSError, ValueError):
            pass
        return data

    def save(self, changes: dict) -> dict:
        with self._lock:
            data = self.load()
            data.update({k: v for k, v in changes.items() if k in DEFAULT_SETTINGS})
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1)
            os.replace(tmp, self.path)
            return data


# ------------------------------------------------------------------- history


class History:
    """Remembers which items were archived, and where."""

    def __init__(self, path: str | None = None) -> None:
        self.path = path or os.path.join(app_dir(), "history.sqlite3")
        with self._db() as db:
            db.execute("CREATE TABLE IF NOT EXISTS archived ("
                       " fp TEXT PRIMARY KEY, archive_path TEXT NOT NULL, archived_at TEXT NOT NULL)")

    @contextmanager
    def _db(self) -> Iterator[sqlite3.Connection]:
        con = sqlite3.connect(self.path, timeout=10)
        try:
            with con:  # commits, or rolls back on error
                yield con
        finally:
            con.close()

    def still_archived(self) -> set[str]:
        """Fingerprints whose archive folder still exists on this laptop."""
        with self._db() as db:
            rows = db.execute("SELECT fp, archive_path FROM archived").fetchall()
        exists: dict[str, bool] = {}
        out: set[str] = set()
        for fp, path in rows:
            if path not in exists:
                exists[path] = os.path.isdir(path)
            if exists[path]:
                out.add(fp)
        return out

    def filter_new(self, items: list[ItemInfo]) -> tuple[list[ItemInfo], int]:
        done = self.still_archived()
        if not done:
            return items, 0
        new = [it for it in items if it.fingerprint() not in done]
        return new, len(items) - len(new)

    def record(self, items: Iterable[ItemInfo], archive_path: str) -> None:
        now = dt.datetime.now().isoformat(timespec="seconds")
        rows = [(it.fingerprint(), archive_path, now) for it in items]
        if rows:
            with self._db() as db:
                db.executemany("INSERT OR REPLACE INTO archived VALUES (?, ?, ?)", rows)

    def count(self) -> int:
        return len(self.still_archived())

    def forget(self) -> None:
        with self._db() as db:
            db.execute("DELETE FROM archived")
