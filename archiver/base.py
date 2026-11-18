"""The steps every archive run goes through, whatever the data source.

A backend (real Outlook, or the demo data) only has to know how to read
folders and save one item; the order of work, the progress messages, the
Excel list and the error report live here.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import traceback
from dataclasses import dataclass, field
from typing import Callable, Iterator

from .core import (
    INDEX_COLUMNS,
    Criteria,
    ExportOptions,
    ExportResult,
    ItemInfo,
    PathAllocator,
    file_stem,
    index_row,
    relative_dir,
    safe_name,
    unique_dir,
)

Progress = Callable[[str, int, int], None]  # message, done, total
IsCancelled = Callable[[], bool]

FORMAT_ORDER = ("msg", "mht", "pdf", "txt")


class BackendError(Exception):
    """A problem to show the user as-is (already written in plain words)."""


class Cancelled(Exception):
    pass


@dataclass
class FolderNode:
    key: str
    name: str
    count: int = 0
    children: list["FolderNode"] = field(default_factory=list)
    is_store: bool = False


class Backend:
    label = "Data source"

    # ---- to be provided by each backend -------------------------------
    def connect(self) -> str:
        raise NotImplementedError

    def list_tree(self) -> list[FolderNode]:
        raise NotImplementedError

    def folder_label(self, key: str) -> str:
        raise NotImplementedError

    def iter_folder(self, key: str, criteria: Criteria, progress: Progress,
                    cancelled: IsCancelled) -> Iterator[ItemInfo]:
        """Yield the items in one folder that pass every filter."""
        raise NotImplementedError

    def open_item(self, item: ItemInfo):
        raise NotImplementedError

    def attachment_count(self, handle) -> int:
        raise NotImplementedError

    def pst_begin(self, path: str, display_name: str) -> str:
        """Create and open the PST; return the path of the file written."""
        raise NotImplementedError

    def pst_copy(self, handle, item: ItemInfo, include_store: bool) -> None:
        raise NotImplementedError

    def pst_end(self, keep_open: bool) -> None:
        raise NotImplementedError

    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        raise NotImplementedError

    def save_attachments(self, handle, item: ItemInfo, directory: str,
                         alloc: PathAllocator) -> list[str]:
        raise NotImplementedError

    def finish_export(self) -> None:
        """Release helpers such as Word. Called once after every export."""

    # ---- shared work ----------------------------------------------------
    def scan(self, folder_keys: list[str], criteria: Criteria, progress: Progress,
             cancelled: IsCancelled) -> list[ItemInfo]:
        found: list[ItemInfo] = []
        for i, key in enumerate(folder_keys):
            if cancelled():
                raise Cancelled()
            progress(f"Looking in {self.folder_label(key)} ...", i, len(folder_keys))
            for item in self.iter_folder(key, criteria, progress, cancelled):
                found.append(item)
        progress(f"Finished looking: {len(found):,} items match.", len(folder_keys), len(folder_keys))
        found.sort(key=lambda it: (it.store_name, it.folder_parts,
                                   it.display_date or dt.datetime.min))
        return found

    def export(self, items: list[ItemInfo], opts: ExportOptions, progress: Progress,
               cancelled: IsCancelled) -> ExportResult:
        opts.validate()
        res = ExportResult()
        name = safe_name(opts.archive_name, 80, "Archive")
        base = unique_dir(opts.output_dir, name)
        os.makedirs(base)
        res.folder = base
        include_store = len({it.store_id for it in items}) > 1
        alloc = PathAllocator()
        formats = [f for f in FORMAT_ORDER if f in opts.file_formats]
        files_dir = os.path.join(base, "Files")
        rows: list[list[str]] = []
        pst_open = False
        pdf_failed = False

        try:
            if opts.make_pst:
                res.pst_path = os.path.join(base, name + ".pst")
                progress("Creating the PST file ...", 0, len(items))
                res.pst_path = self.pst_begin(res.pst_path, opts.archive_name.strip() or name)
                pst_open = True

            for n, item in enumerate(items):
                if cancelled():
                    res.cancelled = True
                    break
                if n % 5 == 0:
                    progress(f"Archiving {n + 1:,} of {len(items):,}: {item.subject[:70]}", n, len(items))
                where = f"'{item.subject[:60]}' ({item.store_name}/{item.folder_path})"
                try:
                    handle = self.open_item(item)
                    if item.attachments is None:
                        item.attachments = self.attachment_count(handle)
                except Exception as e:  # noqa: BLE001 - report and move on
                    res.errors.append(f"Could not open {where}: {_short(e)}")
                    continue

                done_something = False
                if pst_open:
                    try:
                        self.pst_copy(handle, item, include_store)
                        done_something = True
                    except Exception as e:  # noqa: BLE001
                        res.errors.append(f"Could not copy {where} into the PST: {_short(e)}")

                files: list[str] = []
                if formats or opts.save_attachments:
                    directory = os.path.join(files_dir, *relative_dir(item, opts.layout, include_store))
                    os.makedirs(directory, exist_ok=True)
                    stem = file_stem(item)
                    for fmt in formats:
                        if fmt == "pdf" and pdf_failed:
                            continue
                        try:
                            files.append(self.save_file(handle, item, fmt, directory, stem, alloc))
                        except BackendError as e:
                            if fmt == "pdf":
                                pdf_failed = True
                            res.errors.append(str(e))
                        except Exception as e:  # noqa: BLE001
                            res.errors.append(f"Could not save {where} as {fmt.upper()}: {_short(e)}")
                    if opts.save_attachments and item.attachments:
                        try:
                            files.extend(self.save_attachments(handle, item, directory, alloc))
                        except Exception as e:  # noqa: BLE001
                            res.errors.append(f"Could not save the attachments of {where}: {_short(e)}")
                res.files_written += len(files)
                if done_something or files or (opts.write_index and not formats and not pst_open):
                    res.copied[item.kind] = res.copied.get(item.kind, 0) + 1
                rows.append(index_row(item, files, base))
        finally:
            if pst_open:
                try:
                    self.pst_end(opts.keep_pst_open)
                except Exception as e:  # noqa: BLE001
                    res.errors.append(f"Could not close the PST file cleanly: {_short(e)}")
            self.finish_export()
            if opts.write_index:
                _write_index(os.path.join(base, "Archive list.csv"), rows)
            _write_readme(base, res, opts, items)

        progress("Done." if not res.cancelled else "Stopped.", len(items), len(items))
        return res


def _short(e: Exception) -> str:
    text = str(e).strip() or e.__class__.__name__
    return text if len(text) < 300 else text[:300] + "..."


def _write_index(path: str, rows: list[list[str]]) -> None:
    # utf-8-sig lets Excel show Arabic and other languages correctly.
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(INDEX_COLUMNS)
        w.writerows(rows)


def _write_readme(base: str, res: ExportResult, opts: ExportOptions, items: list[ItemInfo]) -> None:
    lines = [
        f"Outlook archive: {opts.archive_name}",
        f"Made on {dt.datetime.now():%Y-%m-%d %H:%M}",
        "",
        f"Items archived: {res.total:,} of {len(items):,} chosen"
        + (" (stopped before the end)" if res.cancelled else ""),
        "",
        "What is in this folder",
        "----------------------",
    ]
    if res.pst_path:
        lines += [
            f"* {os.path.basename(res.pst_path)}",
            "  An Outlook data file. To open it in Outlook:",
            "  File > Open & Export > Open Outlook Data File, then choose this file.",
            "  It works with no internet connection.",
        ]
    if opts.file_formats or opts.save_attachments:
        lines += ["* Files\\", "  Each item saved as its own file, sorted into folders."]
    if opts.write_index:
        lines += ["* Archive list.csv", "  A list of every archived item. Opens in Excel; you can sort and filter it."]
    if res.errors:
        lines += ["", f"Problems ({len(res.errors)})", "--------"] + [f"- {e}" for e in res.errors]
    lines += ["", "Nothing was deleted from Outlook: these are copies."]
    with open(os.path.join(base, "README.txt"), "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")


def describe_exception(e: BaseException) -> str:
    if isinstance(e, BackendError):
        return str(e)
    return f"Unexpected problem: {_short(e)}\n\n" + "".join(
        traceback.format_exception(type(e), e, e.__traceback__)[-3:])
