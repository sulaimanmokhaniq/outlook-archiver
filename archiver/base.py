"""The steps every archive run goes through, whatever the data source.

A backend (real Outlook, or the demo data) only has to know how to read
folders and handle one item; the order of work, the safety checks before
removing anything, the Excel list and the README live here.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
from dataclasses import dataclass, field
from typing import Callable, Iterator

from .core import (
    INDEX_KEYS,
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
from .i18n import Translator

Progress = Callable[[str, int, int], None]  # message, done, total
IsCancelled = Callable[[], bool]
OnArchived = Callable[[list[ItemInfo], str], None]

FORMAT_ORDER = ("msg", "mht", "pdf", "txt")


class BackendError(Exception):
    """A problem to show the user as-is (already written in plain words)."""


class Cancelled(Exception):
    pass


@dataclass
class FolderNode:
    key: str
    name: str
    label: str = ""  # "Mailbox/Inbox/Projects" - stable between runs
    count: int = 0
    children: list["FolderNode"] = field(default_factory=list)
    is_store: bool = False

    def to_dict(self) -> dict:
        return {"key": self.key, "name": self.name, "label": self.label, "count": self.count,
                "is_store": self.is_store, "children": [c.to_dict() for c in self.children]}


class Backend:
    label = "Data source"
    is_demo = False

    def __init__(self) -> None:
        self.t = Translator()

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

    def pst_verify(self) -> str | None:
        """Check the PST holds everything copied. Returns a problem, or None."""
        raise NotImplementedError

    def pst_end(self, keep_open: bool) -> None:
        raise NotImplementedError

    def save_file(self, handle, item: ItemInfo, fmt: str, directory: str,
                  stem: str, alloc: PathAllocator) -> str:
        raise NotImplementedError

    def save_attachments(self, handle, item: ItemInfo, directory: str,
                         alloc: PathAllocator) -> list[str]:
        raise NotImplementedError

    def in_protected_folder(self, item: ItemInfo) -> bool:
        """True for folders we never remove from (e.g. Deleted Items)."""
        return False

    def remove_item(self, item: ItemInfo) -> None:
        """Move the original to Deleted Items."""
        raise NotImplementedError

    def finish_export(self) -> None:
        """Release helpers such as Word. Called once after every export."""

    # ---- shared work ----------------------------------------------------
    def scan(self, folder_keys: list[str], criteria: Criteria, progress: Progress,
             cancelled: IsCancelled) -> list[ItemInfo]:
        t = self.t
        found: list[ItemInfo] = []
        for i, key in enumerate(folder_keys):
            if cancelled():
                raise Cancelled()
            progress(t("p_looking", folder=self.folder_label(key)), i, len(folder_keys))
            for item in self.iter_folder(key, criteria, progress, cancelled):
                found.append(item)
        progress(t("p_found", n=len(found)), len(folder_keys), len(folder_keys))
        found.sort(key=lambda it: (it.store_name, it.folder_parts,
                                   it.display_date or dt.datetime.min))
        return found

    def export(self, items: list[ItemInfo], opts: ExportOptions, progress: Progress,
               cancelled: IsCancelled, on_archived: OnArchived | None = None) -> ExportResult:
        t = self.t = Translator(opts.lang)
        opts.validate()
        res = ExportResult()
        name = safe_name(opts.archive_name, 80, "Archive")
        base = unique_dir(opts.output_dir, name)
        os.makedirs(base)
        res.folder = base
        include_store = len({it.store_id for it in items}) > 1
        alloc = PathAllocator()
        formats = [f for f in FORMAT_ORDER if f in opts.file_formats]
        files_dir = os.path.join(base, t("files_dir"))
        rows: dict[int, list[str]] = {}
        in_pst: list[tuple[int, ItemInfo]] = []
        archived: list[ItemInfo] = []
        removed_ids: set[int] = set()
        pst_open = False
        pdf_failed = False

        try:
            if opts.make_pst:
                progress(t("p_pst"), 0, len(items))
                res.pst_path = self.pst_begin(os.path.join(base, name + ".pst"),
                                              opts.archive_name.strip() or name)
                pst_open = True

            # ---- 1. copy everything
            for n, item in enumerate(items):
                if cancelled():
                    res.cancelled = True
                    break
                progress(t("p_archiving", i=n + 1, n=len(items), subject=item.subject[:70]), n, len(items))
                what = f"'{item.subject[:60]}' ({item.store_name}/{item.folder_path})"
                try:
                    handle = self.open_item(item)
                    if item.attachments is None:
                        item.attachments = self.attachment_count(handle)
                except Exception as e:  # noqa: BLE001 - report and move on
                    res.errors.append(t("e_open", what=what, err=_short(e)))
                    continue

                ok = True
                if pst_open:
                    try:
                        self.pst_copy(handle, item, include_store)
                        in_pst.append((n, item))
                    except Exception as e:  # noqa: BLE001
                        ok = False
                        res.errors.append(t("e_pst_copy", what=what, err=_short(e)))

                files: list[str] = []
                if formats or opts.save_attachments:
                    directory = os.path.join(files_dir, *relative_dir(item, opts.layout, include_store, t))
                    os.makedirs(directory, exist_ok=True)
                    stem = file_stem(item, t)
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
                            ok = False
                            res.errors.append(t("e_save", what=what, fmt=fmt.upper(), err=_short(e)))
                    if opts.save_attachments and item.attachments:
                        try:
                            files.extend(self.save_attachments(handle, item, directory, alloc))
                        except Exception as e:  # noqa: BLE001
                            ok = False
                            res.errors.append(t("e_atts", what=what, err=_short(e)))
                res.files_written += len(files)
                if ok:
                    res.copied[item.kind] = res.copied.get(item.kind, 0) + 1
                    archived.append(item)
                rows[n] = index_row(item, files, base, t)

            # ---- 2. only if asked, and only after the PST checks out
            if opts.move_after and pst_open and in_pst and not res.cancelled:
                progress(t("p_checking"), len(items), len(items))
                problem = self.pst_verify()
                if problem:
                    res.errors.insert(0, t("e_verify", detail=problem))
                else:
                    removable = [(n, it) for n, it in in_pst
                                 if it.removal_blocker() is None and not self.in_protected_folder(it)]
                    res.kept = len(in_pst) - len(removable)
                    for i, (n, item) in enumerate(removable):
                        if cancelled():
                            res.cancelled = True
                            break
                        progress(t("p_removing", i=i + 1, n=len(removable)), i, len(removable))
                        try:
                            self.remove_item(item)
                            removed_ids.add(n)
                            res.removed += 1
                        except Exception as e:  # noqa: BLE001
                            what = f"'{item.subject[:60]}' ({item.store_name}/{item.folder_path})"
                            res.errors.append(t("e_remove", what=what, err=_short(e)))
        finally:
            if pst_open:
                try:
                    self.pst_end(opts.keep_pst_open)
                except Exception as e:  # noqa: BLE001
                    res.errors.append(t("e_pst_close", err=_short(e)))
            self.finish_export()
            progress(t("p_writing"), len(items), len(items))
            if opts.write_index:
                for n in removed_ids:
                    rows[n][-1] = t("yes")
                _write_index(os.path.join(base, t("index_file")), [rows[k] for k in sorted(rows)], t)
            _write_readme(base, res, opts, items, t)
            if on_archived and archived:
                on_archived(archived, base)

        progress(t("p_stopped") if res.cancelled else t("p_done"), len(items), len(items))
        return res


def _short(e: Exception) -> str:
    text = str(e).strip() or e.__class__.__name__
    return text if len(text) < 300 else text[:300] + "..."


def _write_index(path: str, rows: list[list[str]], t: Translator) -> None:
    # utf-8-sig lets Excel show Arabic and other languages correctly.
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([t(k) for k in INDEX_KEYS])
        w.writerows(rows)


def _write_readme(base: str, res: ExportResult, opts: ExportOptions, items: list[ItemInfo],
                  t: Translator) -> None:
    head = t("r_inside")
    lines = [
        t("r_title", name=opts.archive_name),
        t("r_made", when=f"{dt.datetime.now():%Y-%m-%d %H:%M}"),
        "",
        t("r_count", done=res.total, total=len(items)) + (t("r_stopped") if res.cancelled else ""),
        "",
        head,
        "-" * len(head),
    ]
    if res.pst_path:
        lines.append(t("r_pst", file=os.path.basename(res.pst_path)))
    if opts.file_formats or opts.save_attachments:
        lines.append(t("r_files", dir=t("files_dir")))
    if opts.write_index:
        lines.append(t("r_index", file=t("index_file")))
    lines.append("")
    if res.removed:
        lines.append(t("r_moved", n=res.removed))
    else:
        lines.append(t("r_copies"))
    if res.kept:
        lines.append(t("r_kept", n=res.kept))
    if res.errors:
        title = t("r_problems", n=len(res.errors))
        lines += ["", title, "-" * len(title)] + [f"- {e}" for e in res.errors]
    with open(os.path.join(base, t("readme_file")), "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")
