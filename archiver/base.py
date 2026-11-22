"""The steps every archive run goes through, whatever the data source.

A backend (real Outlook, or the demo data) only has to know how to read
folders and handle one item; the order of work, the safety checks before
removing anything, the Excel list and the README live here.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
from collections import Counter
from dataclasses import dataclass, field
from typing import Callable, Iterator

from .core import (
    INDEX_KEYS,
    Criteria,
    check_key,
    human_size,
    ExportOptions,
    ExportResult,
    ItemInfo,
    PathAllocator,
    file_stem,
    folder_size,
    index_row,
    relative_dir,
    safe_name,
    unique_dir,
)
from .i18n import Translator

Progress = Callable[[str, int, int], None]  # message, done, total
IsCancelled = Callable[[], bool]
OnArchived = Callable[[list[ItemInfo], str], None]

FORMAT_ORDER = ("eml", "msg", "mht", "pdf", "txt")


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
    size: int | None = None  # bytes in this folder (for a mailbox: the whole mailbox)
    file_path: str = ""  # mailbox only: Outlook's data file on this laptop (.ost / .pst)

    def to_dict(self) -> dict:
        file_size = None
        if self.file_path:
            try:
                file_size = os.path.getsize(self.file_path)
            except OSError:
                pass
        return {"key": self.key, "name": self.name, "label": self.label, "count": self.count,
                "is_store": self.is_store, "size": self.size, "file_size": file_size,
                "file_ext": os.path.splitext(self.file_path)[1].lower() if self.file_path else "",
                "children": [c.to_dict() for c in self.children]}


class Backend:
    label = "Data source"
    is_demo = False
    supports_pst = True  # classic Outlook can write PST files; Microsoft 365 online cannot
    kinds = ("mail", "meeting", "appointment", "contact", "task", "note", "other")
    formats = ("msg", "mht", "pdf", "txt")
    account = ""

    def can_remove(self, item: ItemInfo) -> bool:
        """Whether this data source may remove this kind of item (after archiving)."""
        return True

    def sign_out(self) -> None:
        """Forget a saved sign-in (online mode only)."""

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

    def pst_inventory(self) -> list[ItemInfo]:
        """Read back every item now in the PST (subject, dates, sender, size)."""
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

    def placeholder_size(self) -> int:
        """Demo only: the size a real PST would have had (the demo cannot write one)."""
        return 0

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
        status: dict[int, str] = {}  # item number -> check result key ("chk_ok", ...)
        files_of: dict[int, list[str]] = {}
        wanted_files: dict[int, int] = {}
        in_pst: list[tuple[int, ItemInfo]] = []
        removed_ids: set[int] = set()
        pst_open = False
        pdf_failed = False
        processed = 0

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
                processed = n + 1
                progress(t("p_archiving", i=n + 1, n=len(items), subject=item.subject[:70]), n, len(items))
                what = _what(item)
                try:
                    handle = self.open_item(item)
                    if item.attachments is None:
                        item.attachments = self.attachment_count(handle)
                except Exception as e:  # noqa: BLE001 - report and move on
                    res.errors.append(t("e_open", what=what, err=_short(e)))
                    status[n] = "chk_not_copied"
                    continue

                if pst_open:
                    try:
                        self.pst_copy(handle, item, include_store)
                        in_pst.append((n, item))
                    except Exception as e:  # noqa: BLE001
                        status[n] = "chk_not_copied"
                        res.errors.append(t("e_pst_copy", what=what, err=_short(e)))

                files: list[str] = []
                wanted = 0
                if formats or opts.save_attachments:
                    directory = os.path.join(files_dir, *relative_dir(item, opts.layout, include_store, t))
                    os.makedirs(directory, exist_ok=True)
                    stem = file_stem(item, t)
                    for fmt in formats:
                        if fmt == "pdf" and pdf_failed:
                            continue
                        wanted += 1
                        try:
                            files.append(self.save_file(handle, item, fmt, directory, stem, alloc))
                        except BackendError as e:
                            if fmt == "pdf":  # Word missing: a known gap, reported once
                                pdf_failed = True
                                wanted -= 1
                            res.errors.append(str(e))
                        except Exception as e:  # noqa: BLE001
                            res.errors.append(t("e_save", what=what, fmt=fmt.upper(), err=_short(e)))
                    if opts.save_attachments and item.attachments:
                        try:
                            files.extend(self.save_attachments(handle, item, directory, alloc))
                        except Exception as e:  # noqa: BLE001
                            wanted += 1
                            res.errors.append(t("e_atts", what=what, err=_short(e)))
                files_of[n] = files
                wanted_files[n] = wanted
                res.files_written += len(files)

            # ---- 2. check: is every chosen item really in the archive?
            progress(t("p_verifying"), len(items), len(items))
            res.check = self._check(items, processed, status, in_pst, files_of, wanted_files, pst_open, res)
            for n, key in status.items():
                if key == "chk_ok":
                    kind = items[n].kind
                    res.copied[kind] = res.copied.get(kind, 0) + 1

            # ---- 3. only if asked, and only what the check confirmed is safely archived:
            #         in the PST (classic Outlook) or as an original .eml file (online)
            if pst_open:
                candidates = in_pst
            elif "eml" in formats:
                candidates = [(n, items[n]) for n in range(processed)]
            else:
                candidates = []
            if opts.move_after and candidates and not res.cancelled:
                pst_check = res.check.get("pst") or {}
                if pst_open and (not pst_check.get("checked") or pst_check.get("found", 0) < len(in_pst)):
                    missing = len(in_pst) - pst_check.get("found", 0)
                    res.errors.insert(0, t("e_verify", detail=t("v_missing_n", n=max(missing, 0))))
                else:
                    progress(t("p_checking"), len(items), len(items))
                    safe = [(n, it) for n, it in candidates if status.get(n) == "chk_ok"]
                    removable = [(n, it) for n, it in safe
                                 if it.removal_blocker() is None and not self.in_protected_folder(it)
                                 and self.can_remove(it)]
                    res.kept = len(candidates) - len(removable)
                    for i, (n, item) in enumerate(removable):
                        if cancelled():
                            res.cancelled = True
                            break
                        progress(t("p_removing", i=i + 1, n=len(removable)), i, len(removable))
                        try:
                            self.remove_item(item)
                            removed_ids.add(n)
                            res.removed += 1
                            res.removed_size += item.size
                        except Exception as e:  # noqa: BLE001
                            res.errors.append(t("e_remove", what=_what(item), err=_short(e)))
        finally:
            if pst_open:
                try:
                    self.pst_end(opts.keep_pst_open)
                except Exception as e:  # noqa: BLE001
                    res.errors.append(t("e_pst_close", err=_short(e)))
            self.finish_export()
            progress(t("p_writing"), len(items), len(items))
            for n in range(processed, len(items)):
                status.setdefault(n, "chk_stopped")
            if opts.write_index:
                rows = [index_row(item, files_of.get(n, []), base, t, removed=n in removed_ids,
                                  status=t(status.get(n, "chk_not_checked")))
                        for n, item in enumerate(items)]
                _write_index(os.path.join(base, t("index_file")), rows, t)
            _write_readme(base, res, opts, items, t)
            res.archive_size = folder_size(base) + self.placeholder_size()
            done = [items[n] for n, key in status.items() if key == "chk_ok"]
            if on_archived and done:
                on_archived(done, base)

        progress(t("p_stopped") if res.cancelled else t("p_done"), len(items), len(items))
        return res

    def _check(self, items: list[ItemInfo], processed: int, status: dict[int, str],
               in_pst: list[tuple[int, ItemInfo]], files_of: dict[int, list[str]],
               wanted_files: dict[int, int], pst_open: bool, res: ExportResult) -> dict:
        """Compare what was chosen with what is really in the archive, item by item."""
        t = self.t
        check: dict = {"chosen": len(items), "chosen_size": sum(it.size for it in items)}

        if pst_open:
            pst = {"checked": False, "expected": len(items), "found": 0,
                   "expected_size": check["chosen_size"], "found_size": 0}
            try:
                inventory = self.pst_inventory()
                pst["checked"] = True
            except Exception as e:  # noqa: BLE001
                inventory = []
                res.errors.insert(0, t("e_check", err=_short(e)))
            if pst["checked"]:
                pst["found_size"] = sum(it.size for it in inventory)
                waiting = Counter(check_key(it) for it in inventory)
                for n, item in in_pst:
                    key = check_key(item)
                    if waiting[key] > 0:
                        waiting[key] -= 1
                        pst["found"] += 1
                    else:
                        status[n] = "chk_not_in_pst"
            else:
                for n, _item in in_pst:
                    status[n] = "chk_not_checked"
            check["pst"] = pst

        if any(wanted_files.values()) or any(files_of.values()):
            expected = found = 0
            for n, paths in files_of.items():
                good = [p for p in paths if os.path.isfile(p) and os.path.getsize(p) > 0]
                need = max(wanted_files.get(n, 0), len(paths))
                expected += need
                found += len(good)
                if len(good) < need and status.get(n) is None:
                    status[n] = "chk_file_missing"
            check["files"] = {"expected": expected, "found": found}

        for n in range(processed):
            status.setdefault(n, "chk_ok")
        for n in range(processed, len(items)):
            status[n] = "chk_stopped"
        ok_items = sum(1 for k in status.values() if k == "chk_ok")
        check["ok_items"] = ok_items
        check["ok_size"] = sum(items[n].size for n, k in status.items() if k == "chk_ok")
        check["ok"] = ok_items == len(items)
        check["missing"] = [
            {"subject": items[n].subject, "folder": f"{items[n].store_name}/{items[n].folder_path}",
             "date": items[n].display_date.strftime("%Y-%m-%d %H:%M") if items[n].display_date else "",
             "reason": t(k)}
            for n, k in sorted(status.items()) if k != "chk_ok"
        ][:500]
        check["missing_count"] = len(items) - ok_items
        return check


def _what(item: ItemInfo) -> str:
    return f"'{item.subject[:60]}' ({item.store_name}/{item.folder_path})"


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
    c = res.check
    if c:
        title = t("r_check")
        lines += ["", title, "-" * len(title), t("r_chk_items", n=c["chosen"], size=human_size(c["chosen_size"]))]
        if c.get("pst"):
            p = c["pst"]
            lines.append(t("r_chk_pst", found=p["found"], n=p["expected"],
                           size_a=human_size(p["expected_size"]), size_b=human_size(p["found_size"])))
        if c.get("files"):
            lines.append(t("r_chk_files", found=c["files"]["found"], n=c["files"]["expected"]))
        lines.append(t("r_chk_ok") if c["ok"] else t("r_chk_bad", n=c["missing_count"], col=t("col_check")))
        if c["missing"]:
            title = t("r_missing")
            lines += ["", title, "-" * len(title)]
            lines += [f"- {m['date']}  {m['subject']}  ({m['folder']}) - {m['reason']}" for m in c["missing"]]
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
