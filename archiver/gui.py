"""The window: pick folders, item types, dates and filters, then archive."""
from __future__ import annotations

import datetime as dt
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from .base import Backend, BackendError, Cancelled, FolderNode, describe_exception
from .core import (
    DATE_FIELDS,
    FILE_FORMATS,
    ITEM_KINDS,
    LAYOUTS,
    Criteria,
    ExportOptions,
    ExportResult,
    ItemInfo,
    parse_date,
    summarize,
)

APP_TITLE = "Outlook Archiver"
BLUE, GREY, WHITE = "#1f6feb", "#6b6b6b", "#ffffff"


def _box_image(state: str) -> tk.PhotoImage:
    """A 16x16 tick box drawn in code, so it looks the same on every PC."""
    img = tk.PhotoImage(width=16, height=16)
    edge = BLUE if state != "off" else GREY
    img.put(edge, to=(1, 1, 15, 15))
    img.put(BLUE if state == "on" else WHITE, to=(2, 2, 14, 14))
    if state == "on":
        for x, y in [(4, 8), (5, 9), (6, 10), (7, 9), (8, 8), (9, 7), (10, 6), (11, 5)]:
            img.put(WHITE, to=(x, y, x + 2, y + 2))
    elif state == "some":
        img.put(BLUE, to=(5, 5, 11, 11))
    return img


class App:
    def __init__(self, root: tk.Tk, make_backend: Callable[[], Backend]) -> None:
        self.root = root
        self.make_backend = make_backend
        self.backend: Backend | None = None
        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.jobs: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.pending = 0
        self.nodes: dict[str, FolderNode] = {}
        self.parents: dict[str, str] = {}
        self.checked: set[str] = set()
        self.last_folder = ""

        root.title(APP_TITLE)
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"{min(1280, sw - 40)}x{min(820, sh - 90)}+10+10")
        root.minsize(900, 560)
        self.boxes = {s: _box_image(s) for s in ("on", "off", "some")}
        self._build()
        threading.Thread(target=self._worker, daemon=True).start()
        self._run(self._job_connect)
        root.after(100, self._poll)

    # ================================================================ layout
    def _build(self) -> None:
        r = self.root
        r.columnconfigure(0, weight=1)
        r.rowconfigure(1, weight=1)

        top = ttk.Frame(r, padding=(12, 10, 12, 0))
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text=APP_TITLE, font=("Segoe UI", 15, "bold")).pack(side="left")
        self.status = ttk.Label(top, text="Connecting to Outlook ...", foreground="#555")
        self.status.pack(side="left", padx=16)

        body = ttk.PanedWindow(r, orient="horizontal")
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)

        # ---- left: folders
        left = ttk.LabelFrame(body, text=" 1. Choose folders ", padding=8)
        body.add(left, weight=1)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        ttk.Label(left, text="Click a folder to tick or untick it.\nIts sub-folders follow.",
                  foreground="#555").grid(row=0, column=0, columnspan=2, sticky="w")
        self.tree = ttk.Treeview(left, show="tree", selectmode="none")
        self.tree.column("#0", width=250)
        self.tree.grid(row=1, column=0, sticky="nsew", pady=4)
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=4)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<space>", self._on_tree_space)
        btns = ttk.Frame(left)
        btns.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Button(btns, text="Tick all", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(btns, text="Untick all", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        ttk.Button(btns, text="Reload folders", command=lambda: self._run(self._job_connect)).pack(side="left")

        r.after(150, lambda: body.sashpos(0, 290))

        # ---- right: choices, in two columns; scrolls on small screens
        right = self._scrollable(body)
        right.columnconfigure(0, weight=1)
        right.columnconfigure(1, weight=1)
        col_a = ttk.Frame(right)
        col_a.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        col_a.columnconfigure(0, weight=1)
        col_b = ttk.Frame(right)
        col_b.grid(row=0, column=1, sticky="nsew")
        col_b.columnconfigure(0, weight=1)

        # 2. types
        types = ttk.LabelFrame(col_a, text=" 2. What to include ", padding=8)
        types.grid(row=0, column=0, sticky="ew")
        self.kind_vars: dict[str, tk.BooleanVar] = {}
        for i, (k, label) in enumerate(ITEM_KINDS.items()):
            v = tk.BooleanVar(value=True)
            self.kind_vars[k] = v
            ttk.Checkbutton(types, text=label, variable=v).grid(row=i, column=0, sticky="w")

        # 3. dates
        dates = ttk.LabelFrame(col_a, text=" 3. Time period ", padding=8)
        dates.grid(row=1, column=0, sticky="ew", pady=8)
        ttk.Label(dates, text="Quick pick:").grid(row=0, column=0, sticky="w")
        self.quick = ttk.Combobox(dates, state="readonly", width=34, values=self._quick_options())
        self.quick.set("Choose a year or period ...")
        self.quick.grid(row=0, column=1, columnspan=3, sticky="w", padx=6)
        self.quick.bind("<<ComboboxSelected>>", self._on_quick)
        ttk.Label(dates, text="From:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.date_from = ttk.Entry(dates, width=14)
        self.date_from.grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(dates, text="To:").grid(row=1, column=2, sticky="e", pady=(6, 0))
        self.date_to = ttk.Entry(dates, width=14)
        self.date_to.grid(row=1, column=3, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(dates, text="Dates as YYYY-MM-DD, e.g. 2021-12-31. Leave empty for no limit.",
                  foreground="#555").grid(row=2, column=0, columnspan=5, sticky="w")
        ttk.Label(dates, text="Use the:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.date_field = ttk.Combobox(dates, state="readonly", width=34, values=list(DATE_FIELDS.values()))
        self.date_field.current(0)
        self.date_field.grid(row=3, column=1, columnspan=3, sticky="w", padx=6, pady=(6, 0))

        # 4. filters
        flt = ttk.LabelFrame(col_a, text=" 4. Extra filters (optional) ", padding=8)
        flt.grid(row=2, column=0, sticky="ew")
        flt.columnconfigure(1, weight=1)
        self.filter_entries: dict[str, ttk.Entry] = {}
        for i, (key, label) in enumerate([
            ("sender", "From contains:"),
            ("recipients", "To / CC contains:"),
            ("subject", "Subject contains:"),
            ("body_words", "Words in the text:"),
            ("category", "Category contains:"),
        ]):
            ttk.Label(flt, text=label).grid(row=i, column=0, sticky="w", pady=1)
            e = ttk.Entry(flt)
            e.grid(row=i, column=1, sticky="ew", padx=6, pady=1)
            self.filter_entries[key] = e
        ttk.Label(flt, text="Attachments:").grid(row=5, column=0, sticky="w", pady=1)
        self.attach = ttk.Combobox(flt, state="readonly", width=24,
                                   values=["Any", "Only with attachments", "Only without attachments"])
        self.attach.current(0)
        self.attach.grid(row=5, column=1, sticky="w", padx=6, pady=1)
        ttk.Label(flt, text="Separate several names or words with ;  (any of them matches). "
                            "'Words in the text' is slower.", foreground="#555"
                  ).grid(row=6, column=0, columnspan=2, sticky="w")

        # 5. output
        out = ttk.LabelFrame(col_b, text=" 5. Save as ", padding=8)
        out.grid(row=0, column=0, sticky="nsew")
        out.columnconfigure(1, weight=1)
        self.make_pst = tk.BooleanVar(value=True)
        self.keep_open = tk.BooleanVar(value=True)
        ttk.Checkbutton(out, text="PST file (opens in Outlook)",
                        variable=self.make_pst).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(out, text="Show it in Outlook when done",
                        variable=self.keep_open).grid(row=1, column=0, columnspan=3, sticky="w", padx=(22, 0))
        self.fmt_vars: dict[str, tk.BooleanVar] = {}
        for i, (k, label) in enumerate(FILE_FORMATS.items()):
            v = tk.BooleanVar(value=False)
            self.fmt_vars[k] = v
            ttk.Checkbutton(out, text=label, variable=v).grid(row=3 + i, column=0, columnspan=3, sticky="w")
        ttk.Label(out, text="Also save each item as its own file:", foreground="#555"
                  ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.save_atts = tk.BooleanVar(value=False)
        self.write_index = tk.BooleanVar(value=True)
        ttk.Checkbutton(out, text="Attachments as separate files", variable=self.save_atts
                        ).grid(row=7, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(out, text="Excel list of all items (.csv)", variable=self.write_index
                        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Label(out, text="Sort files by:").grid(row=9, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.layout = ttk.Combobox(out, state="readonly", width=26, values=list(LAYOUTS.values()))
        self.layout.current(0)
        self.layout.grid(row=10, column=0, columnspan=3, sticky="w")
        ttk.Label(out, text="Save in folder:").grid(row=11, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.out_dir = ttk.Entry(out, width=24)
        self.out_dir.insert(0, os.path.join(os.path.expanduser("~"), "Outlook Archives"))
        self.out_dir.grid(row=12, column=0, columnspan=2, sticky="ew")
        ttk.Button(out, text="Browse", width=7, command=self._browse).grid(row=12, column=2, padx=(4, 0))
        ttk.Label(out, text="Archive name:").grid(row=13, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.name = ttk.Entry(out, width=24)
        self.name.insert(0, f"Outlook archive {dt.date.today():%Y-%m-%d}")
        self.name.grid(row=14, column=0, columnspan=3, sticky="ew")

        # ---- bottom: actions and log
        bottom = ttk.Frame(r, padding=(12, 0, 12, 10))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        actions = ttk.Frame(bottom)
        actions.grid(row=0, column=0, sticky="ew")
        self.btn_count = ttk.Button(actions, text="Count matching items", command=self._on_count)
        self.btn_count.pack(side="left")
        self.btn_go = ttk.Button(actions, text="Start archiving", command=self._on_archive)
        self.btn_go.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(actions, text="Stop", command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left")
        self.btn_open = ttk.Button(actions, text="Open archive folder", command=self._open_folder,
                                   state="disabled")
        self.btn_open.pack(side="right")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=6)
        self.progress_text = ttk.Label(bottom, text="", foreground="#555")
        self.progress_text.grid(row=2, column=0, sticky="w")
        logf = ttk.Frame(bottom)
        logf.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        logf.columnconfigure(0, weight=1)
        self.log = tk.Text(logf, height=7, wrap="word", state="disabled", relief="solid", borderwidth=1)
        self.log.grid(row=0, column=0, sticky="ew")
        lsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        lsb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=lsb.set)
        self._say("Nothing is deleted from Outlook - this tool only makes copies.")

    def _scrollable(self, paned: ttk.PanedWindow) -> ttk.Frame:
        outer = ttk.Frame(paned)
        paned.add(outer, weight=4)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        bg = ttk.Style().lookup("TFrame", "background") or self.root.cget("background")
        canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0, background=bg)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew", padx=(8, 0))
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        inner = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def resize(_e=None) -> None:
            w = max(canvas.winfo_width(), inner.winfo_reqwidth())
            h = max(canvas.winfo_height(), inner.winfo_reqheight())
            canvas.itemconfigure(win, width=w, height=h)
            canvas.configure(scrollregion=(0, 0, w, h))

        inner.bind("<Configure>", resize)
        canvas.bind("<Configure>", resize)

        def wheel(e) -> None:
            if inner.winfo_reqheight() > canvas.winfo_height():
                canvas.yview_scroll(-1 if (e.delta > 0 or e.num == 4) else 1, "units")

        canvas.bind("<Enter>", lambda _e: (canvas.bind_all("<MouseWheel>", wheel),
                                           canvas.bind_all("<Button-4>", wheel),
                                           canvas.bind_all("<Button-5>", wheel)))
        canvas.bind("<Leave>", lambda _e: (canvas.unbind_all("<MouseWheel>"),
                                           canvas.unbind_all("<Button-4>"),
                                           canvas.unbind_all("<Button-5>")))
        return inner

    @staticmethod
    def _quick_options() -> list[str]:
        this_year = dt.date.today().year
        opts = ["Everything (no date limit)", "Last 12 months"]
        opts += [f"Older than {n} year{'s' if n > 1 else ''}" for n in (1, 2, 3, 5)]
        opts += [f"Whole year {y}" for y in range(this_year, this_year - 26, -1)]
        return opts

    # ============================================================ folder tree
    def _fill_tree(self, stores: list[FolderNode]) -> None:
        self.tree.delete(*self.tree.get_children())
        self.nodes.clear()
        self.parents.clear()
        self.checked.clear()

        def add(parent_id: str, node: FolderNode) -> None:
            self.nodes[node.key] = node
            if parent_id:
                self.parents[node.key] = parent_id
            self.tree.insert(parent_id, "end", iid=node.key, open=node.is_store)
            for child in node.children:
                add(node.key, child)

        for s in stores:
            add("", s)
        if stores:  # tick the main mailbox by default
            self._set_checked(stores[0].key, True)
        self._refresh_all()

    def _descendants(self, key: str) -> list[str]:
        out = [key]
        for child in self.tree.get_children(key):
            out.extend(self._descendants(child))
        return out

    def _set_checked(self, key: str, on: bool) -> None:
        for k in self._descendants(key):
            (self.checked.add if on else self.checked.discard)(k)

    def _refresh_all(self) -> None:
        for key, node in self.nodes.items():
            sub = self._descendants(key)
            n_on = sum(1 for k in sub if k in self.checked)
            state = "on" if n_on == len(sub) else ("off" if n_on == 0 else "some")
            count = "" if node.is_store else f"  ({node.count:,})"
            self.tree.item(key, text=f" {node.name}{count}", image=self.boxes[state])

    def _toggle(self, key: str) -> None:
        if self.busy or key not in self.nodes:
            return
        sub = self._descendants(key)
        all_on = all(k in self.checked for k in sub)
        self._set_checked(key, not all_on)
        # a parent stays ticked only while it is in the list itself
        self._refresh_all()

    def _on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "tree":
            return None
        if self.tree.identify_element(event.x, event.y) in ("Treeitem.indicator", "indicator"):
            return None  # let the arrow open/close the branch
        key = self.tree.identify_row(event.y)
        if key:
            self._toggle(key)
            self.tree.focus(key)
        return "break"

    def _on_tree_space(self, _event):
        key = self.tree.focus()
        if key:
            self._toggle(key)
        return "break"

    def _set_all(self, on: bool) -> None:
        if self.busy:
            return
        for key in self.tree.get_children(""):
            self._set_checked(key, on)
        self._refresh_all()

    def _chosen_folders(self) -> list[str]:
        return [k for k in self.nodes if k in self.checked and not self.nodes[k].is_store]

    # ============================================================== choices
    def _on_quick(self, _event=None) -> None:
        choice = self.quick.get()
        today = dt.date.today()
        a, b = "", ""
        if choice.startswith("Whole year"):
            y = int(choice.split()[-1])
            a, b = f"{y}-01-01", f"{y}-12-31"
        elif choice == "Last 12 months":
            a, b = (today - dt.timedelta(days=365)).isoformat(), today.isoformat()
        elif choice.startswith("Older than"):
            n = int(choice.split()[2])
            try:
                cut = today.replace(year=today.year - n)
            except ValueError:  # 29 February
                cut = today.replace(year=today.year - n, day=28)
            b = (cut - dt.timedelta(days=1)).isoformat()
        for entry, value in ((self.date_from, a), (self.date_to, b)):
            entry.delete(0, "end")
            entry.insert(0, value)
        if self.name.get().startswith("Outlook archive"):
            self.name.delete(0, "end")
            suffix = choice.replace("Whole year ", "").replace("Everything (no date limit)", "all")
            self.name.insert(0, f"Outlook archive {suffix}".strip())

    def _criteria(self) -> Criteria:
        kinds = {k for k, v in self.kind_vars.items() if v.get()}
        if not kinds:
            raise ValueError("Tick at least one kind of item in step 2.")
        a = parse_date(self.date_from.get())
        b = parse_date(self.date_to.get())
        if a and b and a > b:
            raise ValueError("The 'From' date is after the 'To' date.")
        field_key = list(DATE_FIELDS)[self.date_field.current()]
        attach = ["any", "with", "without"][self.attach.current()]
        texts = {k: e.get() for k, e in self.filter_entries.items()}
        return Criteria(kinds=kinds, date_field=field_key, date_from=a, date_to=b,
                        attachments=attach, **texts)

    def _options(self) -> ExportOptions:
        out_dir = self.out_dir.get().strip()
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        opts = ExportOptions(
            output_dir=out_dir,
            archive_name=self.name.get().strip(),
            make_pst=self.make_pst.get(),
            keep_pst_open=self.keep_open.get(),
            file_formats={k for k, v in self.fmt_vars.items() if v.get()},
            save_attachments=self.save_atts.get(),
            layout=list(LAYOUTS)[self.layout.current()],
            write_index=self.write_index.get(),
        )
        opts.validate()
        return opts

    def _browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self.out_dir.get() or os.path.expanduser("~"))
        if d:
            self.out_dir.delete(0, "end")
            self.out_dir.insert(0, os.path.normpath(d))

    # ============================================================== actions
    def _gather(self) -> tuple[list[str], Criteria] | None:
        folders = self._chosen_folders()
        if not folders:
            messagebox.showinfo(APP_TITLE, "Tick at least one folder in step 1.")
            return None
        try:
            return folders, self._criteria()
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return None

    def _on_count(self) -> None:
        got = self._gather()
        if got:
            folders, crit = got
            self._say(f"Counting items ({crit.describe()}) ...")
            self._run(lambda: self._job_count(folders, crit))

    def _on_archive(self) -> None:
        got = self._gather()
        if not got:
            return
        try:
            opts = self._options()
        except (ValueError, OSError) as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        folders, crit = got
        self._say(f"Finding items to archive ({crit.describe()}) ...")
        self._run(lambda: self._job_find_for_archive(folders, crit, opts))

    def _on_stop(self) -> None:
        self.cancel.set()
        self._say("Stopping after the current item ...")

    def _open_folder(self) -> None:
        if not self.last_folder:
            return
        if sys.platform == "win32":
            os.startfile(self.last_folder)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", self.last_folder])  # noqa: S603,S607

    # ======================================================= worker thread
    def _run(self, job: Callable[[], None], follow_on: bool = False) -> None:
        """Queue work for the background thread (one job at a time)."""
        if self.busy and not follow_on:
            return
        self.cancel.clear()
        self.pending += 1
        self._set_busy(True)
        self.jobs.put(job)

    def _worker(self) -> None:
        while True:
            job = self.jobs.get()
            try:
                job()
            except Cancelled:
                self.events.put(("log", "Stopped. Nothing more was copied."))
            except BackendError as e:
                self.events.put(("error", str(e)))
            except Exception as e:  # noqa: BLE001
                self.events.put(("error", describe_exception(e)))
            finally:
                self.events.put(("idle",))

    def _progress(self, text: str, done: int, total: int) -> None:
        self.events.put(("progress", text, done, total))

    def _job_connect(self) -> None:
        if self.backend is None:
            self.backend = self.make_backend()
        msg = self.backend.connect()
        self.events.put(("status", msg))
        self.events.put(("progress", "Reading the folder list ...", 0, 0))
        self.events.put(("tree", self.backend.list_tree()))
        self.events.put(("progress", "Ready.", 0, 0))

    def _need_backend(self) -> Backend:
        if self.backend is None:
            raise BackendError("Not connected to Outlook yet. Click 'Reload folders'.")
        return self.backend

    def _job_count(self, folders: list[str], crit: Criteria) -> None:
        items = self._need_backend().scan(folders, crit, self._progress, self.cancel.is_set)
        self.events.put(("log", summarize(items)))

    def _job_find_for_archive(self, folders: list[str], crit: Criteria, opts: ExportOptions) -> None:
        items = self._need_backend().scan(folders, crit, self._progress, self.cancel.is_set)
        self.events.put(("confirm", items, opts))

    def _job_export(self, items: list[ItemInfo], opts: ExportOptions) -> None:
        res = self._need_backend().export(items, opts, self._progress, self.cancel.is_set)
        self.events.put(("finished", res))

    # ============================================================ UI thread
    def _poll(self) -> None:
        try:
            while True:
                self._handle(self.events.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _handle(self, ev: tuple) -> None:
        kind = ev[0]
        if kind == "status":
            self.status.configure(text=ev[1])
            self._say(ev[1])
        elif kind == "tree":
            self._fill_tree(ev[1])
        elif kind == "progress":
            _, text, done, total = ev
            self.progress_text.configure(text=text)
            if total:
                self.progress.configure(maximum=total, value=done)
            else:
                self.progress.configure(maximum=1, value=0)
        elif kind == "log":
            self._say(ev[1])
        elif kind == "error":
            self.status.configure(text="Problem - see the message.")
            self._say(ev[1])
            messagebox.showerror(APP_TITLE, ev[1])
        elif kind == "confirm":
            self._confirm(ev[1], ev[2])
        elif kind == "finished":
            self._finished(ev[1])
        elif kind == "idle":
            self.pending -= 1
            if self.pending <= 0:
                self.pending = 0
                self._set_busy(False)

    def _confirm(self, items: list[ItemInfo], opts: ExportOptions) -> None:
        self._say(summarize(items))
        if not items:
            messagebox.showinfo(APP_TITLE, "No items match these choices, so there is nothing to archive.")
            return
        what = []
        if opts.make_pst:
            what.append("a PST file")
        if opts.file_formats:
            what.append("separate files (" + ", ".join(sorted(f.upper() for f in opts.file_formats)) + ")")
        if opts.save_attachments:
            what.append("attachments")
        if opts.write_index:
            what.append("an Excel list")
        msg = (f"Copy {len(items):,} items into {', '.join(what)}?\n\n"
               f"Saved in: {os.path.join(opts.output_dir, opts.archive_name)}\n\n"
               "Nothing will be deleted from Outlook. Large archives can take a while; "
               "you can keep working, but leave Outlook open.")
        if messagebox.askyesno(APP_TITLE, msg):
            self._run(lambda: self._job_export(items, opts), follow_on=True)
        else:
            self._say("Archiving cancelled.")

    def _finished(self, res: ExportResult) -> None:
        self.last_folder = res.folder
        self.btn_open.configure(state="normal")
        head = "Stopped part way. " if res.cancelled else "Finished. "
        lines = [f"{head}{res.total:,} items archived into {res.folder}"]
        if res.pst_path:
            lines.append(f"PST file: {res.pst_path}")
        if res.files_written:
            lines.append(f"Files written: {res.files_written:,}")
        if res.errors:
            lines.append(f"{len(res.errors)} problem(s) - the list is in README.txt in the archive folder:")
            lines += ["  - " + e for e in res.errors[:10]]
            if len(res.errors) > 10:
                lines.append(f"  ... and {len(res.errors) - 10} more.")
        self._say("\n".join(lines))
        self.status.configure(text=head.strip())
        messagebox.showinfo(APP_TITLE, "\n".join(lines[:3]) + (
            f"\n\n{len(res.errors)} item(s) had problems - see README.txt." if res.errors else ""))

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.btn_count.configure(state=state)
        self.btn_go.configure(state=state)
        self.btn_stop.configure(state="normal" if busy else "disabled")
        self.root.configure(cursor="watch" if busy else "")

    def _say(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"[{dt.datetime.now():%H:%M:%S}] {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def run(make_backend: Callable[[], Backend]) -> None:
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # sharp text on high-DPI screens
        except Exception:  # noqa: BLE001
            pass
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    App(root, make_backend)
    root.mainloop()
