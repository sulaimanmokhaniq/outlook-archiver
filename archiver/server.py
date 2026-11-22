"""Runs the program's screen: a web page served only to this laptop.

The page (web/) talks to this server; all Outlook work happens on one
background thread. Nothing is reachable from other computers: the server
listens on 127.0.0.1 only, checks the Host header, and every /api call must
carry a random token that only the window opened by this program knows.
"""
from __future__ import annotations

import json
import os
import queue
import secrets
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Callable
from urllib.parse import urlparse

from . import __version__
from .base import Backend, BackendError, Cancelled
from .core import Criteria, ExportOptions, estimate_archive_size, free_space, preview_row, summarize
from .i18n import LANGS, Translator
from .storage import History, Settings, app_dir
from .win_icon import WindowIcon

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")
STATIC = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/index.html": ("index.html", "text/html; charset=utf-8"),
    "/app.js": ("app.js", "text/javascript; charset=utf-8"),
    "/style.css": ("style.css", "text/css; charset=utf-8"),
    "/app.ico": ("app.ico", "image/x-icon"),
    "/favicon.ico": ("app.ico", "image/x-icon"),
    **{f"/icon-{s}.png": (f"icon-{s}.png", "image/png") for s in (32, 48, 64, 128, 192, 256, 512)},
}
PREVIEW_LIMIT = 500
IDLE_LIMIT = 180  # seconds without the window asking for news before quitting


class Busy(Exception):
    pass


class Controller:
    """Holds the program's state and runs one job at a time."""

    def __init__(self, engines, settings: Settings | None = None,
                 history: History | None = None) -> None:
        """engines: {"classic": factory, "online": factory}; a factory gets the settings dict."""
        if callable(engines):  # a single data source (used by the tests)
            engines = {"classic": lambda _settings, make=engines: make()}
        self.engines: dict[str, Callable[[dict], Backend]] = engines
        self.engine = ""  # the connection in use: classic | online
        self.classic_failed = False
        self._classic: Backend | None = None
        self.settings = settings or Settings()
        self.history = history or History()
        self.backend: Backend | None = None
        self.lock = threading.RLock()
        self.jobs: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.job: dict | None = None
        self.connected = False
        self.demo = False
        self.tree: list[dict] = []
        self.tree_version = 0
        self.scan: dict | None = None
        self.scan_seq = 0
        self.result: dict | None = None
        self.result_seq = 0
        self.error: str | None = None
        self.error_seq = 0
        self.stopped_seq = 0
        self.history_count = self.history.count()
        threading.Thread(target=self._worker, daemon=True, name="outlook-worker").start()

    # ------------------------------------------------------------- helpers
    def t(self) -> Translator:
        return Translator(self.settings.load().get("lang") or "en")

    def submit(self, kind: str, fn: Callable[[], None]) -> None:
        with self.lock:
            if self.busy:
                raise Busy()
            self.busy = True
            self.cancel.clear()
            self.error = None
            self.job = {"kind": kind, "text": "", "done": 0, "total": 0, "started": time.time()}
        self.jobs.put(fn)

    def _worker(self) -> None:
        while True:
            fn = self.jobs.get()
            try:
                fn()
            except Cancelled:
                with self.lock:
                    self.stopped_seq += 1
            except BackendError as e:
                self._fail(str(e))
            except Exception as e:  # noqa: BLE001
                _log_problem()
                self._fail(self.t()("e_unexpected", err=str(e) or e.__class__.__name__))
            finally:
                with self.lock:
                    self.busy = False
                    self.job = None

    def _fail(self, text: str) -> None:
        with self.lock:
            self.error = text
            self.error_seq += 1

    def progress(self, text: str, done: int, total: int) -> None:
        with self.lock:
            if self.job is not None:
                self.job.update(text=text, done=done, total=total)

    def _backend(self) -> Backend:
        if self.backend is None or not self.connected:
            raise BackendError(self.t()("e_not_connected"))
        self.backend.t = self.t()
        return self.backend

    # ---------------------------------------------------------------- jobs
    def connect(self, engine: str | None = None) -> None:
        """Connect to classic Outlook or to Microsoft 365 online.

        Without a choice, the saved one is used; "auto" means classic Outlook,
        and if it is missing the screen offers the online mode instead (it
        never opens a sign-in page by itself).
        """
        settings = self.settings.load()
        chosen = engine or settings.get("engine") or "auto"
        if chosen not in self.engines:
            chosen = "classic" if "classic" in self.engines else next(iter(self.engines))

        def run() -> None:
            t = self.t()
            with self.lock:
                self.connected = False
                self.tree = []
                self.tree_version += 1
                self.scan = None
            if chosen == "classic":
                if self._classic is None:
                    self._classic = self.engines["classic"](settings)
                backend = self._classic
            else:
                backend = self.engines[chosen](self.settings.load())
            backend.t = t
            self.progress(t("p_connecting") if chosen == "classic" else t("p_signing_in"), 0, 0)
            try:
                backend.connect()
            except BackendError as e:
                if chosen == "classic" and "online" in self.engines:
                    with self.lock:
                        self.classic_failed = True
                    raise BackendError(str(e) + "\n\n" + t("e_try_online")) from e
                raise
            self.progress(t("p_reading_tree"), 0, 0)
            tree = [n.to_dict() for n in backend.list_tree()]
            with self.lock:
                self.backend = backend
                self.engine = chosen
                self.connected = True
                self.classic_failed = chosen != "classic" and self.classic_failed
                self.demo = backend.is_demo
                self.tree = tree
                self.tree_version += 1
            if engine:  # remembered only when picked on purpose
                self.settings.save({"engine": engine})
        self.submit("connect", run)

    def sign_out(self) -> None:
        with self.lock:
            if self.busy:
                raise Busy()
            backend = self.backend
            if self.engine == "online" and backend is not None:
                backend.sign_out()
                self.connected = False
                self.backend = None
                self.engine = ""
                self.tree = []
                self.tree_version += 1
                self.scan = None
        self.settings.save({"engine": "auto"})

    def capabilities(self) -> dict:
        b = self.backend
        if not b or not self.connected:
            return {}
        return {"pst": b.supports_pst, "kinds": list(b.kinds), "formats": list(b.formats)}

    def count(self, body: dict) -> None:
        folders = [k for k in body.get("folders", []) if isinstance(k, str)]
        crit_raw = body.get("criteria") or {}
        crit = Criteria.from_dict(crit_raw)  # raises ValueError on a bad date
        skip = bool(body.get("skip_archived", True))
        key = json.dumps([sorted(folders), crit_raw, skip], sort_keys=True, ensure_ascii=False)

        def run() -> None:
            items = self._backend().scan(folders, crit, self.progress, self.cancel.is_set)
            skipped = 0
            if skip:
                items, skipped = self.history.filter_new(items)
            with self.lock:
                self.scan_seq += 1
                self.scan = {
                    "id": self.scan_seq, "key": key, "items": items, "skipped": skipped,
                    "summary": summarize(items),
                    "preview": [preview_row(i) for i in items[:PREVIEW_LIMIT]],
                }
        self.submit("count", run)

    def archive(self, body: dict) -> None:
        t = self.t()
        with self.lock:
            scan = self.scan
        if not scan or scan["id"] != body.get("scan_id"):
            raise ValueError(t("e_stale"))
        raw = dict(body.get("options") or {})
        raw["lang"] = t.lang
        opts = ExportOptions.from_dict(raw)
        backend = self.backend
        if backend is not None:
            if opts.make_pst and not backend.supports_pst:
                raise ValueError(t("e_no_pst_online"))
            opts.file_formats = {f for f in opts.file_formats if f in backend.formats}
        if opts.output_dir:
            try:
                os.makedirs(opts.output_dir, exist_ok=True)
            except OSError:
                pass
        opts.validate()  # raises ValueError with a translated message

        def record(items, path) -> None:
            self.history.record(items, path)

        def run() -> None:
            backend = self._backend()
            try:
                res = backend.export(scan["items"], opts, self.progress, self.cancel.is_set, record)
            finally:
                with self.lock:
                    self.scan = None  # the mailbox has changed; count again next time
            with self.lock:
                self.result = res.to_dict()
                self.result_seq += 1
                self.history_count = self.history.count()
        self.submit("archive", run)

    def estimate(self, body: dict) -> dict:
        """Space before and after, shown before the archive starts."""
        t = self.t()
        with self.lock:
            scan = self.scan
            tree = list(self.tree)
        if not scan or scan["id"] != body.get("scan_id"):
            raise ValueError(t("e_stale"))
        opts = ExportOptions.from_dict(body.get("options") or {})
        items = scan["items"]
        backend = self.backend
        archive = estimate_archive_size(items, opts)
        free = free_space(opts.output_dir)
        moving = opts.move_after and opts.make_pst

        def removable(it) -> bool:
            return it.removal_blocker() is None and not (backend and backend.in_protected_folder(it))

        mailboxes = []
        for node in tree:
            mine = [it for it in items if it.store_name == node["name"]]
            if not mine:
                continue
            mailboxes.append({
                "name": node["name"], "size": node.get("size"),
                "chosen": sum(it.size for it in mine),
                "removed": sum(it.size for it in mine if removable(it)) if moving else 0,
                "file_size": node.get("file_size"), "file_ext": node.get("file_ext", ""),
            })
        drive = os.path.splitdrive(os.path.abspath(opts.output_dir or "."))[0]
        return {
            "items_size": sum(it.size for it in items), "archive": archive, "free": free,
            "enough": free is None or archive + 200 * 1024 * 1024 < free,
            "move": moving, "drive": drive, "mailboxes": mailboxes,
        }

    def forget_history(self) -> None:
        self.history.forget()
        with self.lock:
            self.history_count = 0
            self.scan = None

    # --------------------------------------------------------------- state
    def state(self) -> dict:
        with self.lock:
            job = None
            if self.job:
                job = dict(self.job)
                job["elapsed"] = round(time.time() - job.pop("started"), 1)
            scan = None
            if self.scan:
                scan = {k: self.scan[k] for k in ("id", "key", "skipped", "summary")}
            return {
                "version": __version__,
                "busy": self.busy,
                "job": job,
                "connected": self.connected,
                "demo": self.demo,
                "engine": self.engine,
                "engines": list(self.engines),
                "classic_failed": self.classic_failed,
                "needs_client_id": bool(getattr(self.engines.get("online"), "needs_client_id", False)),
                "account": self.backend.account if self.backend and self.connected else "",
                "capabilities": self.capabilities(),
                "store_count": len(self.tree),
                "tree_version": self.tree_version,
                "scan": scan,
                "result": self.result,
                "result_seq": self.result_seq,
                "error": self.error,
                "error_seq": self.error_seq,
                "stopped_seq": self.stopped_seq,
                "history_count": self.history_count,
            }

    def scan_preview(self) -> dict:
        with self.lock:
            if not self.scan:
                return {"id": None, "rows": [], "more": 0}
            rows = self.scan["preview"]
            return {"id": self.scan["id"], "rows": rows,
                    "more": max(0, self.scan["summary"]["total"] - len(rows))}


def _log_problem() -> None:
    """Keep technical details in log.txt next to the settings, for support."""
    try:
        with open(os.path.join(app_dir(), "log.txt"), "a", encoding="utf-8") as f:
            f.write(f"--- {time.strftime('%Y-%m-%d %H:%M:%S')} (version {__version__})\n")
            traceback.print_exc(file=f)
    except OSError:
        pass


# ------------------------------------------------------------------- web


def pick_folder(initial: str) -> str:
    """Show Windows' own 'choose a folder' window."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        path = filedialog.askdirectory(parent=root, mustexist=False,
                                       initialdir=initial if os.path.isdir(initial or "") else None)
    finally:
        root.destroy()
    return os.path.normpath(path) if path else ""


def open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", path])  # noqa: S603,S607


class Server:
    def __init__(self, controller: Controller, port: int = 0) -> None:
        self.controller = controller
        self.token = secrets.token_urlsafe(24)
        self.last_seen = time.time()
        self.bye_at: float | None = None
        handler = type("Handler", (_Handler,), {"app": self})
        self.httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
        self.httpd.daemon_threads = True
        self.port = self.httpd.server_address[1]
        self.url = f"http://127.0.0.1:{self.port}/?t={self.token}"

    def start(self) -> None:
        threading.Thread(target=self.httpd.serve_forever, daemon=True, name="http").start()

    def stop(self) -> None:
        self.httpd.shutdown()

    def should_quit(self, window: subprocess.Popen | None, window_started: float) -> bool:
        if self.controller.busy:
            return False
        now = time.time()
        if window is not None and window.poll() is not None and now - window_started > 8:
            return True  # the app window was closed
        if self.bye_at and now - self.bye_at > 6 and self.last_seen < self.bye_at:
            return True  # the page was closed and did not come back
        return now - self.last_seen > IDLE_LIMIT


class _Handler(BaseHTTPRequestHandler):
    app: Server
    server_version = "OutlookArchiver"
    protocol_version = "HTTP/1.1"

    def log_message(self, *args) -> None:  # keep the console quiet
        pass

    # ---- plumbing
    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").lower()
        return host in (f"127.0.0.1:{self.app.port}", f"localhost:{self.app.port}")

    def _send(self, code: int, body: bytes, ctype: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, data) -> None:
        self._send(code, json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _authorized(self) -> bool:
        return secrets.compare_digest(self.headers.get("X-Token", ""), self.app.token)

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if n > 2_000_000:
            raise ValueError("Request too large")
        raw = self.rfile.read(n) if n else b"{}"
        data = json.loads(raw.decode("utf-8") or "{}")
        return data if isinstance(data, dict) else {}

    # ---- routes
    def do_GET(self) -> None:  # noqa: N802
        if not self._host_ok():
            return self._send(403, b"Forbidden", "text/plain")
        path = urlparse(self.path).path
        if path in STATIC:
            name, ctype = STATIC[path]
            with open(os.path.join(WEB_DIR, name), "rb") as f:
                return self._send(200, f.read(), ctype)
        if not path.startswith("/api/"):
            return self._send(404, b"Not found", "text/plain")
        if not self._authorized():
            return self._json(401, {"error": "unauthorized"})
        self.app.last_seen = time.time()
        c = self.app.controller
        if path == "/api/state":
            return self._json(200, c.state())
        if path == "/api/settings":
            return self._json(200, c.settings.load())
        if path == "/api/tree":
            with c.lock:
                return self._json(200, {"version": c.tree_version, "tree": c.tree})
        if path == "/api/scan":
            return self._json(200, c.scan_preview())
        return self._json(404, {"error": "not found"})

    def do_POST(self) -> None:  # noqa: N802
        if not self._host_ok():
            return self._send(403, b"Forbidden", "text/plain")
        path = urlparse(self.path).path
        if not self._authorized():
            return self._json(401, {"error": "unauthorized"})
        self.app.last_seen = time.time()
        c = self.app.controller
        try:
            body = self._body()
            if path == "/api/connect":
                engine = body.get("engine")
                c.connect(engine if engine in c.engines else None)
            elif path == "/api/signout":
                c.sign_out()
            elif path == "/api/count":
                c.count(body)
            elif path == "/api/archive":
                c.archive(body)
            elif path == "/api/stop":
                c.cancel.set()
            elif path == "/api/settings":
                if body.get("lang") not in (None, "", *LANGS):
                    body.pop("lang")
                return self._json(200, c.settings.save(body))
            elif path == "/api/estimate":
                return self._json(200, c.estimate(body))
            elif path == "/api/forget":
                c.forget_history()
            elif path == "/api/browse":
                return self._json(200, {"path": pick_folder(str(body.get("current", "")))})
            elif path == "/api/open":
                target = str(body.get("path", ""))
                if not os.path.isdir(target):
                    return self._json(400, {"error": "not a folder"})
                open_path(target)
            elif path == "/api/bye":
                self.app.bye_at = time.time()
                self.app.last_seen = self.app.bye_at - 0.001
            else:
                return self._json(404, {"error": "not found"})
        except Busy:
            return self._json(409, {"error": c.t()("e_busy")})
        except ValueError as e:
            return self._json(400, {"error": str(e)})
        return self._json(200, {"ok": True})


# ---------------------------------------------------------------- window


def find_app_browser() -> str | None:
    """Edge (built into Windows) or Chrome, which can show a page as an app window."""
    if sys.platform != "win32":
        return None
    for env in ("PROGRAMFILES(X86)", "PROGRAMFILES", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if not base:
            continue
        for rel in (r"Microsoft\Edge\Application\msedge.exe", r"Google\Chrome\Application\chrome.exe"):
            p = os.path.join(base, rel)
            if os.path.isfile(p):
                return p
    return None


def open_window(url: str) -> subprocess.Popen | None:
    exe = find_app_browser()
    if exe:
        profile = os.path.join(app_dir(), "window")
        try:
            return subprocess.Popen([  # noqa: S603
                exe, f"--app={url}", f"--user-data-dir={profile}", "--no-first-run",
                "--no-default-browser-check", "--disable-features=Translate", "--window-size=1360,900",
            ])
        except OSError:
            pass
    webbrowser.open(url)
    return None


def run_app(engines, show_window: bool = True, port: int = 0) -> None:
    controller = Controller(engines)
    server = Server(controller, port)
    server.start()
    controller.connect()
    print(f"Outlook Archiver {__version__} is running at {server.url}", flush=True)
    window = open_window(server.url) if show_window else None
    started = time.time()
    icon = WindowIcon(os.path.join(WEB_DIR, "app.ico")) if show_window else None
    try:
        while not server.should_quit(window, started):
            if icon:
                icon.apply()  # sharp taskbar icon (Windows only)
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    server.stop()
