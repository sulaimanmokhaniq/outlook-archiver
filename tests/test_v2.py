"""Remembering, skipping, moving, Arabic output and the local server."""
import csv
import datetime as dt
import http.client
import json
import os
import time

import pytest

from archiver.base import BackendError
from archiver.core import Criteria, ExportOptions, ItemInfo
from archiver.demo_backend import DemoBackend
from archiver.i18n import MESSAGES
from archiver.server import Controller, Server
from archiver.storage import History, Settings


@pytest.fixture(autouse=True)
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("OUTLOOK_ARCHIVER_HOME", str(tmp_path / "home"))
    return tmp_path


def _noop(*_a):
    return None


def _all_keys(b):
    tree = b.list_tree()
    keys = []

    def walk(n):
        if not n.is_store:
            keys.append(n.key)
        for c in n.children:
            walk(c)
    for s in tree:
        walk(s)
    return keys


def test_both_languages_have_the_same_messages():
    assert set(MESSAGES["en"]) == set(MESSAGES["ar"])


def test_settings_remember_and_ignore_unknown(tmp_path):
    s = Settings(str(tmp_path / "s.json"))
    assert s.load()["lang"] == ""
    s.save({"lang": "ar", "kinds": ["mail"], "evil": 1})
    data = Settings(str(tmp_path / "s.json")).load()
    assert data["lang"] == "ar" and data["kinds"] == ["mail"] and "evil" not in data


def test_removal_rules():
    now = dt.datetime(2026, 1, 1)
    mail = ItemInfo("1", "s", "f", ("Inbox",), "M", "IPM.Note")
    assert mail.removal_blocker(now) is None
    contact = ItemInfo("1", "s", "f", ("C",), "M", "IPM.Contact")
    assert contact.removal_blocker(now) == "kind"
    past = ItemInfo("1", "s", "f", ("Cal",), "M", "IPM.Appointment",
                    start=dt.datetime(2020, 1, 1), end=dt.datetime(2020, 1, 1, 1))
    assert past.removal_blocker(now) is None
    past.recurring = True
    assert past.removal_blocker(now) == "recurring"
    future = ItemInfo("1", "s", "f", ("Cal",), "M", "IPM.Appointment",
                      start=dt.datetime(2027, 1, 1), end=dt.datetime(2027, 1, 1, 1))
    assert future.removal_blocker(now) == "future"


def test_skip_already_archived(tmp_path):
    b = DemoBackend()
    keys = _all_keys(b)
    crit = Criteria(kinds={"mail"}, date_from=dt.date(2020, 1, 1), date_to=dt.date(2020, 12, 31))
    items = b.scan(keys, crit, _noop, lambda: False)
    hist = History(str(tmp_path / "h.sqlite3"))
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="A", write_index=True, make_pst=False)
    res = b.export(items[:10], opts, _noop, lambda: False, hist.record)
    new, skipped = hist.filter_new(items)
    assert skipped == 10 and len(new) == len(items) - 10
    # deleting the old archive folder means those items are archived again
    import shutil
    shutil.rmtree(res.folder)
    new, skipped = hist.filter_new(items)
    assert skipped == 0


def test_move_removes_only_safe_items(tmp_path):
    b = DemoBackend()
    keys = _all_keys(b)
    crit = Criteria(kinds={"mail", "appointment", "contact"}, date_to=dt.date(2020, 12, 31))
    items = b.scan(keys, crit, _noop, lambda: False)
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="Move", move_after=True)
    res = b.export(items, opts, _noop, lambda: False)
    assert not res.errors
    assert res.removed > 0 and res.kept > 0
    kept_kinds = {"contact"}
    after = b.scan(keys, crit, _noop, lambda: False)
    left = {(it.kind, it.recurring, it.folder_parts[0]) for it in after}
    # what is left: contacts, recurring meetings, and Deleted Items content
    for kind, recurring, folder in left:
        assert kind in kept_kinds or recurring or folder == "Deleted Items"
    with open(os.path.join(res.folder, "Archive list.csv"), encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    assert sum(1 for r in rows[1:] if r[-1] == "Yes") == res.removed


def test_move_needs_pst(tmp_path):
    with pytest.raises(ValueError):
        ExportOptions(output_dir=str(tmp_path), archive_name="x", make_pst=False, move_after=True).validate()


def test_cancelled_run_removes_nothing(tmp_path):
    b = DemoBackend()
    keys = _all_keys(b)
    items = b.scan(keys, Criteria(kinds={"mail"}), _noop, lambda: False)
    calls = {"n": 0}

    def cancel_midway():
        calls["n"] += 1
        return calls["n"] > 20

    opts = ExportOptions(output_dir=str(tmp_path), archive_name="C", move_after=True)
    res = b.export(items, opts, _noop, cancel_midway)
    assert res.cancelled and res.removed == 0


def test_arabic_output(tmp_path):
    b = DemoBackend()
    keys = _all_keys(b)
    items = b.scan(keys, Criteria(kinds={"mail"}, date_from=dt.date(2021, 1, 1),
                                  date_to=dt.date(2021, 3, 31)), _noop, lambda: False)
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="أرشيف", lang="ar",
                         file_formats={"txt"})
    res = b.export(items, opts, _noop, lambda: False)
    names = os.listdir(res.folder)
    assert "قائمة الأرشيف.csv" in names and "اقرأني.txt" in names and "الملفات" in names
    with open(os.path.join(res.folder, "قائمة الأرشيف.csv"), encoding="utf-8-sig") as f:
        assert next(csv.reader(f))[0] == "النوع"


# ------------------------------------------------------------------ server

class Client:
    def __init__(self, server):
        self.s = server

    def call(self, method, path, body=None, token=True, host=None):
        c = http.client.HTTPConnection("127.0.0.1", self.s.port, timeout=10)
        headers = {"Host": host or f"127.0.0.1:{self.s.port}"}
        if token:
            headers["X-Token"] = self.s.token
        data = None
        if body is not None:
            data = json.dumps(body).encode()
            headers["Content-Type"] = "application/json"
        c.request(method, path, body=data, headers=headers)
        r = c.getresponse()
        raw = r.read()
        try:
            return r.status, json.loads(raw)
        except ValueError:
            return r.status, raw

    def wait_idle(self):
        for _ in range(200):
            _, st = self.call("GET", "/api/state")
            if not st["busy"]:
                return st
            time.sleep(0.05)
        raise AssertionError("still busy")


@pytest.fixture()
def client(tmp_path):
    ctl = Controller(DemoBackend, Settings(str(tmp_path / "s.json")), History(str(tmp_path / "h.sqlite3")))
    srv = Server(ctl)
    srv.start()
    ctl.connect()
    cl = Client(srv)
    cl.wait_idle()
    yield cl
    srv.stop()


def test_server_blocks_strangers(client):
    assert client.call("GET", "/api/state", token=False)[0] == 401
    assert client.call("GET", "/api/state", host="evil.example:80")[0] == 403
    status, body = client.call("GET", "/", token=False)
    assert status == 200 and b"app.js" in body


def test_server_count_then_archive(client, tmp_path):
    _, tree = client.call("GET", "/api/tree")
    inbox = [n for n in tree["tree"][0]["children"] if n["name"] == "Inbox"][0]
    req = {"folders": [inbox["key"]], "criteria": {"kinds": ["mail"], "date_from": "2019-01-01",
                                                   "date_to": "2019-12-31"}, "skip_archived": True}
    assert client.call("POST", "/api/count", req)[0] == 200
    st = client.wait_idle()
    total = st["scan"]["summary"]["total"]
    assert total > 0
    _, preview = client.call("GET", "/api/scan")
    assert len(preview["rows"]) == total
    # archiving with an old scan id is refused
    assert client.call("POST", "/api/archive", {"scan_id": 999, "options": {}})[0] == 400
    opts = {"output_dir": str(tmp_path / "out"), "archive_name": "Test", "make_pst": True,
            "file_formats": ["txt"], "write_index": True}
    status, _ = client.call("POST", "/api/archive", {"scan_id": st["scan"]["id"], "options": opts})
    assert status == 200
    st = client.wait_idle()
    assert st["result"]["total"] == total and not st["result"]["errors"]
    assert st["history_count"] == total
    # counting again skips what was just archived
    client.call("POST", "/api/count", req)
    st = client.wait_idle()
    assert st["scan"]["summary"]["total"] == 0 and st["scan"]["skipped"] == total


def test_server_settings_roundtrip(client):
    client.call("POST", "/api/settings", {"lang": "ar", "period": "older2", "lang_bad": 1})
    _, s = client.call("GET", "/api/settings")
    assert s["lang"] == "ar" and s["period"] == "older2"
    client.call("POST", "/api/settings", {"lang": "xx"})
    assert client.call("GET", "/api/settings")[1]["lang"] == "ar"


class NoOutlook(DemoBackend):
    def connect(self):
        raise BackendError(self.t("e_no_outlook"))


def test_connect_failure_is_reported(tmp_path):
    ctl = Controller(NoOutlook, Settings(str(tmp_path / "s.json")), History(str(tmp_path / "h.sqlite3")))
    ctl.connect()
    for _ in range(100):
        if not ctl.busy:
            break
        time.sleep(0.02)
    st = ctl.state()
    assert not st["connected"] and "classic Outlook" in st["error"] and st["error_seq"] == 1


def test_icons_are_served(client):
    for path in ("/app.ico", "/favicon.ico", "/icon-32.png", "/icon-256.png"):
        status, body = client.call("GET", path, token=False)
        assert status == 200 and len(body) > 500, path


def test_window_icon_is_harmless_off_windows():
    from archiver.win_icon import WindowIcon
    assert WindowIcon("archiver/web/app.ico").apply() == 0


def test_estimate_before_and_after(client, tmp_path):
    _, tree = client.call("GET", "/api/tree")
    assert tree["tree"][0]["size"] > 0
    inbox = [n for n in tree["tree"][0]["children"] if n["name"] == "Inbox"][0]
    client.call("POST", "/api/count", {"folders": [inbox["key"]], "criteria": {"kinds": ["mail"],
                "date_to": "2019-12-31"}, "skip_archived": False})
    st = client.wait_idle()
    scan_id = st["scan"]["id"]
    assert st["scan"]["summary"]["size"] > 0
    opts = {"output_dir": str(tmp_path), "archive_name": "x", "make_pst": True, "write_index": True}
    _, copy_only = client.call("POST", "/api/estimate", {"scan_id": scan_id, "options": opts})
    assert copy_only["archive"] >= copy_only["items_size"] and copy_only["free"] > 0
    assert copy_only["mailboxes"][0]["removed"] == 0 and not copy_only["move"]
    _, more = client.call("POST", "/api/estimate", {"scan_id": scan_id, "options": dict(opts, file_formats=["msg"])})
    assert more["archive"] > copy_only["archive"]
    _, moving = client.call("POST", "/api/estimate", {"scan_id": scan_id, "options": dict(opts, move_after=True)})
    mb = moving["mailboxes"][0]
    assert moving["move"] and 0 < mb["removed"] <= mb["chosen"] <= mb["size"]
    client.call("POST", "/api/archive", {"scan_id": scan_id, "options": dict(opts, move_after=True)})
    st = client.wait_idle()
    r = st["result"]
    assert r["removed_size"] == mb["removed"] and r["archive_size"] > 0


def test_check_confirms_every_item(tmp_path):
    b = DemoBackend()
    keys = _all_keys(b)
    items = b.scan(keys, Criteria(kinds={"mail", "appointment", "contact"}, date_to=dt.date(2018, 12, 31)),
                   _noop, lambda: False)
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="Check", file_formats={"txt"})
    res = b.export(items, opts, _noop, lambda: False)
    c = res.check
    assert c["ok"] and c["missing_count"] == 0
    assert c["pst"]["found"] == c["pst"]["expected"] == len(items)
    assert c["pst"]["found_size"] == c["pst"]["expected_size"] == sum(it.size for it in items)
    assert c["files"]["found"] == c["files"]["expected"] == len(items)
    assert res.total == len(items)
    with open(os.path.join(res.folder, "Archive list.csv"), encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    assert rows[0][9] == "Check" and all(r[9] == "OK" for r in rows[1:])


def test_check_catches_missing_items_and_blocks_removal(tmp_path, monkeypatch):
    monkeypatch.setenv("OUTLOOK_ARCHIVER_DEMO_LOSE", "3")
    b = DemoBackend()
    keys = _all_keys(b)
    crit = Criteria(kinds={"mail"}, date_to=dt.date(2018, 6, 30))
    items = b.scan(keys, crit, _noop, lambda: False)
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="Lost", move_after=True)
    res = b.export(items, opts, _noop, lambda: False)
    c = res.check
    assert not c["ok"] and c["missing_count"] == 3 and len(c["missing"]) == 3
    assert c["pst"]["found"] == len(items) - 3
    assert res.removed == 0  # nothing removed when anything is missing
    assert res.total == len(items) - 3
    assert b.scan(keys, crit, _noop, lambda: False) == items  # mailbox untouched
    with open(os.path.join(res.folder, "README.txt"), encoding="utf-8-sig") as f:
        readme = f.read()
    assert "3 item(s) are missing" in readme and "Not found in the PST" in readme


def test_check_counts_stopped_items(tmp_path):
    b = DemoBackend()
    items = b.scan(_all_keys(b), Criteria(kinds={"mail"}), _noop, lambda: False)[:50]
    calls = {"n": 0}

    def stop_after_10():
        calls["n"] += 1
        return calls["n"] > 10

    res = b.export(items, ExportOptions(output_dir=str(tmp_path), archive_name="S"), _noop, stop_after_10)
    assert res.cancelled and res.check["ok_items"] == 10 and res.check["missing_count"] == 40


# ---------------------------------------------------- Microsoft 365 online mode

def test_online_mode_end_to_end(tmp_path):
    from archiver.demo_online import DemoOnlineBackend
    b = DemoOnlineBackend()
    assert "manager@example.com" in b.connect()
    tree = b.list_tree()
    names = [n.name for n in tree[0].children]
    assert {"Inbox", "Sent Items", "Calendar", "Contacts", "Tasks"} <= set(names)
    keys = list(b._folders)
    items = b.scan(keys, Criteria(date_from=dt.date(2020, 1, 1), date_to=dt.date(2020, 12, 31)), _noop, lambda: False)
    kinds = {it.kind for it in items}
    assert {"mail", "appointment", "contact"} <= kinds and all(it.size >= 0 for it in items)
    assert all(it.display_date.year == 2020 for it in items if it.kind in ("mail", "appointment"))
    with pytest.raises(ValueError):  # no PST online, and removal needs a safe copy
        ExportOptions(output_dir=str(tmp_path), archive_name="x", make_pst=False, move_after=True).validate()
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="Online", make_pst=False,
                         file_formats={"eml", "mht"}, save_attachments=True, move_after=True)
    res = b.export(items, opts, _noop, lambda: False)
    assert not res.errors and res.check["ok"] and res.check["files"]["found"] == res.check["files"]["expected"]
    exts = {os.path.splitext(f)[1] for _r, _d, fs in os.walk(res.folder) for f in fs}
    assert {".eml", ".ics", ".vcf", ".html", ".csv", ".txt"} <= exts
    moved = [m for m in b.http.messages.values() if m["parentFolderId"] == "mf-deleteditems"]
    assert res.removed > 0 and len(moved) >= res.removed
    assert res.kept == sum(1 for it in items if it.kind not in ("mail", "meeting", "other")) + \
        sum(1 for it in items if it.folder_parts[:1] == ("Deleted Items",))


def test_online_graph_paging_and_dates(tmp_path):
    from archiver.demo_online import DemoOnlineBackend
    b = DemoOnlineBackend()
    b.connect()
    b.list_tree()
    inbox = [k for k, v in b._folders.items() if v[2] == ("Inbox",)][0]
    items = list(b.iter_folder(inbox, Criteria(), _noop, lambda: False))
    assert len(items) == 700  # 14 pages of 50 were followed
    demo = DemoBackend()
    original = {i.subject: i.received for i, _b, _a in demo._items["f0"]}
    assert all(original[it.subject] == it.received for it in items if it.received)


def test_server_switches_connection(tmp_path):
    from archiver.demo_online import DemoOnlineBackend
    ctl = Controller({"classic": lambda s: DemoBackend(), "online": lambda s: DemoOnlineBackend()},
                     Settings(str(tmp_path / "s.json")), History(str(tmp_path / "h.sqlite3")))

    def wait():
        for _ in range(200):
            if not ctl.busy:
                return ctl.state()
            time.sleep(0.02)

    ctl.connect()
    st = wait()
    assert st["engine"] == "classic" and st["capabilities"]["pst"]
    ctl.connect("online")
    st = wait()
    assert st["engine"] == "online" and not st["capabilities"]["pst"] and "eml" in st["capabilities"]["formats"]
    assert st["account"] == "manager@example.com"
    assert ctl.settings.load()["engine"] == "online"
    ctl.sign_out()
    assert not ctl.state()["connected"] and ctl.settings.load()["engine"] == "auto"


def test_classic_missing_offers_online(tmp_path):
    from archiver.demo_online import DemoOnlineBackend
    ctl = Controller({"classic": lambda s: NoOutlook(), "online": lambda s: DemoOnlineBackend()},
                     Settings(str(tmp_path / "s.json")), History(str(tmp_path / "h.sqlite3")))
    ctl.connect()
    for _ in range(100):
        if not ctl.busy:
            break
        time.sleep(0.02)
    st = ctl.state()
    assert st["classic_failed"] and "Microsoft 365 online" in st["error"] and not st["connected"]


# ------------------------------------- classic Outlook next to the new Outlook

class FakeRegistry:
    HKCU, HKLM = "HKCU", "HKLM"

    def __init__(self, values=None, keys=(), subkeys=()):
        self.values = dict(values or {})
        self.keys, self.subkeys = set(keys), set(subkeys)
        self.writes = []

    def get_dword(self, root, path, name):
        return self.values.get((root, path, name))

    def set_dword(self, root, path, name, value):
        self.writes.append(value)
        self.values[(root, path, name)] = value

    def key_exists(self, root, path):
        return (root, path) in self.keys

    def has_subkeys(self, root, path):
        return (root, path) in self.subkeys


def test_status_reads_the_new_outlook_switch():
    from archiver import outlook_setup as o
    reg = FakeRegistry({("HKCU", o.PREFS, o.SWITCH): 1}, keys={("HKLM", o.APP_PATH)},
                       subkeys={("HKCU", o.PROFILES)})
    assert o.status(reg) == {"classic_installed": True, "new_outlook_on": True, "has_profile": True}
    assert o.status(FakeRegistry())["new_outlook_on"] is False


def test_switch_is_put_back_even_if_starting_fails():
    from archiver import outlook_setup as o
    reg = FakeRegistry({("HKCU", o.PREFS, o.SWITCH): 1})
    with pytest.raises(RuntimeError):
        with o.classic_can_start(reg) as changed:
            assert changed and reg.get_dword("HKCU", o.PREFS, o.SWITCH) == 0
            raise RuntimeError("classic Outlook did not start")
    assert reg.get_dword("HKCU", o.PREFS, o.SWITCH) == 1 and reg.writes == [0, 1]
    off = FakeRegistry({("HKCU", o.PREFS, o.SWITCH): 0})
    with o.classic_can_start(off) as changed:
        assert not changed
    assert off.writes == []  # nothing touched when the switch is already off


def test_background_mode_is_reported_and_outlook_closed(tmp_path):
    closed = []

    class Background(DemoBackend):
        def connect(self):
            self.background = True
            return "ok"

        def close(self):
            closed.append(True)

    ctl = Controller(lambda: Background(), Settings(str(tmp_path / "s.json")), History(str(tmp_path / "h.sqlite3")))
    ctl.connect()
    for _ in range(100):
        if not ctl.busy:
            break
        time.sleep(0.02)
    st = ctl.state()
    assert st["connected"] and st["background"] and st["capabilities"]["pst"]
    ctl.shutdown(wait=5)
    assert closed == [True]


@pytest.mark.skipif(os.name != "nt", reason="Windows registry")
def test_switch_handling_on_the_real_windows_registry():
    """Runs on the Windows build machine: set the switch, start, check it is put back."""
    from archiver import outlook_setup as o
    reg = o.Registry()
    before = reg.get_dword(reg.HKCU, o.PREFS, o.SWITCH)
    try:
        reg.set_dword(reg.HKCU, o.PREFS, o.SWITCH, 1)
        assert o.status(reg)["new_outlook_on"] is True
        with o.classic_can_start(reg) as changed:
            assert changed and reg.get_dword(reg.HKCU, o.PREFS, o.SWITCH) == 0
        assert reg.get_dword(reg.HKCU, o.PREFS, o.SWITCH) == 1
    finally:
        reg.set_dword(reg.HKCU, o.PREFS, o.SWITCH, before if before is not None else 0)
