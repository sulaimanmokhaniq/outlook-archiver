import csv
import datetime as dt
import os

import pytest

from archiver.core import (
    Criteria,
    ExportOptions,
    ItemInfo,
    PathAllocator,
    classify,
    clean_date,
    file_stem,
    index_row,
    matches_deep,
    matches_quick,
    parse_date,
    relative_dir,
    safe_name,
)
from archiver.demo_backend import DemoBackend


def item(**kw) -> ItemInfo:
    base = dict(entry_id="1", store_id="s", folder_key="f0", folder_parts=("Inbox",),
                store_name="Me", message_class="IPM.Note", subject="Budget",
                sender="Ahmed Ali", recipients="Sara", categories="",
                received=dt.datetime(2021, 5, 3, 9, 30))
    base.update(kw)
    return ItemInfo(**base)


@pytest.mark.parametrize("mc,kind", [
    ("IPM.Note", "mail"),
    ("IPM.Note.SMIME", "mail"),
    ("IPM.Schedule.Meeting.Request", "meeting"),
    ("IPM.Appointment", "appointment"),
    ("IPM.Contact", "contact"),
    ("IPM.DistList", "contact"),
    ("IPM.Task", "task"),
    ("IPM.TaskRequest.Accept", "task"),
    ("IPM.StickyNote", "note"),
    ("REPORT.IPM.Note.NDR", "other"),
    ("", "other"),
])
def test_classify(mc, kind):
    assert classify(mc) == kind


def test_clean_date_handles_outlook_no_date():
    assert clean_date(dt.datetime(4501, 1, 1)) is None
    assert clean_date(None) is None
    tz = dt.datetime(2020, 1, 2, 3, 4, 5, tzinfo=dt.timezone.utc)
    assert clean_date(tz) == dt.datetime(2020, 1, 2, 3, 4, 5)


def test_date_range_is_inclusive():
    c = Criteria(date_from=dt.date(2021, 5, 3), date_to=dt.date(2021, 5, 3))
    assert matches_quick(item(), c)
    c = Criteria(date_from=dt.date(2021, 5, 4))
    assert not matches_quick(item(), c)


def test_auto_date_uses_meeting_start_and_draft_fallback():
    appt = item(message_class="IPM.Appointment", received=None,
                created=dt.datetime(2019, 1, 1), start=dt.datetime(2022, 6, 1))
    assert appt.display_date == dt.datetime(2022, 6, 1)
    draft = item(received=None, modified=dt.datetime(2020, 2, 2))
    assert draft.display_date == dt.datetime(2020, 2, 2)


def test_undated_items_excluded_when_dates_set():
    c = Criteria(date_from=dt.date(2000, 1, 1))
    assert not matches_quick(item(received=None), c)
    c.include_undated = True
    assert matches_quick(item(received=None), c)


def test_text_filters_any_term_case_insensitive():
    assert matches_quick(item(), Criteria(sender="sara; AHMED"))
    assert not matches_quick(item(), Criteria(sender="omar"))
    assert matches_quick(item(), Criteria(subject="budg"))
    assert not matches_quick(item(), Criteria(kinds={"appointment"}))


def test_deep_filters():
    c = Criteria(body_words="contract", attachments="with")
    assert c.needs_item_open
    assert matches_deep("x", "the CONTRACT is attached", 1, c)
    assert not matches_deep("x", "the contract", 0, c)
    assert not matches_deep("x", "nothing", 2, c)


def test_parse_date_formats():
    assert parse_date("2024-03-05") == dt.date(2024, 3, 5)
    assert parse_date("05/03/2024") == dt.date(2024, 3, 5)
    assert parse_date("") is None
    with pytest.raises(ValueError):
        parse_date("March")


def test_safe_names():
    assert safe_name('RE: a/b "c"?') == "RE_ a_b _c__"
    assert safe_name("") == "(no subject)"
    assert safe_name("con") == "_con"
    assert len(safe_name("x" * 500)) == 80


def test_layouts_and_stem():
    it = item(folder_parts=("Inbox", "A/B"))
    assert relative_dir(it, "folder_year", False) == ["Inbox", "A_B", "2021"]
    assert relative_dir(it, "folder_year", True)[0] == "Me"
    assert relative_dir(it, "type_year", False) == ["Emails", "2021"]
    assert relative_dir(it, "year_month", False) == ["2021", "05"]
    assert file_stem(it) == "2021-05-03 0930 - Budget"


def test_path_allocator_unique_and_short(tmp_path):
    alloc = PathAllocator()
    a = alloc.allocate(str(tmp_path), "same", ".msg")
    b = alloc.allocate(str(tmp_path), "same", ".msg")
    assert a != b and b.endswith("same (2).msg")
    long = alloc.allocate(str(tmp_path), "y" * 400, ".msg")
    assert len(long) <= 250


def test_index_row_blocks_excel_formulas():
    row = index_row(item(subject="=HYPERLINK(\"x\")"), [], "/")
    assert row[5].startswith("'=")


def test_options_validation(tmp_path):
    with pytest.raises(ValueError):
        ExportOptions(output_dir=str(tmp_path / "missing"), archive_name="a").validate()
    with pytest.raises(ValueError):
        ExportOptions(output_dir=str(tmp_path), archive_name="a", make_pst=False,
                      write_index=False).validate()


def _noop(*_a):
    return None


def test_demo_end_to_end(tmp_path):
    b = DemoBackend()
    b.connect()
    tree = b.list_tree()
    keys = [n.key for n in tree[0].children]
    keys += [c.key for n in tree[0].children for c in n.children]
    crit = Criteria(kinds={"mail", "meeting"}, date_from=dt.date(2019, 1, 1),
                    date_to=dt.date(2019, 12, 31))
    found = b.scan(keys, crit, _noop, lambda: False)
    assert found
    assert all(it.kind in ("mail", "meeting") for it in found)
    assert all(it.display_date.year == 2019 for it in found)
    opts = ExportOptions(output_dir=str(tmp_path), archive_name="Test 2019",
                         file_formats={"mht", "txt"}, save_attachments=True)
    res = b.export(found, opts, _noop, lambda: False)
    assert not res.errors
    assert res.total == len(found)
    assert os.path.exists(os.path.join(res.folder, "README.txt"))
    with open(os.path.join(res.folder, "Archive list.csv"), encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    assert len(rows) == len(found) + 1
    assert os.path.isdir(os.path.join(res.folder, "Files", "Inbox", "2019"))
    # second run with the same name must not overwrite the first
    res2 = b.export(found[:3], opts, _noop, lambda: False)
    assert res2.folder != res.folder
