"""data.json (ручне) + data_auto.json (веде програма), компактний JSON і
штабні рішення, що оновлюються з карти Дельти - синтетичні дані."""
import json
import os

from bn_fixtures import make_data
from content.data_store import load_data, merge_data, migrate_layout, save_auto_data, save_data, split_data
from content.staff import StaffEntry, staff_to_data
from utils.json_format import dumps_compact

A, B = "31U DQ 40100 10100", "31U DQ 40200 10200"


def test_dumps_compact_keeps_short_containers_on_one_line():
    value = {"title": "Межа", "mgrs": [A, B], "long": ["x" * 60, "y" * 60], "empty": [], "n": 1}
    text = dumps_compact(value, width=80)
    lines = text.split("\n")
    assert '  "mgrs": ["31U DQ 40100 10100", "31U DQ 40200 10200"],' in lines
    assert '  "long": [' in lines and lines[-1] == "}"
    assert json.loads(text) == value
    assert dumps_compact({"a": 1}) == '{"a": 1}' and dumps_compact("рядок") == '"рядок"'
    assert all(len(line) <= 80 for line in lines if '"x' not in line and '"y' not in line)


def test_split_and_merge_roundtrip():
    data = make_data()
    data["bn"].update({"KSP_MGRS": A, "ZKP_MGRS": B})
    data["map"]["STAFF_SIDC"] = {"межі": ["1", "межа"]}
    data["staff"] = {"межі": {"межа-1": {"title": "Межа", "mgrs": [A]}}}
    data["pbd"] = {"points": []}
    data["власний_ключ"] = {"x": 1}
    manual, auto = split_data(data)
    assert set(manual) == {"unit", "bn", "map", "власний_ключ"}
    assert manual["bn"] == {"NUMBER": "5", "TIME": "06.00"}
    assert set(manual["map"]) == {"STAFF_COMMENT", "RELIABILITY", "PLATFORM_TYPE"}
    assert list(auto) == ["previous", "map", "staff", "bn_document", "pbd"]
    assert auto["previous"] == {"KSP_MGRS": A, "ZKP_MGRS": B} and "SIDC_BY_TYPE" in auto["map"]
    assert merge_data(manual, auto) == data
    # Ручне значення важливіше за автоматичне, кеш ПБД - завжди з data_auto.json.
    merged = merge_data({"map": {"LINE_COLOR": "#123456"}, "pbd": {"old": 1}}, {"map": {"LINE_COLOR": "#000000"}, "pbd": {"new": 1}})
    assert merged == {"map": {"LINE_COLOR": "#123456"}, "pbd": {"new": 1}}
    assert merge_data({}, {}) == {}


def test_load_and_save_two_files(tmp_path):
    data = make_data()
    data["pbd"] = {"points": [{"type": "ВП", "name": "АЛЬФА", "mgrs": A}]}
    manual, auto = split_data(data)
    save_data(str(tmp_path / "data.json"), manual)
    save_auto_data(str(tmp_path / "data_auto.json"), data)
    assert load_data(str(tmp_path / "data.json"), str(tmp_path / "data_auto.json")) == data
    assert load_data(str(tmp_path / "data.json")) == manual
    assert load_data(str(tmp_path / "немає.json"), str(tmp_path / "data_auto.json")) == {}
    on_disk = json.loads((tmp_path / "data_auto.json").read_text(encoding="utf-8"))
    assert "unit" not in on_disk and on_disk["pbd"] == data["pbd"]


def test_migrate_layout_splits_old_single_file(tmp_path):
    data = make_data()
    data["pbd"] = {"points": []}
    path, auto_path = str(tmp_path / "data.json"), str(tmp_path / "data_auto.json")
    save_data(path, data)
    assert migrate_layout(path, auto_path) is True
    manual = json.loads((tmp_path / "data.json").read_text(encoding="utf-8"))
    assert set(manual) == {"unit", "bn", "map"} and "document" not in manual["bn"]
    assert load_data(path, auto_path) == data
    assert migrate_layout(path, auto_path) is False          # уже розділено
    assert sorted(os.listdir(tmp_path)) == ["data.json", "data_auto.json"]


def test_staff_to_data_keeps_other_fields():
    original = {"межі": {"м-1": {"title": "Межа", "text": "дорога ({1})", "mgrs": [A], "on_map": False}},
                "не_словник": "рядок"}
    entries = [StaffEntry("межі", "м-1", title="Межа", text="", mgrs=[B]),
               StaffEntry("точки", "т-1", title="Точка", mgrs=[A]),
               StaffEntry("не_словник", "х", mgrs=[A])]
    result = staff_to_data(entries, original)
    assert result["межі"]["м-1"] == {"title": "Межа", "text": "", "mgrs": [B], "on_map": False}
    assert result["точки"]["т-1"] == {"title": "Точка", "mgrs": [A]}
    assert result["не_словник"] == "рядок"
    assert original["межі"]["м-1"]["mgrs"] == [A]           # оригінал не змінено
