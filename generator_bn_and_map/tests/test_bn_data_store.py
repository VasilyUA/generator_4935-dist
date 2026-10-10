import json
import os

import pytest

from bn_fixtures import PBD_PARAGRAPHS, write_docx, write_rop_vop
from content.data_store import build_pbd_cache, load_data, normalize_name, save_data
from content.pbd_reader import read_pbd
from content.rop_vop_reader import read_rop_vop


def test_load_missing_file(tmp_path):
    assert load_data(str(tmp_path / "немає.json")) == {}


def test_save_keeps_backup_and_unicode(tmp_path):
    path = tmp_path / "data.json"
    save_data(str(path), {"unit": {"A": "перше"}})
    save_data(str(path), {"unit": {"A": "друге"}})
    assert json.loads(path.read_text(encoding="utf-8")) == {"unit": {"A": "друге"}}
    assert json.loads((tmp_path / "data.json.bak").read_text(encoding="utf-8")) == {"unit": {"A": "перше"}}
    assert "друге" in path.read_text(encoding="utf-8")
    assert sorted(os.listdir(tmp_path)) == ["data.json", "data.json.bak"]


def test_save_failure_leaves_original_and_no_temp(tmp_path, monkeypatch):
    path = tmp_path / "data.json"
    save_data(str(path), {"a": 1})

    def broken_dump(*args, **kwargs):
        raise RuntimeError("збій запису")

    monkeypatch.setattr("content.data_store.json.dump", broken_dump)
    with pytest.raises(RuntimeError):
        save_data(str(path), {"a": 2})
    assert json.loads(path.read_text(encoding="utf-8")) == {"a": 1}
    assert not [name for name in os.listdir(tmp_path) if name.startswith(".data-")]


def test_normalize_name():
    assert normalize_name(" «Альфа-1» ") == "АЛЬФА1"
    assert normalize_name(None) == ""


def _cache(tmp_path, previous=None, with_table=True):
    report = read_pbd(str(write_docx(tmp_path / "ПБД.docx", PBD_PARAGRAPHS)))
    table = read_rop_vop(str(write_rop_vop(tmp_path / "РОП_ВОП.xlsx"))) if with_table else None
    return build_pbd_cache(report, table, previous, "09.10.2026 тест")


def test_cache_merges_types_quantities_and_reports_differences(tmp_path):
    cache, warnings = _cache(tmp_path)
    points = {(p["type"], p["name"]): p for p in cache["points"]}
    # Тип з РОП_ВОП точніший ("Вогнева засідка"), кількість - теж звідти.
    assert points[("ВЗ", "ЧАРЛІ")]["quantity"] == 3
    assert ("ВП", "ЧАРЛІ") not in points
    assert points[("КСП", "АЛЬФА")]["quantity"] == 4
    # Перейменований хаб знайдено за координатами.
    assert points[("ХАБ", "ГОТЕЛЬ")]["quantity"] == 7
    assert any("ГОТЕЛЬ" in w and "ГОТЕЛІ" in w for w in warnings)
    # Позиція лише з РОП_ВОП додається з попередженням; ЗКП - мовчки.
    assert points[("СП", "ІНДІЯ")]["group"] == "РОП_ВОП"
    assert any("ІНДІЯ" in w for w in warnings)
    assert points[("ЗКП", "")]["mgrs"] == "31U DQ 40800 10800"
    assert any("ГОЛЬФ" in w and "немає в РОП_ВОП" in w for w in warnings)
    assert cache["zkp_mgrs"] == "31U DQ 40800 10800"
    assert cache["report_datetime"] == "2026-10-09T20:00:00"
    assert cache["folder"] == "09.10.2026 тест"
    assert cache["source"] == "ПБД.docx" and cache["rop_vop_source"] == "РОП_ВОП.xlsx"


def test_cache_default_line_visibility(tmp_path):
    cache, _ = _cache(tmp_path)
    on_map = {line["kind"]: line["on_map"] for line in cache["lines"]}
    assert on_map == {"bro": True, "flot": True, "boundary_right": True, "boundary_left": True,
                      "rop": True, "flot_rop": False, "vop": False, "flot_vop": False}


def test_cache_preserves_manual_edits(tmp_path):
    previous = {
        "points": [{"type": "ТЗ", "name": "ФОКСТРОТ", "on_map": False, "sidc": "10031000001100000000"}],
        "lines": [{"kind": "vop", "ref": "ВОП «БРАВО»", "on_map": True, "name": "Контур"}],
    }
    cache, _ = _cache(tmp_path, previous)
    point = next(p for p in cache["points"] if p["name"] == "ФОКСТРОТ")
    assert point["on_map"] is False and point["sidc"] == "10031000001100000000"
    line = next(l for l in cache["lines"] if l["kind"] == "vop")
    assert line["on_map"] is True and line["name"] == "Контур"


def test_cache_without_table(tmp_path):
    cache, warnings = _cache(tmp_path, with_table=False)
    assert all(p["quantity"] is None for p in cache["points"])
    assert cache["zkp_mgrs"] == "" and cache["rop_vop_source"] == ""
    assert warnings == []


def test_cache_keeps_reserve_areas(tmp_path):
    cache, _ = _cache(tmp_path)
    assert cache["reserve_areas"] == {"БнГ": "лісосмуга (31U DQ 40700 10700)"}
    assert "signature" in cache["blocks"]
