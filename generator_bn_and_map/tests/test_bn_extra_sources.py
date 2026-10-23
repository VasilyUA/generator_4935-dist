"""Нові джерела БН (синтетичні дані): тека ПБД/, позатермінові донесення з
перейменуваннями, розпорядження з розвідувального забезпечення (РР), таблиця
сигналів, пошук експорту Дельти."""
import os
from datetime import date

import pytest
from docx import Document

import generators.pipeline as pipeline
from bn_fixtures import (
    PBD_PARAGRAPHS, RENAME_PARAGRAPHS, RR_PARAGRAPHS, SIGNAL_ROWS, write_docx, write_rop_vop, write_signal_table,
)
from content.position_refresh import PositionRefresher, build_rename_map, is_position_gap
from content.rename_reader import Rename, parse_renames, position_type, read_renames
from content.rr_reader import read_rr
from content.signal_table import read_signal_table, signal_key, split_signal
from content.sources import (
    date_in_name, find_all_by_glob, find_delta_export, find_latest_by_glob, find_latest_pbd, find_rop_vop,
)


# --- пошук файлів --------------------------------------------------------------
@pytest.fixture
def resources(tmp_path):
    res = tmp_path / "resources"
    pbd = res / "ПБД"
    old = res / "01.10.2026 стара"
    for folder in (pbd, old):
        folder.mkdir(parents=True)
    write_docx(old / "ПБД 01.10.2026.docx", PBD_PARAGRAPHS)
    write_docx(pbd / "Додаток №1 ПБД 08.10.2026.docx", PBD_PARAGRAPHS)
    write_docx(pbd / "Додаток №1 ПБД 09.10.2026.docx", PBD_PARAGRAPHS)
    write_rop_vop(pbd / "Додаток 2 (РОП_ВОП) 08.10.2026.xlsx")
    write_rop_vop(pbd / "Додаток 2 (РОП_ВОП) 09.10.2026.xlsx")
    (pbd / "~$Додаток №1 ПБД 10.10.2026.docx").write_bytes(b"")
    return res


def test_date_in_name():
    assert date_in_name("Додаток ПБД 01.02.2026 і 09.10.2026.docx") == date(2026, 10, 9)
    assert date_in_name("без дати.docx") is None
    assert date_in_name("31.02.2026.docx") is None


def test_latest_pbd_from_pbd_folder(resources):
    path, folder = find_latest_pbd(str(resources))
    assert os.path.basename(path) == "Додаток №1 ПБД 09.10.2026.docx" and os.path.basename(folder) == "ПБД"
    rop_vop = find_rop_vop(folder, ("РОП", "ВОП"), path)
    assert os.path.basename(rop_vop) == "Додаток 2 (РОП_ВОП) 09.10.2026.xlsx"
    assert find_rop_vop(None, ("РОП",)) is None
    assert find_latest_pbd(str(resources / "немає")) == (None, None)


def test_refresh_pbd_cache_uses_pbd_folder(resources):
    cache, _warnings, used = pipeline.refresh_pbd_cache({}, str(resources))
    assert [os.path.basename(p) for p in used] == ["Додаток №1 ПБД 09.10.2026.docx", "Додаток 2 (РОП_ВОП) 09.10.2026.xlsx"]
    assert cache["folder"] == "ПБД" and cache["points"]


def test_refresh_error_mentions_pbd_folder(tmp_path):
    with pytest.raises(pipeline.SourceError, match="ПБД"):
        pipeline.refresh_pbd_cache({}, str(tmp_path))


def test_latest_and_all_by_glob(tmp_path):
    for name in ("РР №1 від 01.08.2026.docx", "РР №2 від 17.08.2026.docx", "РР без дати.docx"):
        (tmp_path / name).write_bytes(b"x")
    assert os.path.basename(find_latest_by_glob(str(tmp_path), "РР*.docx")) == "РР №2 від 17.08.2026.docx"
    assert find_latest_by_glob(str(tmp_path), "Немає*.docx") is None
    assert len(find_all_by_glob(str(tmp_path), "РР*.docx")) == 3


def test_find_delta_export(tmp_path):
    assert find_delta_export(str(tmp_path)) is None
    folder = tmp_path / "Дельта"
    folder.mkdir()
    older = folder / "експорт.csv"
    newer = tmp_path / "БП - дельта.csv"
    older.write_text("x", encoding="utf-8")
    newer.write_text("x", encoding="utf-8")
    (tmp_path / "інший.csv").write_text("x", encoding="utf-8")
    os.utime(older, (1_000_000, 1_000_000))
    assert find_delta_export(str(tmp_path)) == str(newer)
    os.utime(older, (4_000_000_000, 4_000_000_000))
    assert find_delta_export(str(tmp_path)) == str(older)


# --- перейменування -------------------------------------------------------------
def test_parse_renames_and_types():
    renames = parse_renames(RENAME_PARAGRAPHS)
    assert renames == [Rename("ВП", "СТАРА", "НОВА"), Rename("ПВ", "БЕТА", "ГАММА"), Rename("ВОП", "ОМЕГА", "СИГМА")]
    assert position_type("Спостережна позиція") == "СП" and position_type("Вогнева засідка") == "ВЗ"
    assert position_type("щось інше") == ""
    assert parse_renames(["ВП «ТА САМА» перейменовано на «та сама»"]) == []


def test_read_renames(tmp_path):
    report = read_renames(write_docx(tmp_path / "ДОНЕСЕННЯ_ПОЗАТЕРМІНОВЕ 17.docx", RENAME_PARAGRAPHS))
    assert (report.number, report.date, report.label) == ("17", "11.08.2026", "№17 від 11.08.2026")
    assert len(report.renames) == 3


def test_rename_map_follows_chains():
    mapping = build_rename_map([Rename("ВП", "А1", "Б1"), Rename("ВП", "Б1", "В1")])
    assert mapping == {"А1": "В1", "Б1": "В1"}


def test_refresher_applies_reported_rename_then_coordinates():
    refresher = PositionRefresher([{"type": "ПВ", "name": "ГАММА", "mgrs": "31U DQ 40300 10300"}],
                                  renames=[Rename("ПВ", "БЕТА", "ГАММА")])
    result = refresher.refresh("ПВ «БЕТА» за координатами (31U DQ 49999 19999).")
    assert result.text == "ПВ «ГАММА» за координатами (31U DQ 40300 10300)."
    assert result.reported == [("ПВ «БЕТА»", "ПВ «ГАММА»")] and result.renamed == []


def test_refresher_renames_strongpoints_but_not_current_names():
    renames = [Rename("ВОП", "ОМЕГА", "СИГМА")]
    result = PositionRefresher([], renames=renames).refresh("зайняти ротний опорний пункт «ОМЕГА»")
    assert result.text == "зайняти ротний опорний пункт «СИГМА»"
    assert result.reported == [("РОП «ОМЕГА»", "РОП «СИГМА»")]
    kept = PositionRefresher([], renames=renames, current_names=["ОМЕГА"]).refresh("ВОП «ОМЕГА»")
    assert kept.text == "ВОП «ОМЕГА»" and kept.reported == []


def test_reported_name_missing_in_pbd_is_unknown():
    refresher = PositionRefresher([{"type": "ВП", "name": "ІНША", "mgrs": "31U DQ 40100 10100"}],
                                  renames=[Rename("ВП", "СТАРА", "НОВА")])
    result = refresher.refresh("ВП «СТАРА» (31U DQ 45000 15000)")
    assert result.text == "ВП «НОВА» (31U DQ 45000 15000)" and result.unknown == ["ВП «НОВА»"]


def test_coordinates_of_a_new_place_are_not_the_position():
    assert is_position_gap(" за координатами (") and is_position_gap(" мати околиці н.п. «ТЕСТОВЕ» за координатами (")
    assert not is_position_gap(" на проміжний рубіж за координатами (")
    refresher = PositionRefresher([{"type": "ВП", "name": "АЛЬФА", "mgrs": "31U DQ 40100 10100"}])
    text = "перемістити ВП «АЛЬФА» на проміжний рубіж за координатами (31U DQ 41111 11111; 31U DQ 42222 12222)"
    assert refresher.refresh(text).text == text


# --- РР ---------------------------------------------------------------------------
def test_read_rr(tmp_path):
    order = read_rr(write_docx(tmp_path / "РР 77.docx", RR_PARAGRAPHS))
    assert (order.number, order.date, order.warnings) == ("77", "17.08.2026", [])
    sizes = {name: len(items) for name, items in order.blocks.items()}
    assert sizes == {
        "rr_stage:1:all": 6, "rr_stage:1:zone": 1, "rr_stage:1:detailed": 1, "rr_stage:1:rou": 1, "rr_stage:1:rpr": 1,
        "rr_stage:2:all": 3, "rr_stage:2:zone": 1, "rr_stage:2:rpr": 1,
        "rr_stage:1:tasks": 1, "rr_stage:2:tasks": 1,
    }
    assert order.blocks["rr_stage:1:rou"] == ["РОУ №1/1 – ЧАРЛІ (31U DQ 40300 10300);"]
    zone = order.blocks["rr_stage:2:zone"][0]
    assert zone.startswith("Зону відповідальності") and "ЕХО (31U DQ 40500 10500)" in zone


def test_read_rr_without_zones(tmp_path):
    order = read_rr(write_docx(tmp_path / "РР.docx", ["РОЗПОРЯДЖЕННЯ № 5 від 01.01.2026", "Текст."]))
    assert order.number == "5" and order.blocks == {} and order.warnings


# --- таблиця сигналів -------------------------------------------------------------
def test_signal_helpers():
    assert signal_key("‘‘ГРІМ – 101’’") == "ГРІМ101"
    assert split_signal("ГРІМ-101") == ("ГРІМ", "101")
    assert split_signal("ДВА СЛОВА 4444") == ("ДВА СЛОВА", "4444")
    assert split_signal("БЕЗЦИФР") == ("БЕЗЦИФР", "")


def test_read_signal_table(tmp_path):
    table = read_signal_table(write_signal_table(tmp_path / "Таблиця сигналів.docx", SIGNAL_ROWS))
    assert [row.name for row in table.rows] == ["ГРІМ-101", "ХМАРА-201", "ХМАРА-202", "ХМАРА-221", "ЗОРЯ-301", "ЗАМОК-41"]
    by_name = {row.name: row for row in table.rows}
    assert by_name["ГРІМ-101"].section == "1. ЗАГАЛЬНІ СИГНАЛИ"
    assert (by_name["ХМАРА-221"].group, by_name["ХМАРА-221"].line) == ("3.1. Для ЗВРез", "На проміжному рубежі")
    assert by_name["ЗОРЯ-301"].group == "3.2. Для БнГ" and by_name["ЗОРЯ-301"].line == ""
    assert table.find("ГРІМ101") is by_name["ГРІМ-101"] and table.find("НЕМАЄ1") is None
    assert "ЗАМОК41" in table.keys


def test_read_document_without_signal_table(tmp_path):
    path = tmp_path / "порожня.docx"
    Document().save(path)
    assert read_signal_table(str(path)).warnings
