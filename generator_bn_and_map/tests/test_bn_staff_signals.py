"""Штабні рішення (data_auto.json → staff), виправлена таблиця сигналів і їх
використання в DocumentBuilder - синтетичні дані."""
import os

from content.delta_reader import DeltaExport, DeltaObject
from content.signal_table import read_signal_table, signal_key
from content.staff import (
    SOURCE_DELTA, StaffEntry, apply_delta, document_placeholders, duplicate_coordinates, invalid_coordinates,
    load_staff, staff_values,
)
from bn_fixtures import SIGNAL_ROWS, write_signal_table
from generators.document_builder import (
    ORIGIN_MANUAL, ORIGIN_REFRESHED, ORIGIN_STAFF, ORIGIN_STANDARD, DocumentBuilder, SignalMention,
)
from generators.generate_signals import find_duplicates, signal_meaning, write_corrected_table
from utils.mgrs import latlon_to_mgrs, mgrs_to_latlon

A, B, C = "31U DQ 40100 10100", "31U DQ 40200 10200", "31U DQ 40300 10300"
STAFF_DATA = {"staff": {
    "межі": {"межа-1": {"title": "Межа A0002 – A0003", "text": "дорога ({1}), ліс ({2})", "mgrs": [A, B]}},
    "точки": {"т-1": {"title": "Точка збору", "mgrs": C}, "порожня": {"title": "Стабпункт", "mgrs": []}},
    "не_словник": "рядок",
    "погані": {"т-2": "рядок", "т-3": {"title": "Погана", "mgrs": ["не координата"]}},
}}


def _entries():
    return {entry.placeholder: entry for entry in load_staff(STAFF_DATA)}


def _mention(name, text, unit=""):
    return SignalMention(name, signal_key(name), text, unit, "", False)


# --- штабні рішення ----------------------------------------------------------------
def test_load_and_render_staff():
    entries = _entries()
    assert set(entries) == {"staff:межі:межа-1", "staff:точки:т-1", "staff:точки:порожня", "staff:погані:т-3"}
    boundary = entries["staff:межі:межа-1"]
    assert boundary.render() == f"дорога ({A}), ліс ({B})" and boundary.shape == "line"
    point = entries["staff:точки:т-1"]
    assert point.render() == f"({C})" and point.shape == "point" and point.mgrs == [C]
    assert StaffEntry("г", "к", text="лише текст").render() == "лише текст"
    assert StaffEntry("г", "к", text="({1}) ({3})", mgrs=[A]).render() == f"({A}) ({{3}})"
    assert StaffEntry("г", "к", mgrs=[A, B, C], geometry="area").shape == "area"
    assert load_staff({}) == [] and load_staff(None) == []


def test_invalid_and_duplicate_coordinates():
    entries = list(_entries().values())
    assert invalid_coordinates(entries) == ["staff:погані:т-3: не MGRS - не координата"]
    same = [StaffEntry("вр", "1", mgrs=[A, B]), StaffEntry("вр", "2", mgrs=[A, B]),
            StaffEntry("межі", "1", mgrs=[A, B]), StaffEntry("межі", "2", mgrs=[A, B])]
    assert duplicate_coordinates(same) == [["staff:вр:1", "staff:вр:2"]]


def _delta(name, coords, geometry="LINESTRING"):
    return DeltaExport("x.csv", objects=[DeltaObject({}, "", "", name, geometry, coords)])


def test_apply_delta_keeps_descriptions_when_point_count_matches():
    entries = load_staff(STAFF_DATA)
    coords = [mgrs_to_latlon("31U DQ 41000 11000"), mgrs_to_latlon("31U DQ 42000 12000")]
    notes = apply_delta(entries, _delta("межа a0002 - a0003", coords))
    boundary = next(e for e in entries if e.key == "межа-1")
    assert notes == [] and boundary.source == SOURCE_DELTA
    assert boundary.render() == "дорога (31U DQ 41000 11000), ліс (31U DQ 42000 12000)"
    assert next(e for e in entries if e.key == "т-1").source == "data_auto.json"
    assert apply_delta(entries, None) == []


def test_apply_delta_with_other_point_count_and_closed_area():
    entry = StaffEntry("райони", "р", title="Район", text="({1}), ({2})", mgrs=[A, B], geometry="area")
    ring = [mgrs_to_latlon(m) for m in ("31U DQ 41000 11000", "31U DQ 42000 11000", "31U DQ 42000 12000")]
    notes = apply_delta([entry], _delta("Район", ring + ring[:1], "POLYGON"))
    assert len(entry.mgrs) == 3 and entry.text == "" and len(notes) == 1
    assert entry.render() == "(31U DQ 41000 11000), (31U DQ 42000 11000), (31U DQ 42000 12000)"


def test_document_placeholders_are_staff_decisions_of_the_order():
    # На карту йдуть лише штабні рішення, плейсхолдери яких є в тексті наказу.
    data = {"bn": {"document": {
        "header": ["КСП – {ksp_place}"],
        "sections": [{"heading": "Завдання", "content": [
            "ВР №1 - {staff:вогневі_рубежі:вр-1}",
            {"text": "межа {staff:межі:межа-1}, точка {staff:точки:т-1}", "bold": ["межа"]},
            [{"block": "pbd:positions"}, "район {staff:райони:р-1}"],
        ]}],
    }}}
    assert document_placeholders(data) == {
        "staff:вогневі_рубежі:вр-1", "staff:межі:межа-1", "staff:точки:т-1", "staff:райони:р-1"}
    assert document_placeholders({}) == set()


def test_latlon_to_mgrs_matches_staff_format():
    assert latlon_to_mgrs(*mgrs_to_latlon(A)) == A
    assert latlon_to_mgrs(*mgrs_to_latlon(A), digits=3) == "31U DQ 401 101"


# --- DocumentBuilder ----------------------------------------------------------------
def test_builder_renders_staff_placeholders_and_skips_empty_ones():
    entries = load_staff(STAFF_DATA)
    builder = DocumentBuilder({}, {}, staff=staff_values(entries))
    doc = builder.build({"sections": [
        "праворуч: {staff:межі:межа-1}, відповідальний сусід.",
        "- стабпункт за координатами {staff:точки:порожня};",
        "Точка: {staff:точки:т-1}.",
    ]})
    texts = [p.text for p in doc.paragraphs]
    assert texts == [f"праворуч: дорога ({A}), ліс ({B}), відповідальний сусід.", f"Точка: ({C})."]
    assert [entry.origin for entry in builder.log] == [ORIGIN_STAFF, ORIGIN_STAFF]
    assert builder.staff_paragraphs()[0].structured_values == (A, B)
    assert builder.manual_paragraphs() == []
    assert set(builder.used_staff) == {"staff:межі:межа-1", "staff:точки:т-1"}
    assert set(builder.skipped_staff) == {"staff:точки:порожня"}
    staff_chars = len(A) + len(B) + len(C)
    all_chars = sum(len(entry.text) for entry in builder.log)
    assert builder.automatic_percent() == round(100 - 100 * staff_chars / all_chars, 1)


def test_builder_treats_delta_coordinates_as_source():
    entries = load_staff(STAFF_DATA)
    apply_delta(entries, _delta("Межа A0002 – A0003", [mgrs_to_latlon(A), mgrs_to_latlon(B)]))
    builder = DocumentBuilder({}, {}, staff=staff_values(entries))
    builder.build({"sections": ["праворуч: {staff:межі:межа-1}."]})
    assert builder.log[0].origin == ORIGIN_REFRESHED and builder.automatic_percent() == 100.0


def test_builder_tracks_units_and_known_signals():
    builder = DocumentBuilder({}, {"порожній": []}, known_signals={"ГРІМ101"})
    builder.build({"sections": [
        "Командиру ЗВРез A0001:",
        "ВР №1 - дорога сигнал висування – «ГРОЗА-501»",
        "За сигналом «ГРІМ-101» відійти.",
        {"heading": "ІНШЕ", "content": ["За сигналом «ДОЩ-17» діяти.", {"block": "порожній", "optional": True}]},
    ]})
    mentions = {m.name: m for m in builder.signal_mentions}
    assert mentions["ГРОЗА-501"].unit == "Командиру ЗВРез A0001" and not mentions["ГРОЗА-501"].known
    assert mentions["ГРІМ-101"].known and mentions["ДОЩ-17"].unit == ""
    origins = [entry.origin for entry in builder.log if "сигнал" in entry.text.lower()]
    assert origins == [ORIGIN_MANUAL, ORIGIN_STANDARD, ORIGIN_MANUAL]
    assert builder.missing_blocks == []


# --- таблиця сигналів ----------------------------------------------------------------
def test_signal_meaning():
    text = ("За сигналом «ДОЩ-7» A0002 здійснює поступовий відхід на «С» + 01.20 під прикриттям вогню, "
            "здійснює відхід на проміжний рубіж оборони, із зайняттям опорного пункту.")
    assert signal_meaning(text, "ДОЩ-7", "Командиру A0002:") == "Здійснює відхід на проміжний рубіж оборони (командиру A0002)"
    assert signal_meaning("ВР №21 - дорога (31U DQ 40100 10100) сигнал – «ХМАРА-9»", "ХМАРА-9") == "Вийти на вогневий рубіж №21"
    plain = signal_meaning("Діяти за сигналом «ВІЯ-3» (31U DQ 40100 10100) самостійно", "ВІЯ-3")
    assert plain == "Діяти самостійно"
    timed = signal_meaning("з отриманням сигналу ‘‘ВІЯ-13’’ до ‘‘С’’ + 01.30 з дозволу старшого начальника.", "ВІЯ-13")
    assert timed == "З дозволу старшого начальника"
    moved = signal_meaning("За сигналом «ВІЯ-14» здійснити відхід позицій ВОП «БРАВО» за координатами (31U DQ 40100 10100).",
                           "ВІЯ-14")
    assert moved == "Здійснити відхід позицій ВОП «БРАВО»"


def test_corrected_signal_table(tmp_path):
    source = write_signal_table(tmp_path / "Таблиця сигналів A0001.docx", SIGNAL_ROWS)
    table = read_signal_table(str(source))
    unit = "Командиру ЗВРез командира A0001"
    mentions = [
        _mention("ГРОЗА-501", "ВР №1 - дорога (31U DQ 40100 10100) сигнал висування – «ГРОЗА-501»", unit),
        _mention("ГРОЗА-522", "ВР №22 - дорога (31U DQ 40200 10200) сигнал висування – «ГРОЗА-522»", unit),
        _mention("ГРІМ-101", "За сигналом «ГРІМ-101» перейти до оборони."),
        _mention("ДОЩ-7", "За сигналом «ДОЩ-7» A0002 здійснити відхід на проміжний рубіж, під прикриттям вогню.",
                 "Командиру A0002"),
        _mention("ДОЩ-7", "Повторна згадка: за сигналом «ДОЩ-7» A0002 здійснити відхід на проміжний рубіж."),
    ]
    output = tmp_path / "output"
    result = write_corrected_table(table, mentions, str(output), "бойового наказу №5 від 24.08.2026")
    assert result.path == os.path.join(str(output), "Таблиця сигналів A0001.docx")
    assert result.updated == [("3.1. Для ЗВРез", "Вийти на вогневий рубіж №1", "ХМАРА-201", "ГРОЗА-501")]
    assert [name for name, _ in result.added] == ["ГРОЗА-522", "ДОЩ-7"]

    corrected = read_signal_table(result.path)
    names = [row.name for row in corrected.rows]
    assert names == ["ГРІМ-101", "ГРОЗА-501", "ХМАРА-202", "ХМАРА-221", "ГРОЗА-522", "ЗОРЯ-301", "ЗАМОК-41", "ДОЩ-7"]
    by_name = {row.name: row for row in corrected.rows}
    assert by_name["ГРОЗА-522"].line == "На проміжному рубежі"
    assert by_name["ДОЩ-7"].section.startswith("6. СИГНАЛИ БОЙОВОГО НАКАЗУ №5 ВІД 24.08.2026")
    assert by_name["ДОЩ-7"].meaning == "Здійснити відхід на проміжний рубіж (командиру A0002)"
    # Оригінал у resources/ не змінюється.
    assert [row.name for row in read_signal_table(str(source)).rows][1] == "ХМАРА-201"


def test_table_without_changes_is_not_written(tmp_path):
    table = read_signal_table(write_signal_table(tmp_path / "Таблиця.docx", SIGNAL_ROWS))
    result = write_corrected_table(table, [_mention("ГРІМ-101", "За сигналом «ГРІМ-101» діяти.")], str(tmp_path / "out"), "x")
    assert result.path == "" and not (tmp_path / "out").exists()
    assert write_corrected_table(None, [], str(tmp_path), "x").path == ""


def test_find_duplicates():
    unit = "Командиру БнГ A0001"
    mentions = [
        _mention("ЗОРЯ-9", "ВР №3 - дорога сигнал – «ЗОРЯ-9»", unit),
        _mention("ЗОРЯ-9", "ВР №21 - дорога сигнал – «ЗОРЯ-9»", unit),
        _mention("ГРІМ-1", "ВР №1 - дорога сигнал – «ГРІМ-1»", unit),
        _mention("ГРІМ-1", "ВР №1 - дорога сигнал – «ГРІМ-1»", unit),
    ]
    assert find_duplicates(mentions) == [("ЗОРЯ-9", [
        "Вийти на вогневий рубіж №3 (командиру БнГ A0001)", "Вийти на вогневий рубіж №21 (командиру БнГ A0001)",
    ])]
