"""Рубежі оборони з окремого ручного файлу resources/рубежі.json (синтетичні
дані: квадрат 31U DQ, назви - фонетичний алфавіт)."""
import json
from datetime import date

import docx
import pytest

from bn_fixtures import make_data
from content.data_store import load_data, read_lines, split_data
from content.delta_reader import DeltaExport, DeltaObject
from content.staff import (
    LINES_BLOCK, SOURCE_LINES, apply_delta, has_lines_block, line_entries, line_paragraphs, load_staff,
    staff_to_data,
)
from generators.generate_bn import generate_bn
from generators.generate_map import ORIGIN_STAFF, build_map_objects
from utils.mgrs import mgrs_to_latlon

A, B, C = "31U DQ 40100 10100", "31U DQ 40200 10200", "31U DQ 40300 10300"
D, E = "31U DQ 41100 11100", "31U DQ 41200 11200"

LINES = [
    {"name": "Перший проміжний рубіж оборони", "objects": [
        {"key": "опорні_пункти:роп-1", "type": "РОП", "title": "РОП «АЛЬФА» (проміжний рубіж)",
         "text": "лісосмуга ({1}), стик лісосмуг ({2}), дорога ({3})", "mgrs": [A, B, C]},
        {"key": "передні_краї:пк-1", "title": "Передній край «АЛЬФА» (проміжний рубіж)", "mgrs": [A, B]},
    ]},
    {"name": "Другий проміжний рубіж оборони", "objects": [
        {"type": "РОП", "title": "РОП «БРАВО»", "text": "ліс ({1}), край лісу ({2}), яр ({3})", "mgrs": [D, E, C]},
        {"title": "Передній край «БРАВО»", "mgrs": [D, E]},
        {"type": "ВОП", "title": "ВОП «ЧАРЛІ»", "mgrs": []},
    ]},
]


def _write(path, content):
    path.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")
    return str(path)


def test_read_lines(tmp_path):
    assert read_lines(str(tmp_path / "немає.json")) == []
    assert read_lines(_write(tmp_path / "1.json", {"lines": LINES})) == LINES
    assert read_lines(_write(tmp_path / "2.json", LINES)) == LINES
    (tmp_path / "bad.json").write_text('{"lines": [ {"name": "x",} ]}', encoding="utf-8")
    with pytest.raises(ValueError, match="bad.json: помилка в рядку 1"):
        read_lines(str(tmp_path / "bad.json"))
    with pytest.raises(ValueError, match="список рубежів"):
        read_lines(_write(tmp_path / "3.json", {"lines": "ні"}))


def test_lines_are_loaded_but_never_written_to_data_files(tmp_path):
    data_path = _write(tmp_path / "data.json", {"unit": {}})
    data = load_data(data_path, None, _write(tmp_path / "рубежі.json", {"lines": LINES}))
    assert data["lines"] == LINES
    manual, auto = split_data(data)
    assert "lines" not in manual and "lines" not in auto


def test_line_entries_keys_types_and_geometry():
    entries = {entry.title: entry for entry in line_entries(LINES)}
    alpha = entries["РОП «АЛЬФА» (проміжний рубіж)"]
    assert (alpha.placeholder, alpha.shape, alpha.source, alpha.line) == (
        "staff:опорні_пункти:роп-1", "area", SOURCE_LINES, "Перший проміжний рубіж оборони")
    edge = entries["Передній край «БРАВО»"]                 # тип - з початку назви
    assert (edge.kind, edge.shape, edge.group) == ("передній край", "line", "рубежі")
    # Без "key" - сталий ключ з назв рубежу й об'єкта: той самий id на карті щоразу.
    assert edge.key == line_entries(LINES)[3].key and "БРАВО" in edge.key
    assert entries["ВОП «ЧАРЛІ»"].shape == "area"


def test_line_object_replaces_staff_entry_with_the_same_key():
    data = {"staff": {"опорні_пункти": {"роп-1": {"title": "Старий", "mgrs": [E]}, "інший": {"title": "Інший", "mgrs": [E]}}},
            "lines": LINES}
    entries = load_staff(data)
    by_placeholder = {entry.placeholder: entry for entry in entries}
    assert by_placeholder["staff:опорні_пункти:роп-1"].mgrs == [A, B, C]
    assert len([e for e in entries if e.placeholder == "staff:опорні_пункти:роп-1"]) == 1
    assert by_placeholder["staff:опорні_пункти:інший"].source == "data_auto.json"


def test_delta_map_does_not_override_and_lines_stay_out_of_data_auto():
    entries = load_staff({"lines": LINES})
    moved = [mgrs_to_latlon(m) for m in (D, E, A)]
    export = DeltaExport("x.csv", objects=[DeltaObject({}, "", "", "РОП «АЛЬФА» (проміжний рубіж)", "POLYGON", moved)])
    assert apply_delta(entries, export) == []
    assert entries[0].mgrs == [A, B, C] and entries[0].source == SOURCE_LINES
    assert staff_to_data(entries, {"межі": {}}) == {"межі": {}}


def test_new_line_gets_standard_paragraphs():
    entries = load_staff({"lines": LINES})
    paragraphs, empty = line_paragraphs(entries, {"staff:опорні_пункти:роп-1", "staff:передні_краї:пк-1"})
    assert paragraphs == [
        ["Другий проміжний рубіж оборони:", ["Другий проміжний рубіж оборони:"]],
        f"- РОП «БРАВО» за координатами: ліс ({D}), край лісу ({E}), яр ({C});",
        f"- Передній край «БРАВО»: ({D}), ({E}).",
    ]
    assert [entry.title for entry in empty] == ["ВОП «ЧАРЛІ»"]


def _document(with_block=True):
    content = ["Відійти на перший проміжний рубіж та зайняти РОП «АЛЬФА» за координатами {staff:опорні_пункти:роп-1} "
               "з переднім краєм: {staff:передні_краї:пк-1}."]
    if with_block:
        content.append({"block": LINES_BLOCK, "optional": True})
    return {"header": ["БОЙОВИЙ НАКАЗ"], "sections": [{"heading": "БОЙОВИЙ ПОРЯДОК", "content": content}]}


def test_order_keeps_written_text_and_adds_new_line(tmp_path):
    data = make_data()
    data["bn"]["document"] = _document()
    data["lines"] = LINES
    assert has_lines_block(data)
    result = generate_bn(data, None, {"points": [], "lines": [], "blocks": {}}, date(2026, 8, 24), "06.00", "5", str(tmp_path))
    texts = [p.text for p in docx.Document(result.path).paragraphs]
    assert any(f"за координатами лісосмуга ({A}), стик лісосмуг ({B}), дорога ({C}) з переднім краєм: ({A}), ({B})." in t
               for t in texts)
    assert "Другий проміжний рубіж оборони:" in texts
    assert f"- РОП «БРАВО» за координатами: ліс ({D}), край лісу ({E}), яр ({C});" in texts
    assert any(w.startswith("Рубежі з рубежі.json дописано стандартними абзацами") for w in result.warnings)
    assert any("без координат" in w and "ВОП «ЧАРЛІ»" in w for w in result.warnings)
    assert not any(entry.source == SOURCE_LINES for entry in result.unused_staff)


def test_order_without_place_for_new_lines_warns(tmp_path):
    data = make_data()
    data["bn"]["document"] = _document(with_block=False)
    data["lines"] = LINES
    result = generate_bn(data, None, {"points": [], "lines": [], "blocks": {}}, date(2026, 8, 24), "06.00", "5", str(tmp_path))
    assert any(w.startswith("Рубежі з рубежі.json нікуди дописати") for w in result.warnings)


def test_lines_on_the_map():
    data = make_data()
    _, lines, _, _ = build_map_objects({"points": [], "lines": []}, data["map"], "a0001", staff_entries=load_staff({"lines": LINES}))
    by_name = {line.name: line for line in lines}
    assert by_name["РОП «БРАВО»"].area and by_name["РОП «БРАВО»"].origin == ORIGIN_STAFF
    assert not by_name["Передній край «БРАВО»"].area
    assert "ВОП «ЧАРЛІ»" not in by_name                     # без координат - на карті нема чого малювати
