import zipfile

from bn_fixtures import MAP_CONFIG
from generators.generate_map import (
    DELTA_CSV_FILE_NAME, ICONS_TODO_FILE_NAME, build_map_objects, generate_map, load_icons, point_display_name,
)


def _cache(points=None, lines=None):
    return {
        "report_datetime": "2026-10-09T20:00:00",
        "points": points if points is not None else [
            {"type": "ВП", "name": "АЛЬФА-ДВА", "mgrs": "31U DQ 40100 10100", "quantity": 3, "group": "", "on_map": True},
            {"type": "КСП", "name": "", "mgrs": "31U DQ 40200 10200", "quantity": None, "group": "Бойовий порядок"},
            {"type": "ХАБ", "name": "БРАВО", "mgrs": "31U DQ 40300 10300", "quantity": 7, "group": "Логістичні хаби"},
            {"type": "ТЗ", "name": "ЧАРЛІ", "mgrs": "31U DQ 40400 10400", "quantity": 5, "group": "MAVIK"},
            {"type": "ТЗ", "name": "ДЕЛЬТА", "mgrs": "31U DQ 40500 10500", "quantity": 2, "group": "MAVIK",
             "sidc": ["10031000001100000000", "Власний"]},
            {"type": "ВП", "name": "ПРИХОВАНА", "mgrs": "31U DQ 40600 10600", "on_map": False},
            {"type": "ВП", "name": "БИТА", "mgrs": "31U DQ 40600 1060"},
        ],
        "lines": lines if lines is not None else [
            {"kind": "bro", "ref": "", "mgrs": ["31U DQ 40000 10000", "31U DQ 41000 10000", "31U DQ 41000 11000"], "on_map": True},
            {"kind": "boundary_left", "ref": "ліворуч", "mgrs": ["31U DQ 42000 10000", "31U DQ 42000 12000"], "on_map": True},
            {"kind": "flot_rop", "ref": "РОП «А»", "mgrs": ["31U DQ 40100 10300", "31U DQ 40200 10300"], "on_map": True},
            {"kind": "vop", "ref": "ВОП «Б»", "mgrs": ["31U DQ 40110 10110", "31U DQ 40120 10110"], "on_map": True},
            {"kind": "rop", "ref": "РОП «В»", "mgrs": ["31U DQ 40100 10100", "31U DQ 40200 1010"], "on_map": True},
            {"kind": "flot", "ref": "", "mgrs": ["31U DQ 40000 11000", "31U DQ 41000 11500"], "on_map": False},
        ],
    }


def test_point_display_name():
    assert point_display_name("ВП", "АЛЬФА-ДВА") == "ВП Альфа-два"
    assert point_display_name("ХАБ", "БРАВО 12") == "ХАБ Браво 12"
    assert point_display_name("КСП", "") == "КСП"


def test_points_sidc_precedence_and_filters():
    points, _, warnings, _ = build_map_objects(_cache(), MAP_CONFIG, "a0001 бат")
    by_name = {p.name: p for p in points}
    assert set(by_name) == {"ВП Альфа-два", "КСП", "ХАБ Браво", "ТЗ Чарлі", "ТЗ Дельта"}
    assert by_name["ВП Альфа-два"].sidc == "10031020111211010000"       # за типом
    assert by_name["ХАБ Браво"].sidc == "10031000001634000000"          # група важливіша за тип
    assert by_name["ТЗ Дельта"].sidc == "10031000001100000000"          # власний sidc точки
    assert by_name["ТЗ Чарлі"].sidc == "10031000000000000000"           # DEFAULT_SIDC
    assert by_name["ВП Альфа-два"].observed == "2026-10-09T20:00:00"
    assert any("ТЗ" in w and "DEFAULT_SIDC" in w for w in warnings)
    assert any("ВП Бита" in w for w in warnings)


def test_point_ids_are_stable():
    first, *_ = build_map_objects(_cache(), MAP_CONFIG)
    second, *_ = build_map_objects(_cache(), MAP_CONFIG)
    assert [p.uid for p in first] == [p.uid for p in second]
    assert len({p.uid for p in first}) == len(first)


def test_lines():
    _, lines, warnings, _ = build_map_objects(_cache(), MAP_CONFIG, "a0001 бат")
    names = [(l.name, l.sidc) for l in lines]
    assert names[0] == ("БРО A0001 БАТ", "10032500161101030000")
    assert lines[0].coords[0] == lines[0].coords[-1] and len(lines[0].coords) == 4
    assert names[1] == ("", "10032500161101010000")
    assert names[2] == ("", "10032500001401000000")
    assert len(lines) == 3
    assert any("vop" in w and "SIDC" in w for w in warnings)
    assert any("rop" in w and "замало" in w for w in warnings)


def test_line_name_and_sidc_override():
    lines = [{"kind": "bro", "ref": "", "name": "Свій", "sidc": "1234", "mgrs": ["31U DQ 40000 10000", "31U DQ 41000 10000"]}]
    _, built, *_ = build_map_objects(_cache(points=[], lines=lines), MAP_CONFIG, "x")
    assert (built[0].name, built[0].sidc) == ("Свій", "1234")


def test_load_icons(tmp_path):
    (tmp_path / "1.png").write_bytes(b"a")
    assert load_icons(str(tmp_path), {"1", "2"}) == {"1": b"a"}


def test_generate_map_writes_both_files(tmp_path):
    icons = tmp_path / "icons"
    icons.mkdir()
    (icons / "10031020111211010000.png").write_bytes(b"png")
    data = {"unit": {"SHORT_UNIT_BATTALION": "a0001"}, "map": MAP_CONFIG}
    result = generate_map(data, _cache(), str(tmp_path / "map"), str(icons))
    assert result.kmz_path.endswith("ТАКТИКА A0001.kmz") and result.csv_path.endswith(DELTA_CSV_FILE_NAME)
    with zipfile.ZipFile(result.kmz_path) as archive:
        names = archive.namelist()
        assert archive.read("files/10031020111211010000.png") == b"png"
    # Для решти кодів - рамка зі «?», жодна точка не лишається без значка, а програма про це каже.
    assert {f"files/{p.sidc}.png" for p in result.points} <= set(names)
    assert result.warnings[0].startswith("Невідомі знаки:") and result.stats["unknown_symbols"] > 0
    with open(result.csv_path, encoding="utf-8") as f:
        assert len(f.read().strip().split("\n")) == 1 + len(result.points) + len(result.lines)


def test_icons_todo_lists_points_without_real_icons(tmp_path):
    # Які значки дати: код, що це, і точки з ним. ТЗ (свій БпЛА) програма малює повністю - не в списку.
    icons = tmp_path / "icons"
    icons.mkdir()
    (icons / "10031020111211010000.png").write_bytes(b"png")
    config = dict(MAP_CONFIG, SIDC_BY_GROUP={"FPV": ["10030100001103000100", "БпЛА, ударний"],
                                             "Логістичні хаби": ["10031000001634000000", "Пункт постачання"]})
    cache = _cache()
    cache["points"].append({"type": "ТЗ", "name": "ДЕЛЬТА", "mgrs": "31U DQ 40400 10400", "group": "FPV"})
    result = generate_map({"unit": {}, "map": config}, cache, str(tmp_path / "map"), str(icons))
    with open(tmp_path / "map" / ICONS_TODO_FILE_NAME, encoding="utf-8") as f:
        todo = f.read()
    assert "10031000001634000000 - Логістичні хаби (Пункт постачання):\n  - ХАБ Браво\n" in todo
    assert "10031022111211010000 - КСП (Піхотний підрозділ):\n  - КСП\n" in todo
    assert "10031020111211010000" not in todo and "10030100001103000100" not in todo
    assert any(ICONS_TODO_FILE_NAME in w for w in result.warnings)


def test_own_icon_of_a_position_by_its_name(tmp_path):
    # Значок, узятий з Кропиви для однієї позиції (resources/map_icons/<назва>.png): лише для неї.
    icons = tmp_path / "icons"
    icons.mkdir()
    (icons / "хаб браво.png").write_bytes(b"own")
    result = generate_map({"unit": {}, "map": MAP_CONFIG}, _cache(), str(tmp_path / "map"), str(icons))
    hub = next(p for p in result.points if p.name == "ХАБ Браво")
    assert hub.icon.startswith("name-")
    with zipfile.ZipFile(result.kmz_path) as archive:
        assert archive.read(f"files/{hub.icon}.png") == b"own"
        kml = archive.read("doc.kml").decode("utf-8")
    assert f"#sidc-{hub.icon}" in kml
    with open(tmp_path / "map" / ICONS_TODO_FILE_NAME, encoding="utf-8") as f:
        assert "ХАБ Браво" not in f.read()
