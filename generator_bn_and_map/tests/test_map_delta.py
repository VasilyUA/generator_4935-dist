"""Карта з урахуванням останнього експорту Дельти і штабних рішень
(синтетичні дані: координати - квадрат 31U DQ, назви - фонетичний алфавіт)."""
import copy
import csv
import uuid
import zipfile

from bn_fixtures import MAP_CONFIG
from content.delta_reader import learn_type_sidcs, name_key, parse_wkt, read_delta_export
from content.staff import StaffEntry
from generators.delta_csv_writer import DELTA_COLUMNS
from generators.generate_map import ORIGIN_DELTA, ORIGIN_PBD, ORIGIN_STAFF, build_map_objects, follows, generate_map, kmz_file_name
from utils.mgrs import mgrs_to_latlon

U_ALPHA, U_OLD, U_GONE, U_BRAVO = (f"00000000-0000-4000-8000-00000000000{i}" for i in range(1, 5))
SIDC_VP = MAP_CONFIG["SIDC_BY_TYPE"]["ВП"][0]
SIDC_OWN_VP = "10031020111211019999"
SIDC_OWN_TZ = "10031520002017000000"
SIDC_ENEMY = "10061010151211000100"
SIDC_BRO = MAP_CONFIG["LINE_SIDC"]["bro"][0]
SIDC_BOUNDARY = MAP_CONFIG["LINE_SIDC"]["boundary"][0]


def wkt_point(mgrs, shift=(0.0, 0.0), z=None):
    lat, lon = mgrs_to_latlon(mgrs)
    lat, lon = lat + shift[0], lon + shift[1]
    return f"POINT Z({lon:.7f} {lat:.7f} {z})" if z is not None else f"POINT ({lon:.7f} {lat:.7f})"


def wkt_line(mgrs_list):
    coords = [mgrs_to_latlon(m) for m in mgrs_list]
    return "LINESTRING (" + ", ".join(f"{lon:.7f} {lat:.7f}" for lat, lon in coords) + ")"


def write_delta(path, rows):
    columns = DELTA_COLUMNS + ["comment 1"]
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL, lineterminator="\n")
        writer.writerow(columns)
        for row in rows:
            writer.writerow([row.get(column, "") for column in columns])
    return str(path)


def _cache():
    return {
        "report_datetime": "2026-10-09T20:00:00",
        "points": [
            {"type": "ВП", "name": "АЛЬФА", "mgrs": "31U DQ 40100 10100", "quantity": 3, "group": ""},
            {"type": "ПВ", "name": "НОВА", "mgrs": "31U DQ 40300 10300", "quantity": 2, "group": ""},
            {"type": "ТЗ", "name": "БРАВО", "mgrs": "31U DQ 40500 10500", "quantity": 1, "group": "MAVIK"},
        ],
        "lines": [
            {"kind": "bro", "ref": "", "mgrs": ["31U DQ 40000 10000", "31U DQ 41000 10000", "31U DQ 41000 11000"]},
            {"kind": "boundary_left", "ref": "ліворуч", "mgrs": ["31U DQ 42000 10000", "31U DQ 42000 12000"]},
        ],
    }


DELTA_ROWS = [
    {"sidc": SIDC_OWN_VP, "id": U_ALPHA, "name": "ВП Альфа", "staff_comments": "коментар", "comment 1": "примітка",
     "coordinates": wkt_point("31U DQ 40100 10100", (0.00002, 0.0))},
    # Перейменована позиція - на карті стара назва, координати за 10-15 м від нових.
    {"sidc": SIDC_VP, "id": U_OLD, "name": "ПВ Стара", "coordinates": wkt_point("31U DQ 40300 10300", (0.0001, 0.0))},
    {"sidc": SIDC_VP, "id": U_GONE, "name": "ВП Зникла", "coordinates": wkt_point("31U DQ 45000 15000")},
    {"sidc": "10031020111220000000", "id": U_BRAVO, "name": "ТЗ Браво", "coordinates": wkt_point("31U DQ 40500 10500")},
    {"sidc": SIDC_OWN_TZ, "id": "own-tz", "name": "ТЗ Своя", "coordinates": wkt_point("31U DQ 46000 16000")},
    {"sidc": SIDC_ENEMY, "id": "enemy", "name": "Противник", "coordinates": wkt_point("31U DQ 47000 17000", z=120.5)},
    {"sidc": SIDC_BRO, "id": "bro-refined", "name": "БРО A0001 БАТ", "coordinates": wkt_line(
        ["31U DQ 40000 10000", "31U DQ 40500 10000", "31U DQ 41000 10000", "31U DQ 41000 10500", "31U DQ 41000 11000"])},
    {"sidc": SIDC_BOUNDARY, "id": "other-line", "name": " ", "coordinates": wkt_line(["31U DQ 44000 10000", "31U DQ 44000 12000"])},
    {"sidc": SIDC_BOUNDARY, "id": "staff-line", "name": "Межа A0002 – A0003",
     "coordinates": wkt_line(["31U DQ 43000 10000", "31U DQ 43000 12000"])},
    {"sidc": SIDC_ENEMY, "id": "broken", "name": "Без координат", "coordinates": "POINT EMPTY"},
]


def _staff():
    return [
        StaffEntry("вогневі_рубежі", "вр-1", title="ВР-1 БнГ", mgrs=["31U DQ 48000 18000", "31U DQ 48500 18000"]),
        StaffEntry("точки", "т-1", title="Точка збору", mgrs=["31U DQ 48800 18800"]),
        StaffEntry("райони", "р-1", title="Район", geometry="area",
                   mgrs=["31U DQ 49000 19000", "31U DQ 49500 19000", "31U DQ 49500 19500"]),
        StaffEntry("межі", "м-1", title="Межа A0002 – A0003", mgrs=["31U DQ 43000 10000", "31U DQ 43000 12000"],
                   source="Дельта"),
        StaffEntry("точки", "вимкнена", title="Не на карті", mgrs=["31U DQ 48900 18900"], on_map=False),
        StaffEntry("точки", "порожня", title="Порожня"),
    ]


def _config():
    config = copy.deepcopy(MAP_CONFIG)
    config["STAFF_SIDC"] = {"вогневі_рубежі": ["10032500001401000000", "Вогневий рубіж"]}
    return config


def test_parse_wkt_variants():
    assert parse_wkt("POINT Z(30.5 50.4 120.0)") == ("POINT", [(50.4, 30.5)])
    assert parse_wkt("LINESTRING Z(30 50 NaN, 31 51 NaN)") == ("LINESTRING", [(50.0, 30.0), (51.0, 31.0)])
    kind, ring = parse_wkt("POLYGON ((30 50, 31 50, 31 51, 30 50), (30.2 50.2, 30.3 50.2, 30.2 50.2))")
    assert kind == "POLYGON" and len(ring) == 4
    assert parse_wkt("POINT EMPTY") == ("", []) and parse_wkt("") == ("", [])


def test_name_key():
    assert name_key("Межа  a0002 — a0003") == name_key("МЕЖА A0002-A0003") == "МЕЖА A0002-A0003"
    assert name_key("«ВП» Альфа") == "ВП АЛЬФА"


def test_read_delta_export_and_learned_sidc(tmp_path):
    export = read_delta_export(write_delta(tmp_path / "БП - дельта.csv", DELTA_ROWS))
    assert export.columns[-1] == "comment 1" and len(export.objects) == 9
    assert len(export.warnings) == 1 and "рядок 11" in export.warnings[0]
    enemy = next(o for o in export.objects if o.name == "Противник")
    assert enemy.is_point and enemy.type_prefix == "" and enemy.row["coordinates"].startswith("POINT Z(")
    assert export.by_name("вп альфа").uid == U_ALPHA and export.by_name("Немає") is None
    learned = learn_type_sidcs(export, {"ВП", "ТЗ"})
    assert learned["ВП"] in (SIDC_OWN_VP, SIDC_VP) and set(learned) == {"ВП", "ТЗ"}
    assert learn_type_sidcs(None, {"ВП"}) == {}


def test_follows():
    line = [mgrs_to_latlon("31U DQ 40000 10000"), mgrs_to_latlon("31U DQ 41000 10000")]
    assert follows([mgrs_to_latlon("31U DQ 40500 10020")], line)
    assert not follows([mgrs_to_latlon("31U DQ 40500 10200")], line)
    assert not follows([], line[:1])


def test_merge_with_delta(tmp_path):
    export = read_delta_export(write_delta(tmp_path / "БП - дельта.csv", DELTA_ROWS))
    points, lines, warnings, stats = build_map_objects(_cache(), _config(), "a0001 бат", export, _staff())
    by_name = {p.name: p for p in points}
    # На карту - лише те, що є в ПБД і в наказі: «ТЗ Своя» і «Противник» є тільки на карті Дельти.
    assert set(by_name) == {"ВП Альфа", "ПВ Нова", "ТЗ Браво", "Точка збору"}
    alpha = by_name["ВП Альфа"]
    assert (alpha.uid, alpha.sidc, alpha.quantity, alpha.origin) == (U_ALPHA, SIDC_OWN_VP, 3, ORIGIN_PBD)
    assert alpha.row["staff_comments"] == "коментар"
    assert by_name["ПВ Нова"].uid == U_OLD and by_name["ТЗ Браво"].uid == U_BRAVO
    assert by_name["Точка збору"].origin == ORIGIN_STAFF
    assert any("ВП Зникла" in w and "видаліть вручну" in w for w in warnings)
    assert any("не перенесено 3 об'єктів" in w for w in warnings)

    origins = {(line.name, line.origin) for line in lines}
    assert ("БРО A0001 БАТ", ORIGIN_DELTA) in origins          # уточнена на карті - лишається з карти
    assert ("", ORIGIN_PBD) in origins                         # ліва межа з ПБД - нова
    assert {("ВР-1 БнГ", ORIGIN_STAFF), ("Район", ORIGIN_STAFF), ("Межа A0002 – A0003", ORIGIN_DELTA)} <= origins
    bro = next(line for line in lines if line.name == "БРО A0001 БАТ")
    assert len(bro.coords) == 5 and bro.uid == "bro-refined"
    area = next(line for line in lines if line.name == "Район")
    assert area.coords[0] == area.coords[-1] and len(area.coords) == 4
    fire_line = next(line for line in lines if line.name == "ВР-1 БнГ")
    assert fire_line.sidc == "10032500001401000000"
    assert fire_line.uid == str(uuid.uuid5(uuid.UUID("6f1d3c2a-8b4e-4f7a-9c1d-2e5b7a9c0d13"), "staff|вогневі_рубежі|вр-1"))
    assert stats == {"pbd_points": 3, "matched": 3, "staff": 3, "carried": 1, "skipped": 3, "stale": 1}


UAV_RECON = "10030100001103001800"


def test_group_code_beats_symbol_on_delta_map(tmp_path):
    # ТЗ мавіків: на карті Дельти - загальний знак ТЗ, у data_auto.json - код групи
    # (назва групи - як у ПБД, регістр не важить). id об'єкта з карти лишається.
    export = read_delta_export(write_delta(tmp_path / "export.csv", DELTA_ROWS))
    config = _config()
    config["SIDC_BY_GROUP"] = {"mavik": [UAV_RECON, "БпЛА, розвідка"]}
    points, _, warnings, _ = build_map_objects(_cache(), config, "a0001", export)
    bravo = next(p for p in points if p.name == "ТЗ Браво")
    assert (bravo.sidc, bravo.uid) == (UAV_RECON, U_BRAVO)
    assert not any("SIDC_BY_GROUP" in w for w in warnings)


def test_code_of_one_position_by_name(tmp_path):
    # Штаб уточнює знак окремої позиції (data_auto.json → map.SIDC_BY_NAME) - сильніше за знак на карті.
    export = read_delta_export(write_delta(tmp_path / "export.csv", DELTA_ROWS))
    config = _config()
    config["SIDC_BY_NAME"] = {"вп «альфа»": ["10031520001109040000", "Міномет, важкий"]}
    points, *_ = build_map_objects(_cache(), config, "a0001", export)
    alpha = next(p for p in points if p.name == "ВП Альфа")
    assert (alpha.sidc, alpha.description, alpha.uid) == ("10031520001109040000", "Міномет, важкий", U_ALPHA)


def test_group_without_code_is_reported(tmp_path):
    cache = _cache()
    cache["points"].append({"type": "ТЗ", "name": "ЧАРЛІ", "mgrs": "31U DQ 40700 10700", "quantity": 1, "group": "FPV"})
    config = _config()
    config["SIDC_BY_GROUP"] = {"MAVIK": [UAV_RECON, "БпЛА, розвідка"]}
    _, _, warnings, _ = build_map_objects(cache, config, "a0001")
    assert any(w.startswith("ТЗ: для груп FPV") for w in warnings)


def test_bare_codename_on_delta_map_is_the_same_position(tmp_path):
    # На карті Дельти ТЗ підписано лише позивним («ЧАРЛІ» в тій самій точці) - позиція з ПБД
    # оновлює саме цей об'єкт; такий самий позивний за 1.5 км не переноситься.
    rows = [
        {"sidc": SIDC_OWN_TZ, "id": "crew-near", "name": "ЧАРЛІ", "coordinates": wkt_point("31U DQ 40700 10700", (0.00001, 0.0))},
        {"sidc": SIDC_OWN_TZ, "id": "crew-far", "name": "ЧАРЛІ", "coordinates": wkt_point("31U DQ 42200 10700")},
    ]
    export = read_delta_export(write_delta(tmp_path / "export.csv", rows))
    cache = {"report_datetime": "", "lines": [], "points": [
        {"type": "ТЗ", "name": "ЧАРЛІ", "mgrs": "31U DQ 40700 10700", "quantity": 1, "group": "FPV"}]}
    points, _, _, stats = build_map_objects(cache, _config(), "a0001", export)
    assert [(p.uid, p.name, p.origin) for p in points] == [("crew-near", "ТЗ Чарлі", ORIGIN_PBD)]
    assert stats["skipped"] == 1


def test_take_off_point_takes_its_own_object_not_the_crew(tmp_path):
    # Як на реальній карті: у точці і «ТЗ «ЧАРЛІ» FPV», і екіпаж «ЧАРЛІ», і ще один екіпаж. ТЗ з ПБД
    # оновлює свій об'єкт (той самий позивний і тип); екіпажів немає ні в ПБД, ні в наказі - не йдуть.
    rows = [
        {"sidc": SIDC_OWN_TZ, "id": "crew", "name": "ЧАРЛІ", "coordinates": wkt_point("31U DQ 40700 10700", (0.00001, 0.0))},
        {"sidc": "10031020111220000000", "id": "tz", "name": "ТЗ «ЧАРЛІ» FPV", "coordinates": wkt_point("31U DQ 40700 10700")},
        {"sidc": SIDC_OWN_TZ, "id": "other-crew", "name": "ДЕЛЬТА-2", "coordinates": wkt_point("31U DQ 40700 10700")},
    ]
    export = read_delta_export(write_delta(tmp_path / "export.csv", rows))
    cache = {"report_datetime": "", "lines": [], "points": [
        {"type": "ТЗ", "name": "ЧАРЛІ", "mgrs": "31U DQ 40700 10700", "quantity": 1, "group": "FPV"}]}
    points, _, _, stats = build_map_objects(cache, _config(), "a0001", export)
    assert [(p.uid, p.origin) for p in points] == [("tz", ORIGIN_PBD)] and stats["skipped"] == 2


def test_line_with_same_name_but_new_geometry_is_replaced(tmp_path):
    rows = [{"sidc": SIDC_BRO, "id": "bro-old", "name": "БРО A0001 БАТ",
             "coordinates": wkt_line(["31U DQ 30000 10000", "31U DQ 31000 10000"])}]
    export = read_delta_export(write_delta(tmp_path / "д.csv", rows))
    _points, lines, warnings, _stats = build_map_objects(_cache(), _config(), "a0001 бат", export)
    bro = next(line for line in lines if line.name == "БРО A0001 БАТ")
    assert bro.uid == "bro-old" and bro.origin == ORIGIN_PBD and len(bro.coords) == 4
    assert any("замінено на рубіж з ПБД" in w for w in warnings)


def test_kmz_file_name():
    assert kmz_file_name({}, "a0001 бат") == "ТАКТИКА A0001 БАТ.kmz"
    assert kmz_file_name({"KMZ_FILE_NAME": "Карта {battalion}"}, "x") == "Карта X.kmz"
    assert kmz_file_name({"KMZ_FILE_NAME": "Готова.kmz"}, "x") == "Готова.kmz"


def test_generate_map_keeps_delta_rows(tmp_path):
    export = read_delta_export(write_delta(tmp_path / "БП - дельта.csv", DELTA_ROWS))
    data = {"unit": {"SHORT_UNIT_BATTALION": "a0001"}, "map": _config()}
    result = generate_map(data, _cache(), str(tmp_path / "map"), str(tmp_path / "icons"), export, _staff())
    assert result.stats["carried"] == 1
    with open(result.csv_path, encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    assert list(rows[0].keys())[-1] == "comment 1" and len(rows) == len(result.points) + len(result.lines)
    by_name = {row["name"]: row for row in rows}
    assert "Противник" not in by_name and "ТЗ Своя" not in by_name      # немає ні в ПБД, ні в наказі
    alpha = by_name["ВП Альфа"]
    assert (alpha["id"], alpha["staff_comments"], alpha["comment 1"], alpha["quantity"]) == (U_ALPHA, "коментар", "примітка", "3")
    assert alpha["observation_datetime"] == "2026-10-09T20:00:00" and alpha["coordinates"].startswith("POINT (")
    assert by_name["ВР-1 БнГ"]["coordinates"].startswith("LINESTRING (")
    assert by_name["Межа A0002 – A0003"]["id"] == "staff-line"        # штабне рішення, яке штаб веде на карті
    assert any("Дельта, рядок 11" in w for w in result.warnings)
    with zipfile.ZipFile(result.kmz_path) as archive:
        kml = archive.read("doc.kml").decode("utf-8")
    assert "Противник" not in kml and "ВР-1 БнГ" in kml
