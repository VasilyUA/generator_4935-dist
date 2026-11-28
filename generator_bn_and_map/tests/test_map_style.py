"""Вигляд карти: значки (справжні з KMZ Дельти або спрощена рамка за
належністю), кольори ліній, райони-полігони, знак точки для штабних точок і
двовимірний WKT для імпорту в Дельту - синтетичні дані."""
import copy
import struct
import xml.etree.ElementTree as ET
import zipfile
import zlib

from bn_fixtures import MAP_CONFIG
from content.delta_reader import read_kmz_icons
from content.staff import StaffEntry
from generators.delta_csv_writer import line_wkt, polygon_wkt
from generators.generate_map import MapLine, MapPoint, build_map_objects, map_icons
from generators.kmz_writer import build_kml
from generators.pipeline import load_extras, save_missing_icons
from utils.symbol_icons import (
    AIR_SIZE, CIRCLE, DIAMOND, DOME, FRIEND, HOSTILE, NEUTRAL, RECT, UNKNOWN, affiliation, fallback_icon, frame_shape,
)

NS = {"k": "http://www.opengis.net/kml/2.2"}
FRIEND_SIDC, HOSTILE_SIDC = "10031020111211010000", "10061010151211000100"
KML_WITH_STYLES = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2"><Document>
  <Style id="s1"><IconStyle><Icon><href>files/a.png</href></Icon></IconStyle></Style>
  <StyleMap id="m1"><Pair><key>normal</key><styleUrl>#s1</styleUrl></Pair></StyleMap>
  <Placemark><styleUrl>#m1</styleUrl><ExtendedData><Data name="sidc"><value>11111111111111111111</value></Data></ExtendedData></Placemark>
  <Placemark><Style><IconStyle><Icon><href>files/b.png</href></Icon></IconStyle></Style>
    <ExtendedData><Data name="sidc"><value>22222222222222222222</value></Data></ExtendedData></Placemark>
  <Placemark><styleUrl>#немає</styleUrl><ExtendedData><Data name="sidc"><value>33333333333333333333</value></Data></ExtendedData></Placemark>
</Document></kml>"""


def write_kmz(path, members):
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in members.items():
            archive.writestr(name, payload)
    return str(path)


def test_friendly_equipment_gets_circle_frame():
    unit, equipment, hostile_equipment = "10031000001211010000", "10031500002017000000", "10061500002017000000"
    assert [frame_shape(s) for s in (unit, equipment, hostile_equipment)] == [RECT, CIRCLE, DIAMOND]
    assert fallback_icon(unit) != fallback_icon(equipment)


FPV, MAVIC, BOMBER, PLAIN_UAV = "10030100001103000100", "10030100001103001800", "10030100001102000202", "10030100001103000000"


def _pixels(png):
    """PNG з fallback_icon (RGBA, фільтр 0) -> рядки пікселів."""
    width, height = struct.unpack(">II", png[16:24])
    data, pos = b"", 8
    while pos < len(png):
        length = struct.unpack(">I", png[pos:pos + 4])[0]
        if png[pos + 4:pos + 8] == b"IDAT":
            data += png[pos + 8:pos + 8 + length]
        pos += 12 + length
    raw, stride = zlib.decompress(data), 1 + 4 * width
    return [[tuple(raw[y * stride + 1 + 4 * x:y * stride + 5 + 4 * x]) for x in range(width)] for y in range(height)]


def _black(rows, top, bottom):
    """Чорні пікселі в середній смузі купола (там, де літери модифікаторів)."""
    return sum(1 for row in rows[top:bottom] for pixel in row[25:39] if pixel == (0, 0, 0, 255))


def test_friendly_air_symbols_drawn_like_kropyva():
    # ТЗ FPV / мавіків / бомберів: купол, «галочка» БпЛА або «метелик» гвинтокрила, літери A / R / B і M.
    assert {frame_shape(s) for s in (FPV, MAVIC, BOMBER)} == {DOME}
    icons = {s: _pixels(fallback_icon(s)) for s in (FPV, MAVIC, BOMBER, PLAIN_UAV)}
    assert all(len(rows) == len(rows[0]) == AIR_SIZE for rows in icons.values())
    assert _black(icons[FPV], 10, 24) > 0 and _black(icons[PLAIN_UAV], 10, 24) == 0
    assert _black(icons[BOMBER], 46, 60) > 0 and _black(icons[FPV], 46, 60) == 0
    assert len({fallback_icon(s) for s in (FPV, MAVIC, BOMBER)}) == 3
    assert icons[FPV][AIR_SIZE - 4][31] == (128, 224, 255, 255)   # знизу купол відкритий, без лінії


def test_unknown_symbol_is_marked_not_empty():
    # Знак, якого програма не знає, - не порожня рамка (виглядала б як справжній знак), а «?» у ній.
    for sidc in ("10031020111296000000", "10031520002017000000", "10061020111211000000", "10030100001100000000"):
        rows = _pixels(fallback_icon(sidc))
        assert (0, 0, 0, 255) in [pixel for row in rows[22:44] for pixel in row[26:38]]


def test_affiliation_and_fallback_icon():
    assert [affiliation(s) for s in ("10031", "10061", "10041", "10011", "", "10051", "10021")] == [
        FRIEND, HOSTILE, NEUTRAL, UNKNOWN, UNKNOWN, HOSTILE, FRIEND]
    icons = {kind: fallback_icon(sidc) for kind, sidc in
             ((FRIEND, "10031"), (HOSTILE, "10061"), (NEUTRAL, "10041"), (UNKNOWN, "10011"))}
    for png in icons.values():
        assert png[:8] == bytes([0x89]) + b"PNG" + bytes([13, 10, 26, 10])
        assert struct.unpack(">II", png[16:24]) == (AIR_SIZE, AIR_SIZE)
    assert len(set(icons.values())) == 4


def test_read_kmz_icons(tmp_path):
    path = write_kmz(tmp_path / "Дельта.kmz", {
        "doc.kml": KML_WITH_STYLES, "files/a.png": b"A", "files/b.png": b"B",
        f"files/{FRIEND_SIDC}.png": b"F",
    })
    assert read_kmz_icons(path) == {FRIEND_SIDC: b"F", "11111111111111111111": b"A", "22222222222222222222": b"B"}
    broken = tmp_path / "зламаний.kmz"
    broken.write_bytes(b"not a zip")
    assert read_kmz_icons(str(broken)) == {}
    assert read_kmz_icons(write_kmz(tmp_path / "без kml.kmz", {"x.txt": "1"})) == {}
    assert read_kmz_icons(write_kmz(tmp_path / "поганий kml.kmz", {"doc.kml": "<kml"})) == {}


def test_map_icons_priority(tmp_path):
    (tmp_path / f"{FRIEND_SIDC}.png").write_bytes(b"own")
    icons, simplified = map_icons({FRIEND_SIDC, HOSTILE_SIDC, "10041"}, str(tmp_path), {HOSTILE_SIDC: b"delta"})
    assert icons[FRIEND_SIDC] == b"own" and icons[HOSTILE_SIDC] == b"delta"
    assert simplified == ["10041"] and icons["10041"] == fallback_icon("10041")


def test_load_extras_and_save_icons(tmp_path):
    folder = tmp_path / "Дельта"
    folder.mkdir()
    write_kmz(folder / "карта.kmz", {f"files/{FRIEND_SIDC}.png": b"F"})
    extras, used, problems = load_extras(str(tmp_path))
    assert extras.icons == {FRIEND_SIDC: b"F"} and used == [str(folder / "карта.kmz")] and problems == []
    icons_dir = tmp_path / "map_icons"
    assert save_missing_icons(str(icons_dir), extras.icons) == 1
    assert save_missing_icons(str(icons_dir), extras.icons) == 0
    assert save_missing_icons(str(icons_dir), {}) == 0
    assert (icons_dir / f"{FRIEND_SIDC}.png").read_bytes() == b"F"


def test_kml_line_colors_and_areas():
    lines = [
        MapLine("l1", FRIEND_SIDC, "", "Своя", [(48.1, 2.1), (48.2, 2.2)]),
        MapLine("l2", HOSTILE_SIDC, "", "Противник", [(48.1, 2.1), (48.2, 2.2)]),
        MapLine("l3", FRIEND_SIDC, "", "Штабна", [(48.1, 2.1), (48.2, 2.2)], origin="staff"),
        MapLine("l4", FRIEND_SIDC, "", "Район", [(48.1, 2.1), (48.2, 2.2), (48.1, 2.3)], origin="staff", area=True),
    ]
    root = ET.fromstring(build_kml([], lines, set(), "#00C0FF", "#FF0000", "#FFD200").split("\n", 1)[1])
    urls = [p.find("k:styleUrl", NS).text for p in root.findall(".//k:Placemark", NS)]
    assert urls == ["#line", "#line-hostile", "#line-staff", "#area-staff"]
    assert root.find(".//k:Style[@id='line-hostile']/k:LineStyle/k:color", NS).text == "ff0000ff"
    assert root.find(".//k:Style[@id='area-staff']/k:PolyStyle/k:color", NS).text == "4000d2ff"
    ring = root.find(".//k:Polygon/k:outerBoundaryIs/k:LinearRing/k:coordinates", NS).text.split()
    assert len(ring) == 4 and ring[0] == ring[-1]


def test_staff_points_get_point_symbol():
    config = copy.deepcopy(MAP_CONFIG)
    config["STAFF_SIDC"] = {"точки": ["10032500151101030000", "Район (знак лінії)"]}
    entries = [StaffEntry("точки", "т-1", title="Точка", mgrs=["31U DQ 40100 10100"])]
    points, *_ = build_map_objects({"points": [], "lines": []}, config, "", None, entries)
    assert points[0].sidc == MAP_CONFIG["DEFAULT_SIDC"][0]
    config["STAFF_POINT_SIDC"] = ["10031000000000000001", "Точка штабу"]
    points, *_ = build_map_objects({"points": [], "lines": []}, config, "", None, entries)
    assert points[0].sidc == "10031000000000000001"


def test_two_dimensional_wkt():
    assert line_wkt([(48.1, 2.1), (48.2, 2.2)]) == "LINESTRING (2.1 48.1, 2.2 48.2)"
    assert polygon_wkt([(48.1, 2.1), (48.2, 2.2), (48.1, 2.3)]) == "POLYGON ((2.1 48.1, 2.2 48.2, 2.3 48.1, 2.1 48.1))"
    assert polygon_wkt([(48.1, 2.1), (48.2, 2.2), (48.1, 2.1)]).count("2.1 48.1") == 2


def test_carried_polygon_and_point_are_rewritten(tmp_path):
    from generators.delta_csv_writer import write_delta_csv
    point = MapPoint("p", FRIEND_SIDC, "", "Точка", None, "", 48.1, 2.1,
                     row={"id": "p", "name": "Точка", "coordinates": "POINT Z(2.1 48.1 NaN)", "staff_comments": "к"},
                     origin="delta")
    area = MapLine("a", FRIEND_SIDC, "", "Район", [(48.1, 2.1), (48.2, 2.2), (48.1, 2.3)],
                   row={"id": "a", "name": "Район", "coordinates": "POLYGON Z((…))"}, origin="delta", area=True, wkt="POLYGON")
    path = tmp_path / "д.csv"
    write_delta_csv(str(path), [point], [area])
    text = path.read_text(encoding="utf-8")
    assert '"POINT (2.1 48.1)"' in text and '"POLYGON ((2.1 48.1, 2.2 48.2, 2.3 48.1, 2.1 48.1))"' in text
    assert "NaN" not in text and " Z(" not in text and '"к"' in text
