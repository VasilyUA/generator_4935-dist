import csv
import io
import xml.etree.ElementTree as ET
import zipfile

import pytest

from generators.delta_csv_writer import DELTA_COLUMNS, line_wkt, point_wkt, write_delta_csv
from generators.generate_map import MapLine, MapPoint
from generators.kmz_writer import build_kml, format_coordinate, kml_color, write_kmz

NS = {"k": "http://www.opengis.net/kml/2.2"}
POINTS = [
    MapPoint("id-1", "10031020111211010000", "Піхотний підрозділ", "ВП Альфа", 3, "2026-10-09T20:00:00", 48.1234567, 2.5),
    MapPoint("id-2", "10031000000000000000", "Невизначений", "ХАБ Браво", None, "", 48.2, 2.6),
]
LINES = [MapLine("id-3", "10032500161101030000", "Тилова", "БРО A0001", [(48.1, 2.1), (48.2, 2.2), (48.1, 2.1)])]


@pytest.mark.parametrize("value, expected", [(2.5, "2.5"), (48.1234567, "48.1234567"), (36.10000001, "36.1"), (-1.0, "-1.0")])
def test_format_coordinate(value, expected):
    assert format_coordinate(value) == expected


def test_kml_color():
    assert kml_color("#00C0FF") == "ffffc000"
    assert kml_color("bad") == "ffffc000"


def test_build_kml_structure():
    root = ET.fromstring(build_kml(POINTS, LINES, {"10031020111211010000"}).split("\n", 1)[1])
    placemarks = root.findall(".//k:Placemark", NS)
    assert [pm.find("k:name", NS).text for pm in placemarks] == ["БРО A0001", "ВП Альфа", "ХАБ Браво"]
    line, first, second = placemarks
    assert line.find("k:LineString/k:coordinates", NS).text == "2.1,48.1 2.2,48.2 2.1,48.1"
    assert line.find("k:styleUrl", NS).text == "#line"
    assert first.find("k:styleUrl", NS).text == "#sidc-10031020111211010000"
    assert second.find("k:styleUrl", NS) is None
    assert first.find("k:ExtendedData/k:Data/k:value", NS).text == "10031020111211010000"
    assert first.find("k:Point/k:coordinates", NS).text == "2.5,48.1234567"
    assert root.find(".//k:Style[@id='line']/k:LineStyle/k:color", NS).text == "ffffc000"


def test_unnamed_line_gets_dash():
    unnamed = [MapLine("id-4", "1", "", "", [(1.0, 2.0), (3.0, 4.0)])]
    root = ET.fromstring(build_kml([], unnamed, set()).split("\n", 1)[1])
    assert root.find(".//k:Placemark/k:name", NS).text == "-"


def test_write_kmz_includes_only_used_icons(tmp_path):
    path = tmp_path / "карта.kmz"
    write_kmz(str(path), POINTS, LINES, {"10031020111211010000": b"png-1", "99999999999999999999": b"png-2"})
    with zipfile.ZipFile(path) as archive:
        assert sorted(archive.namelist()) == ["doc.kml", "files/10031020111211010000.png"]
        assert archive.read("files/10031020111211010000.png") == b"png-1"


def test_wkt():
    assert point_wkt(48.5, 2.25) == "POINT (2.25 48.5)"
    assert line_wkt([(1.0, 2.0), (3.0, 4.0)]) == "LINESTRING (2.0 1.0, 4.0 3.0)"


def test_write_delta_csv(tmp_path):
    path = tmp_path / "дельта.csv"
    write_delta_csv(str(path), POINTS, LINES, reliability="A2", staff_comment="A0000 /A0001", platform_type="AIRREC")
    raw = path.read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf") and b"\r\n" not in raw
    lines = raw.decode("utf-8").splitlines()
    assert lines[0] == ",".join(f'"{c}"' for c in DELTA_COLUMNS)
    rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8"))))
    assert rows[0]["quantity"] == "3" and rows[0]["staff_comments"] == "A0000 /A0001"
    assert rows[1]["quantity"] == "" and rows[1]["coordinates"] == "POINT (2.6 48.2)"
    assert rows[2]["name"] == "БРО A0001" and rows[2]["reliability_credibility"] == ""
    assert rows[2]["coordinates"].startswith("LINESTRING (2.1 48.1, ")
