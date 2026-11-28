"""Закруглення ліній карти (utils/geometry.py) і його застосування в KMZ/CSV
(синтетичні координати: зсуви в метрах від довільної точки, квадрат 31U DQ)."""
import copy
import csv
import math
import xml.etree.ElementTree as ET
import zipfile

from bn_fixtures import MAP_CONFIG
from content.delta_reader import DeltaExport, DeltaObject
from content.staff import SOURCE_DELTA, StaffEntry, apply_delta, is_rounded_copy
from generators.delta_csv_writer import DELTA_COLUMNS
from generators.generate_map import generate_map
from generators.kmz_writer import KML_NS
from utils.geometry import STEP_DEGREES, smooth_polyline
from utils.mgrs import mgrs_to_latlon

LAT0, LON0 = 48.85, 2.29
METERS_PER_DEGREE = 111320.0
KX = METERS_PER_DEGREE * math.cos(math.radians(LAT0))


def at(east, north):
    """Точка, зсунута на east/north метрів від (LAT0, LON0)."""
    return (LAT0 + north / METERS_PER_DEGREE, LON0 + east / KX)


def xy(point):
    return ((point[1] - LON0) * KX, (point[0] - LAT0) * METERS_PER_DEGREE)


def turns(coords):
    """Кути повороту (градуси) у кожній внутрішній точці ламаної."""
    result = []
    for a, b, c in zip(*(list(map(xy, coords))[k:] for k in range(3))):
        u, v = (b[0] - a[0], b[1] - a[1]), (c[0] - b[0], c[1] - b[1])
        result.append(math.degrees(math.atan2(abs(u[0] * v[1] - u[1] * v[0]), u[0] * v[0] + u[1] * v[1])))
    return result


def _to_segment(p, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    t = max(0.0, min(1.0, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / (dx * dx + dy * dy)))
    return math.hypot(a[0] + t * dx - p[0], a[1] + t * dy - p[1])


def deviation(smooth, original):
    """Найбільша відстань (м) від точок закругленої лінії до вихідної ламаної."""
    return max(min(_to_segment(xy(p), xy(a), xy(b)) for a, b in zip(original, original[1:])) for p in smooth)


def rounded7(coords):
    """Координати так, як їх запише CSV і поверне експорт Дельти (7 знаків)."""
    return [(round(lat, 7), round(lon, 7)) for lat, lon in coords]


def contains_in_order(result, original):
    remaining = iter(result)
    return all(any(point == other for other in remaining) for point in original)


L_SHAPE = [at(0, 0), at(0, 1000), at(1000, 1000)]
ZIGZAG = [at(0, 0), at(400, 600), at(800, 0), at(1200, 600), at(1600, 0)]
SQUARE = [at(0, 0), at(0, 800), at(800, 800), at(800, 0), at(0, 0)]


def test_straight_and_gently_bent_lines_are_unchanged():
    two_points = [at(0, 0), at(0, 1000)]
    straight = [at(0, 0), at(0, 500), at(0, 1000)]
    gentle = [at(0, 0), at(0, 1000), at(100, 2000)]
    for line in (two_points, straight, gentle):
        assert smooth_polyline(line) == line


def test_sharp_corner_is_rounded_through_the_same_vertices():
    result = smooth_polyline(L_SHAPE)
    assert len(result) > len(L_SHAPE)
    # Вершини - ті самі числа, у тому самому порядку; кінці лінії не зсуваються.
    assert result[0] == L_SHAPE[0] and result[-1] == L_SHAPE[-1] and contains_in_order(result, L_SHAPE)
    assert max(turns(result)) <= 2 * STEP_DEGREES
    assert deviation(result, L_SHAPE) <= 500 / 4


def test_closed_area_is_rounded_at_every_corner_including_the_first():
    result = smooth_polyline(SQUARE)
    assert result[0] == result[-1] == SQUARE[0] and contains_in_order(result, SQUARE)
    closing = turns([result[-2], result[0], result[1]])
    assert max(turns(result) + closing) <= 2 * STEP_DEGREES


def test_turn_back_and_corner_at_tiny_segment_stay_sharp():
    turn_back = [at(0, 0), at(0, 1000), at(30, 0)]
    tiny_segment = [at(0, 0), at(0, 1000), at(5, 1000), at(5, 0)]
    for line in (turn_back, tiny_segment):
        assert smooth_polyline(line) == line


def test_repeated_smoothing_adds_nothing_even_after_delta_round_trip():
    for line in (L_SHAPE, ZIGZAG, SQUARE):
        once = smooth_polyline(line)
        assert smooth_polyline(once) == once
        assert smooth_polyline(rounded7(once)) == rounded7(once)


def test_radius_limits_how_far_the_line_moves():
    long_corner = [at(0, 0), at(0, 5000), at(5000, 5000)]
    near = deviation(smooth_polyline(long_corner, radius=200), long_corner)
    far = deviation(smooth_polyline(long_corner, radius=1000), long_corner)
    assert 0 < near <= 200 / 4 and near < far <= 1000 / 4


# --- карта ---

SQUARE_MGRS = ["31U DQ 40000 10000", "31U DQ 40000 11000", "31U DQ 41000 11000", "31U DQ 41000 10000"]


def _data(**map_settings):
    config = copy.deepcopy(MAP_CONFIG)
    config.update(map_settings)
    return {"unit": {"SHORT_UNIT_BATTALION": "a0001"}, "map": config}


def _cache():
    return {"report_datetime": "2026-10-09T20:00:00", "points": [],
            "lines": [{"kind": "bro", "ref": "", "mgrs": SQUARE_MGRS, "on_map": True}]}


def _csv_rows(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def _pair_count(wkt):
    return wkt.count(",") + 1


def test_generate_map_rounds_lines_in_kmz_and_csv(tmp_path):
    result = generate_map(_data(), _cache(), str(tmp_path / "map"), str(tmp_path / "icons"))
    assert result.stats["smoothed"] == 1
    [row] = _csv_rows(result.csv_path)
    assert row["coordinates"].startswith("LINESTRING (") and _pair_count(row["coordinates"]) > 5
    with zipfile.ZipFile(result.kmz_path) as archive:
        root = ET.fromstring(archive.read("doc.kml"))
    coordinates = root.find(".//k:LineString/k:coordinates", {"k": KML_NS}).text.split()
    assert len(coordinates) == _pair_count(row["coordinates"])


def test_smoothing_can_be_switched_off(tmp_path):
    result = generate_map(_data(SMOOTH_LINES=False), _cache(), str(tmp_path / "map"), str(tmp_path / "icons"))
    assert result.stats["smoothed"] == 0
    [row] = _csv_rows(result.csv_path)
    assert _pair_count(row["coordinates"]) == 5   # 4 вершини + замикання контуру


def _delta_with_line(name, coords, sidc="10062500001401000000"):
    row = dict.fromkeys(DELTA_COLUMNS, "")
    row.update({"sidc": sidc, "id": "delta-line", "name": name, "comment 1": "примітка"})
    return DeltaExport("export.csv", columns=DELTA_COLUMNS + ["comment 1"], objects=[
        DeltaObject(row=row, sidc=sidc, uid="delta-line", name=name, geometry="LINESTRING", coords=list(coords)),
    ])


def test_line_from_delta_map_stays_as_drawn_unless_asked(tmp_path):
    # Штабне рішення, яке штаб веде на карті: у Кропиві лінія виглядає саме так, як її намалювали.
    delta = _delta_with_line("Лінія", L_SHAPE)
    staff = [StaffEntry("межі", "м-1", title="Лінія", mgrs=["31U DQ 40000 10000", "31U DQ 40000 11000"], source=SOURCE_DELTA)]
    result = generate_map(_data(), {"points": [], "lines": []}, str(tmp_path / "map"), str(tmp_path / "icons"), delta, staff)
    [written] = _csv_rows(result.csv_path)
    assert result.stats["smoothed"] == 0 and _pair_count(written["coordinates"]) == len(L_SHAPE)

    result = generate_map(_data(SMOOTH_DELTA_LINES=True), {"points": [], "lines": []},
                          str(tmp_path / "map2"), str(tmp_path / "icons"), delta, staff)
    [written] = _csv_rows(result.csv_path)
    assert (written["id"], written["name"], written["comment 1"]) == ("delta-line", "Лінія", "примітка")
    assert result.stats["smoothed"] == 1 and _pair_count(written["coordinates"]) > len(L_SHAPE)
    assert delta.objects[0].coords == L_SHAPE   # сам експорт не змінюється


# --- закруглене штабне рішення повертається з Дельти ---

CORNER_MGRS = ["31U DQ 40000 10000", "31U DQ 40000 11000", "31U DQ 41000 11000"]


def _boundary():
    return StaffEntry("межі", "м-1", title="Межа", text="від ({1}) через ({2}) до ({3})", mgrs=list(CORNER_MGRS))


def test_rounded_staff_line_back_from_delta_keeps_its_text():
    entry = _boundary()
    rounded = rounded7(smooth_polyline([mgrs_to_latlon(m) for m in CORNER_MGRS]))
    notes = apply_delta([entry], _delta_with_line("Межа", rounded))
    assert notes == [] and entry.source == SOURCE_DELTA
    assert entry.mgrs == CORNER_MGRS and entry.render().startswith("від (31U DQ 40000 10000)")


def test_rounded_staff_area_back_from_delta_keeps_its_corners():
    ring_mgrs = CORNER_MGRS + ["31U DQ 41000 10000"]
    entry = StaffEntry("райони", "р-1", title="Район", text="({1}) ({2}) ({3}) ({4})", mgrs=list(ring_mgrs), geometry="area")
    ring = [mgrs_to_latlon(m) for m in ring_mgrs]
    rounded = rounded7(smooth_polyline(ring + ring[:1]))
    assert apply_delta([entry], _delta_with_line("Район", rounded)) == [] and entry.mgrs == ring_mgrs


def test_staff_line_edited_on_delta_map_takes_map_coordinates():
    entry = _boundary()
    corners = [mgrs_to_latlon(m) for m in CORNER_MGRS]
    moved = corners[:2] + [mgrs_to_latlon("31U DQ 41500 11500")]
    edited = rounded7(smooth_polyline(moved))
    notes = apply_delta([entry], _delta_with_line("Межа", edited))
    assert len(notes) == 1 and entry.text == "" and entry.mgrs[-1] == "31U DQ 41500 11500"


def test_is_rounded_copy():
    vertices = [at(0, 0), at(0, 1000), at(1000, 1000)]
    assert is_rounded_copy(vertices, smooth_polyline(vertices))
    assert not is_rounded_copy(vertices, vertices)                     # стільки ж точок - звичайна заміна
    assert not is_rounded_copy(vertices[::-1], smooth_polyline(vertices))
