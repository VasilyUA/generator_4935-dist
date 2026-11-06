"""CSV-імпорт у Дельту - ті самі колонки, лапки й WKT, що й у її експорті.

Об'єкти, взяті з експорту Дельти (MapPoint/MapLine.row), зберігають усі свої
поля (коментарі штабу, достовірність, додаткові колонки експорту). Геометрія
ж завжди пишеться заново у двовимірному WKT ("POINT (lon lat)",
"LINESTRING (…)", "POLYGON ((…))"): у власному експорті Дельти трапляються
"POINT Z(… висота)" і "NaN" замість висоти, а такий рядок імпорт відхиляє."""
import csv

from generators.kmz_writer import format_coordinate

DELTA_COLUMNS = [
    "sidc", "id", "quantity", "name", "observation_datetime", "reliability_credibility",
    "staff_comments", "platform_type", "direction", "speed", "coordinates",
]


def _pairs(coords):
    return ", ".join(f"{format_coordinate(lon)} {format_coordinate(lat)}" for lat, lon in coords)


def point_wkt(lat, lon):
    return f"POINT ({format_coordinate(lon)} {format_coordinate(lat)})"


def line_wkt(coords):
    return f"LINESTRING ({_pairs(coords)})"


def polygon_wkt(coords):
    ring = list(coords)
    if ring and ring[0] != ring[-1]:
        ring.append(ring[0])
    return f"POLYGON (({_pairs(ring)}))"


def _point_values(point, defaults):
    values = dict(point.row) if point.row else dict(defaults)
    values["coordinates"] = point_wkt(point.lat, point.lon)
    if point.origin == "delta" and point.row:
        return values
    values.update({"sidc": point.sidc, "id": point.uid, "name": point.name})
    if point.quantity is not None or not point.row:
        values["quantity"] = "" if point.quantity is None else point.quantity
    if point.observed or not point.row:
        values["observation_datetime"] = point.observed
    return values


def _line_values(line):
    values = dict(line.row) if line.row else {}
    values["coordinates"] = polygon_wkt(line.coords) if line.wkt == "POLYGON" else line_wkt(line.coords)
    if line.origin == "delta" and line.row:
        return values
    # Порожня назва лінії в експорті Дельти - один пробіл.
    values.update({"sidc": line.sidc, "id": line.uid, "name": line.name or " "})
    return values


def write_delta_csv(path, points, lines, reliability="", staff_comment="", platform_type="", extra_columns=()):
    columns = DELTA_COLUMNS + [c for c in extra_columns if c and c not in DELTA_COLUMNS]
    defaults = {"reliability_credibility": reliability, "staff_comments": staff_comment, "platform_type": platform_type}
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL, lineterminator="\n")
        writer.writerow(columns)
        for point in points:
            values = _point_values(point, defaults)
            writer.writerow([values.get(column, "") for column in columns])
        for line in lines:
            values = _line_values(line)
            writer.writerow([values.get(column, "") for column in columns])
