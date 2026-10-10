"""CSV-імпорт у Дельту - ті самі колонки, лапки й WKT, що й у її експорті."""
import csv

from generators.kmz_writer import format_coordinate

DELTA_COLUMNS = [
    "sidc", "id", "quantity", "name", "observation_datetime", "reliability_credibility",
    "staff_comments", "platform_type", "direction", "speed", "coordinates",
]


def point_wkt(lat, lon):
    return f"POINT ({format_coordinate(lon)} {format_coordinate(lat)})"


def line_wkt(coords):
    return "LINESTRING (" + ", ".join(f"{format_coordinate(lon)} {format_coordinate(lat)}" for lat, lon in coords) + ")"


def write_delta_csv(path, points, lines, reliability="", staff_comment="", platform_type=""):
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL, lineterminator="\n")
        writer.writerow(DELTA_COLUMNS)
        for point in points:
            writer.writerow([
                point.sidc, point.uid, "" if point.quantity is None else point.quantity, point.name,
                point.observed, reliability, staff_comment, platform_type, "", "", point_wkt(point.lat, point.lon),
            ])
        for line in lines:
            # Порожня назва лінії в експорті Дельти - один пробіл.
            writer.writerow([line.sidc, line.uid, "", line.name or " ", "", "", "", "", "", "", line_wkt(line.coords)])
