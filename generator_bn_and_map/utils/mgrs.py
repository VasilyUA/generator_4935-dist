"""MGRS -> широта/довгота WGS84 без сторонніх бібліотек (у requirements.txt
немає ні mgrs, ні pyproj).

Координати в документах пишуть по-різному: "31U DQ 12345 67890",
"31UDQ1234567890", "31U DQ 123-678" (100-метрова точність), а частину літер
набирають кирилицею ("33U ХР ..." - кириличні Х і Р виглядають так само, як
латинські). find_mgrs знаходить усі такі входження в довільному тексті,
mgrs_to_latlon перераховує одне в градуси (південно-західний кут квадрата
заданої точності - як і більшість конвертерів)."""
import math
import re
from dataclasses import dataclass

_A = 6378137.0
_F = 1 / 298.257223563
_K0 = 0.9996
_E2 = _F * (2 - _F)
_EP2 = _E2 / (1 - _E2)

_COLUMN_LETTERS = ("ABCDEFGH", "JKLMNPQR", "STUVWXYZ")
_ROW_LETTERS = "ABCDEFGHJKLMNPQRSTUV"
_BAND_LETTERS = "CDEFGHJKLMNPQRSTUVWX"

# Посимвольна заміна (довжина рядка не змінюється, тож позиції збігів у
# нормалізованому тексті збігаються з позиціями в оригіналі).
_CYRILLIC_TO_LATIN = str.maketrans("АВСЕНКМОРТХУ", "ABCEHKMOPTXY")

_MGRS_RE = re.compile(
    r"(?<!\d)(\d{1,2})\s*([C-HJ-NP-X])\s*([A-HJ-NP-Z])\s*([A-HJ-NP-V])\s*(\d{3,5})[\s\-]*(\d{3,5})(?!\d)"
)


@dataclass(frozen=True)
class MgrsMatch:
    start: int
    end: int
    raw: str
    normalized: str
    valid: bool


def _normalize_for_search(text):
    # upper() посимвольно і лише там, де він не змінює довжину (напр. "ß" -> "SS"
    # зсунув би позиції збігів відносно оригінального тексту).
    upper = "".join(c.upper() if len(c.upper()) == 1 else c for c in text)
    return upper.translate(_CYRILLIC_TO_LATIN)


def find_mgrs(text):
    """Усі MGRS-подібні входження в text по порядку. valid=False - збіг за
    формою, але половини мають різну кількість цифр (друкарська помилка) -
    такий збіг не перераховується, а лише повідомляється."""
    if not text:
        return []
    matches = []
    for m in _MGRS_RE.finditer(_normalize_for_search(text)):
        zone, band, column, row, easting, northing = m.groups()
        valid = len(easting) == len(northing) and 1 <= int(zone) <= 60
        normalized = f"{int(zone)}{band} {column}{row} {easting} {northing}"
        matches.append(MgrsMatch(m.start(), m.end(), text[m.start():m.end()], normalized, valid))
    return matches


def _meridian_arc(lat_deg):
    phi = math.radians(lat_deg)
    e2, e4, e6 = _E2, _E2 ** 2, _E2 ** 3
    return _A * (
        (1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256) * phi
        - (3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024) * math.sin(2 * phi)
        + (15 * e4 / 256 + 45 * e6 / 1024) * math.sin(4 * phi)
        - (35 * e6 / 3072) * math.sin(6 * phi)
    )


def utm_to_latlon(zone, easting, northing, northern=True):
    x = easting - 500000.0
    y = northing if northern else northing - 10000000.0
    mu = (y / _K0) / (_A * (1 - _E2 / 4 - 3 * _E2 ** 2 / 64 - 5 * _E2 ** 3 / 256))
    e1 = (1 - math.sqrt(1 - _E2)) / (1 + math.sqrt(1 - _E2))
    phi1 = (
        mu
        + (3 * e1 / 2 - 27 * e1 ** 3 / 32) * math.sin(2 * mu)
        + (21 * e1 ** 2 / 16 - 55 * e1 ** 4 / 32) * math.sin(4 * mu)
        + (151 * e1 ** 3 / 96) * math.sin(6 * mu)
        + (1097 * e1 ** 4 / 512) * math.sin(8 * mu)
    )
    sin1, cos1, tan1 = math.sin(phi1), math.cos(phi1), math.tan(phi1)
    n1 = _A / math.sqrt(1 - _E2 * sin1 ** 2)
    t1 = tan1 ** 2
    c1 = _EP2 * cos1 ** 2
    r1 = _A * (1 - _E2) / (1 - _E2 * sin1 ** 2) ** 1.5
    d = x / (n1 * _K0)
    lat = phi1 - (n1 * tan1 / r1) * (
        d ** 2 / 2
        - (5 + 3 * t1 + 10 * c1 - 4 * c1 ** 2 - 9 * _EP2) * d ** 4 / 24
        + (61 + 90 * t1 + 298 * c1 + 45 * t1 ** 2 - 252 * _EP2 - 3 * c1 ** 2) * d ** 6 / 720
    )
    lon = (
        d
        - (1 + 2 * t1 + c1) * d ** 3 / 6
        + (5 - 2 * c1 + 28 * t1 - 3 * c1 ** 2 + 8 * _EP2 + 24 * t1 ** 2) * d ** 5 / 120
    ) / cos1
    return math.degrees(lat), (zone - 1) * 6 - 180 + 3 + math.degrees(lon)


def mgrs_to_latlon(text):
    """Перше MGRS-входження в text -> (lat, lon) або None, якщо його немає чи
    воно некоректне."""
    matches = find_mgrs(text)
    if not matches or not matches[0].valid:
        return None
    zone_text, band, square, easting_text, northing_text = _split(matches[0].normalized)
    zone = int(zone_text)
    column, row = square
    column_set = _COLUMN_LETTERS[(zone - 1) % 3]
    if column not in column_set:
        return None
    precision = 10 ** (5 - len(easting_text))
    easting = (column_set.index(column) + 1) * 100000 + int(easting_text) * precision
    row_offset = 0 if zone % 2 == 1 else 5
    northing = ((_ROW_LETTERS.index(row) - row_offset) % 20) * 100000 + int(northing_text) * precision

    # 100-кілометрові літери рядка повторюються кожні 2000 км - справжню
    # північну координату обирає широтна смуга (band): найменша можлива
    # northing для смуги - на осьовому меридіані її південної межі.
    band_south_lat = -80 + 8 * _BAND_LETTERS.index(band)
    northern = band >= "N"
    min_northing = _K0 * _meridian_arc(band_south_lat) + (0 if northern else 10000000.0)
    while northing < min_northing - 100000:
        northing += 2000000
    return utm_to_latlon(zone, easting, northing, northern)


def _split(normalized):
    zone_band, square, easting, northing = normalized.split()
    return zone_band[:-1], zone_band[-1], square, easting, northing


def latlon_to_utm(lat, lon):
    """(lat, lon) WGS84 -> (зона, easting, northing) - пряма проєкція
    (формули Снайдера, як і зворотна utm_to_latlon). Норвезькі/шпіцбергенські
    винятки зон не враховуються."""
    zone = min(int((lon + 180) / 6) + 1, 60)
    phi = math.radians(lat)
    lam = math.radians(lon - ((zone - 1) * 6 - 180 + 3))
    sin_phi, cos_phi, tan_phi = math.sin(phi), math.cos(phi), math.tan(phi)
    n = _A / math.sqrt(1 - _E2 * sin_phi ** 2)
    t = tan_phi ** 2
    c = _EP2 * cos_phi ** 2
    a = cos_phi * lam
    easting = _K0 * n * (
        a + (1 - t + c) * a ** 3 / 6 + (5 - 18 * t + t ** 2 + 72 * c - 58 * _EP2) * a ** 5 / 120
    ) + 500000.0
    northing = _K0 * (_meridian_arc(lat) + n * tan_phi * (
        a ** 2 / 2 + (5 - t + 9 * c + 4 * c ** 2) * a ** 4 / 24
        + (61 - 58 * t + t ** 2 + 600 * c - 330 * _EP2) * a ** 6 / 720
    ))
    if lat < 0:
        northing += 10000000.0
    return zone, easting, northing


def latlon_to_mgrs(lat, lon, digits=5):
    """(lat, lon) -> "37U AB 12345 67890" (digits цифр на координату; значення
    відкидаються до південно-західного кута квадрата, як у mgrs_to_latlon)."""
    zone, easting, northing = latlon_to_utm(lat, lon)
    band = _BAND_LETTERS[max(0, min(len(_BAND_LETTERS) - 1, int((lat + 80) // 8)))]
    column = _COLUMN_LETTERS[(zone - 1) % 3][int(easting // 100000) - 1]
    row_offset = 0 if zone % 2 == 1 else 5
    row = _ROW_LETTERS[(int(northing // 100000) + row_offset) % 20]
    scale = 10 ** (5 - digits)
    # +5 см - захист від 12344.999 замість 12345: прямі й зворотні ряди
    # розходяться на міліметри, а MGRS відкидає дробову частину.
    e = int((easting % 100000 + 0.05) // scale)
    n = int((northing % 100000 + 0.05) // scale)
    return f"{zone}{band} {column}{row} {e:0{digits}d} {n:0{digits}d}"
