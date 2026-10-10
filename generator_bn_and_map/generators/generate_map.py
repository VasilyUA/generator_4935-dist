"""Кеш pbd + налаштування map з data.json -> точки й лінії карти ->
output/map/<KMZ для Кропиви> і <CSV для Дельти>."""
import os
import uuid
from dataclasses import dataclass, field

from constants import DELTA_CSV_FILE_NAME, KMZ_FILE_NAME
from generators.delta_csv_writer import write_delta_csv
from generators.kmz_writer import write_kmz
from utils.mgrs import mgrs_to_latlon

# Фіксований простір імен uuid5: однакова позиція отримує той самий id при
# кожній генерації, тож повторний імпорт у Дельту оновлює об'єкт, а не дублює.
_UUID_NAMESPACE = uuid.UUID("6f1d3c2a-8b4e-4f7a-9c1d-2e5b7a9c0d13")
_CLOSED_KINDS = {"bro", "rop", "vop"}
_LINE_SIDC_KEY = {
    "boundary_right": "boundary", "boundary_left": "boundary",
    "flot_rop": "flot", "flot_vop": "flot",
}


@dataclass
class MapPoint:
    uid: str
    sidc: str
    description: str
    name: str
    quantity: int | None
    observed: str
    lat: float
    lon: float


@dataclass
class MapLine:
    uid: str
    sidc: str
    description: str
    name: str
    coords: list


@dataclass
class MapResult:
    kmz_path: str
    csv_path: str
    points: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def title_word(word):
    return word[:1].upper() + word[1:].lower()


def point_display_name(type_, name):
    """("ВП", "АЛЬФА") -> "ВП Альфа" - як у назвах точок Дельти."""
    pretty = " ".join(title_word(word) for word in (name or "").split())
    return f"{type_} {pretty}".strip()


def _sidc_pair(value):
    if isinstance(value, str):
        return value, ""
    if isinstance(value, (list, tuple)) and value:
        return str(value[0]), str(value[1]) if len(value) > 1 else ""
    return None


def _point_sidc(point, config):
    override = _sidc_pair(point.get("sidc"))
    if override:
        return override, False
    for table, key in ((config.get("SIDC_BY_GROUP") or {}, point.get("group", "")),
                       (config.get("SIDC_BY_TYPE") or {}, point.get("type", ""))):
        pair = _sidc_pair(table.get(key))
        if pair:
            return pair, False
    return _sidc_pair(config.get("DEFAULT_SIDC")) or ("", ""), True


def build_map_objects(pbd_cache, config, battalion=""):
    warnings = []
    observed = pbd_cache.get("report_datetime", "")
    points, without_sidc = [], []
    for point in pbd_cache.get("points", []):
        if point.get("on_map") is False:
            continue
        type_, name = point.get("type", ""), point.get("name", "")
        latlon = mgrs_to_latlon(point.get("mgrs", ""))
        if latlon is None:
            warnings.append(f"{point_display_name(type_, name)}: координати не розпізнано - точку пропущено.")
            continue
        (sidc, description), is_default = _point_sidc(point, config)
        if is_default and type_ not in without_sidc:
            without_sidc.append(type_)
        points.append(MapPoint(
            str(uuid.uuid5(_UUID_NAMESPACE, f"point|{type_}|{name}")), sidc, description,
            point_display_name(type_, name), point.get("quantity"), observed, *latlon,
        ))
    if without_sidc:
        warnings.append(
            "Для типів " + ", ".join(without_sidc)
            + " немає SIDC у map.SIDC_BY_TYPE - використано DEFAULT_SIDC; впишіть точні коди в data.json."
        )

    lines = []
    line_sidc = config.get("LINE_SIDC") or {}
    for line in pbd_cache.get("lines", []):
        if line.get("on_map") is False:
            continue
        kind = line.get("kind", "")
        pair = _sidc_pair(line.get("sidc")) or _sidc_pair(line_sidc.get(_LINE_SIDC_KEY.get(kind, kind)))
        if not pair:
            warnings.append(f"Для ліній «{kind}» немає SIDC у map.LINE_SIDC - лінію пропущено.")
            continue
        coords = [mgrs_to_latlon(m) for m in line.get("mgrs", [])]
        coords = [c for c in coords if c is not None]
        if len(coords) < 2:
            warnings.append(f"Лінія «{kind}» {line.get('ref', '')}: замало координат - пропущено.")
            continue
        if kind in _CLOSED_KINDS and coords[0] != coords[-1]:
            coords.append(coords[0])
        default_name = f"БРО {battalion.upper()}".strip() if kind == "bro" else ""
        lines.append(MapLine(
            str(uuid.uuid5(_UUID_NAMESPACE, f"line|{kind}|{line.get('ref', '')}")), pair[0], pair[1],
            line.get("name", default_name), coords,
        ))
    return points, lines, warnings


def load_icons(icons_dir, sidcs):
    icons = {}
    for sidc in sidcs:
        path = os.path.join(icons_dir, f"{sidc}.png")
        if os.path.isfile(path):
            with open(path, "rb") as f:
                icons[sidc] = f.read()
    return icons


def generate_map(data, pbd_cache, output_dir, icons_dir):
    config = data.get("map") or {}
    battalion = data.get("unit", {}).get("SHORT_UNIT_BATTALION", "")
    points, lines, warnings = build_map_objects(pbd_cache, config, battalion)
    os.makedirs(output_dir, exist_ok=True)
    kmz_path = os.path.join(output_dir, KMZ_FILE_NAME)
    csv_path = os.path.join(output_dir, DELTA_CSV_FILE_NAME)
    icons = load_icons(icons_dir, {p.sidc for p in points})
    missing_icons = sorted({p.sidc for p in points} - set(icons))
    if missing_icons:
        warnings.append(f"Немає іконок для {len(missing_icons)} SIDC у resources/map_icons - Кропива покаже типовий значок.")
    write_kmz(kmz_path, points, lines, icons, config.get("LINE_COLOR", "#00C0FF"))
    write_delta_csv(
        csv_path, points, lines,
        reliability=config.get("RELIABILITY", ""), staff_comment=config.get("STAFF_COMMENT", ""),
        platform_type=config.get("PLATFORM_TYPE", ""),
    )
    return MapResult(kmz_path, csv_path, points, lines, warnings)
