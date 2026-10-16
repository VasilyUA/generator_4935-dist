"""Кеш pbd + налаштування map (data.json і data_auto.json) (+ останній експорт Дельти і штабні
рішення data_auto.json → staff) -> точки й лінії карти -> output/map/<KMZ для
Кропиви> і <CSV для Дельти>.

На карту йде лише те, що є в ПБД і в бойовому наказі (штабні рішення з його
тексту). Останній CSV-експорт карти з resources/Дельта/ лише уточнює їх:
  - позиції з ПБД оновлюють свої об'єкти на карті (зіставлення за назвою, за
    позивним - «ТЗ Х» і «ТЗ «Х» Mavik», - або, для перейменованих, за
    координатами до 30 м): id і SIDC лишаються ті, що в Дельті (повторний імпорт
    оновить об'єкт, а не продублює), назва, кількість о/с і координати - з ПБД;
  - лінії з ПБД, які штаб уточнив у Дельті (усі вершини ПБД лежать уздовж лінії
    з карти), лишаються в уточненому вигляді; інакше - беруться з ПБД;
  - штабні рішення, які штаб веде на карті (об'єкт з тією самою назвою),
    беруться з карти;
  - позиції, яких у ПБД уже немає, на нову карту не потрапляють (список - у
    файлі перевірки, у Дельті їх треба видалити вручну);
  - решта об'єктів карти (екіпажі БпЛА, противник, інші лінії) на нову карту
    не переносяться - у файлі перевірки лише їхня кількість.

Перед записом злами ліній і районів, які будує програма (з ПБД і штабних
рішень), закруглюються (utils/geometry.py) - вершини лишаються ті самі,
додаються проміжні точки дуг. Лінії з карти Дельти лишаються такими, як їх
намалював штаб: саме так їх показує і Кропива."""
import hashlib
import math
import os
import uuid
from dataclasses import dataclass, field

from constants import DELTA_CSV_FILE_NAME, ICONS_TODO_FILE_NAME, KMZ_FILE_NAME_TEMPLATE
from content.delta_reader import learn_type_sidcs, name_key
from content.staff import SOURCE_DELTA
from generators.delta_csv_writer import write_delta_csv
from generators.kmz_writer import FRIEND_COLOR, HOSTILE_COLOR, STAFF_COLOR, write_kmz
from utils.geometry import DEFAULT_RADIUS_METERS, smooth_polyline
from utils.mgrs import mgrs_to_latlon
from utils.symbol_icons import fallback_icon, is_full_symbol

# Фіксований простір імен uuid5: однакова позиція отримує той самий id при
# кожній генерації, тож повторний імпорт у Дельту оновлює об'єкт, а не дублює.
_UUID_NAMESPACE = uuid.UUID("6f1d3c2a-8b4e-4f7a-9c1d-2e5b7a9c0d13")
_CLOSED_KINDS = {"bro", "rop", "vop"}
_LINE_SIDC_KEY = {
    "boundary_right": "boundary", "boundary_left": "boundary",
    "flot_rop": "flot", "flot_vop": "flot",
}
# Типи позицій ПБД - точки Дельти з такою першою частиною назви вважаються
# позиціями (а не штабними об'єктами) і звіряються з ПБД.
POSITION_TYPES = {"КСП", "ЗКП", "ВП", "ПВ", "СП", "ВЗ", "ТЗ", "ХАБ", "ПУ", "СПАР"}
_SAME_POSITION_METERS = 30
_LINE_FOLLOW_METERS = 60
_STAFF_SHAPE_SIDC = {"line": "boundary", "area": "rop"}

ORIGIN_PBD = "pbd"
ORIGIN_STAFF = "staff"
ORIGIN_DELTA = "delta"


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
    row: dict | None = None       # рядок експорту Дельти, поля якого зберігаються
    origin: str = ORIGIN_PBD
    icon: str = ""                # власний значок позиції (resources/map_icons/<назва>.png)


@dataclass
class MapLine:
    uid: str
    sidc: str
    description: str
    name: str
    coords: list
    row: dict | None = None
    origin: str = ORIGIN_PBD
    # Замкнений район: у KMZ - полігон із заливкою; wkt - як записати в CSV
    # (POLYGON лише для полігонів з експорту Дельти, решта - LINESTRING).
    area: bool = False
    wkt: str = "LINESTRING"


@dataclass
class MapResult:
    kmz_path: str
    csv_path: str
    points: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)


def title_word(word):
    return word[:1].upper() + word[1:].lower()


def point_display_name(type_, name):
    """("ВП", "АЛЬФА") -> "ВП Альфа" - як у назвах точок Дельти."""
    pretty = " ".join(title_word(word) for word in (name or "").split())
    return f"{type_} {pretty}".strip()


def kmz_file_name(config, battalion):
    template = (config or {}).get("KMZ_FILE_NAME") or KMZ_FILE_NAME_TEMPLATE
    name = template.format(battalion=(battalion or "").upper()).strip()
    return name if name.lower().endswith(".kmz") else f"{name}.kmz"


def _sidc_pair(value):
    if isinstance(value, str):
        return value, ""
    if isinstance(value, (list, tuple)) and value:
        return str(value[0]), str(value[1]) if len(value) > 1 else ""
    return None


def _group_sidc(config, group):
    """Код групи з map.SIDC_BY_GROUP - назва як у ПБД (MAVIK, FPV, БОМБЕРИ,
    Логістичні хаби …), без різниці регістру."""
    key = (group or "").strip().upper()
    if not key:
        return None
    for name, value in (config.get("SIDC_BY_GROUP") or {}).items():
        if name.strip().upper() == key:
            return _sidc_pair(value)
    return None


def _name_sidc(config, type_, name):
    """Код окремої позиції з map.SIDC_BY_NAME: «ВП Альфа» (або лише позивний «АЛЬФА»),
    без різниці регістру й лапок."""
    keys = {name_key(point_display_name(type_, name)), name_key(name)} - {""}
    for key, value in (config.get("SIDC_BY_NAME") or {}).items():
        if name_key(key) in keys:
            return _sidc_pair(value)
    return None


def _point_sidc(point, config, learned=None):
    """-> ((sidc, опис), звідки): "point" (явний у кеші), "name" (map.SIDC_BY_NAME),
    "group" (map.SIDC_BY_GROUP) - ці два сильніші й за знак на карті Дельти,
    "table" (map.SIDC_BY_TYPE), "delta" (найчастіший код цього типу на карті), "default"."""
    override = _sidc_pair(point.get("sidc"))
    if override:
        return override, "point"
    pair = _name_sidc(config, point.get("type", ""), point.get("name", ""))
    if pair:
        return pair, "name"
    pair = _group_sidc(config, point.get("group", ""))
    if pair:
        return pair, "group"
    pair = _sidc_pair((config.get("SIDC_BY_TYPE") or {}).get(point.get("type", "")))
    if pair:
        return pair, "table"
    if learned and point.get("type") in learned:
        return (learned[point["type"]], ""), "delta"
    return _sidc_pair(config.get("DEFAULT_SIDC")) or ("", ""), "default"


def _distance(a, b):
    lat = math.radians((a[0] + b[0]) / 2)
    return math.hypot((a[0] - b[0]) * 111320, (a[1] - b[1]) * 111320 * math.cos(lat))


def _segment_distance(p, a, b):
    """Відстань (м) від точки p до відрізка a-b (локальна площина)."""
    lat0 = math.radians(p[0])
    def xy(q):
        return ((q[1] - p[1]) * 111320 * math.cos(lat0), (q[0] - p[0]) * 111320)
    (ax, ay), (bx, by) = xy(a), xy(b)
    dx, dy = bx - ax, by - ay
    t = 0.0 if dx == dy == 0 else max(0.0, min(1.0, (-ax * dx - ay * dy) / (dx * dx + dy * dy)))
    return math.hypot(ax + t * dx, ay + t * dy)


def follows(points, polyline, tolerance=_LINE_FOLLOW_METERS):
    """Чи лежать усі points уздовж polyline (не далі tolerance метрів)."""
    if len(polyline) < 2:
        return False
    return all(min(_segment_distance(p, polyline[i], polyline[i + 1]) for i in range(len(polyline) - 1)) <= tolerance
               for p in points)


def mentions_codename(obj, codename):
    """Чи в назві об'єкта з карти є той самий позивний (окремими словами) - з
    будь-яким словом-типом і доповненням («ТЗ «Х» БОМБЕР», «ТУ Х») або без них
    (екіпаж «Х» на своїй ТЗ)."""
    wanted = (codename or "").split()
    if not wanted:
        return False
    words, n = name_key(obj.name).split(), len(wanted)
    return any(words[i:i + n] == wanted for i in range(len(words) - n + 1))


def _match_point(type_, display, latlon, candidates, codename=""):
    key = name_key(display)
    for obj in candidates:
        if name_key(obj.name) == key:
            return obj
    best = None
    for obj in candidates:
        named = mentions_codename(obj, codename)
        same_type = obj.type_prefix == type_ or {obj.type_prefix, type_} <= {"ВП", "ПВ", "СП", "ВЗ"}
        if not named and not same_type:
            continue
        distance = _distance(obj.coords[0], latlon)
        # Спершу той самий позивний і тип, далі - той самий позивний, далі - найближча того самого типу.
        rank = (not named, not same_type, distance)
        if distance <= _SAME_POSITION_METERS and (best is None or rank < best[0]):
            best = (rank, obj)
    return best[1] if best else None


def _pbd_points(pbd_cache, config, learned, delta_points, used, warnings):
    observed = pbd_cache.get("report_datetime", "")
    points, sidc_sources, groups = [], {}, {}
    for point in pbd_cache.get("points", []):
        if point.get("on_map") is False:
            continue
        type_, name = point.get("type", ""), point.get("name", "")
        display = point_display_name(type_, name)
        latlon = mgrs_to_latlon(point.get("mgrs", ""))
        if latlon is None:
            warnings.append(f"{display}: координати не розпізнано - точку пропущено.")
            continue
        (sidc, description), source = _point_sidc(point, config, learned)
        uid, row = str(uuid.uuid5(_UUID_NAMESPACE, f"point|{type_}|{name}")), None
        obj = _match_point(type_, display, latlon, [o for o in delta_points if id(o) not in used], name_key(name))
        if obj is not None:
            used.add(id(obj))
            uid, row = obj.uid or uid, obj.row
            # Код позиції чи групи (ТЗ бомберів / мавіків / FPV …) задано явно - він важить більше за знак на карті.
            if source not in ("point", "name", "group") and obj.sidc:
                sidc, source = obj.sidc, "map"
        sidc_sources.setdefault(source, []).append(type_)
        if point.get("group"):
            groups.setdefault(type_, {}).setdefault(point["group"], source in ("group", "name"))
        points.append(MapPoint(uid, sidc, description, display, point.get("quantity"), observed, *latlon, row=row))
    for type_, by_group in sorted(groups.items()):
        missing = sorted(group for group, coded in by_group.items() if not coded)
        if missing and any(by_group.values()):
            warnings.append(
                f"{type_}: для груп " + ", ".join(missing) + " немає коду в map.SIDC_BY_GROUP - на карті загальний "
                f"знак {type_}, а не знак своєї групи; впишіть код у data_auto.json."
            )
    defaults = sorted(set(sidc_sources.get("default", [])))
    if defaults:
        warnings.append(
            "Для типів " + ", ".join(defaults)
            + " немає SIDC ні в map.SIDC_BY_TYPE, ні на карті Дельти - використано DEFAULT_SIDC; впишіть коди в data_auto.json → map.SIDC_BY_TYPE."
        )
    learned_types = sorted(set(sidc_sources.get("delta", [])))
    if learned_types:
        warnings.append("SIDC для " + ", ".join(learned_types) + " взято з карти Дельти (найчастіший код цього типу).")
    return points


def _pbd_lines(pbd_cache, config, battalion, delta_lines, used, warnings):
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
        coords = [c for c in (mgrs_to_latlon(m) for m in line.get("mgrs", [])) if c is not None]
        if len(coords) < 2:
            warnings.append(f"Лінія «{kind}» {line.get('ref', '')}: замало координат - пропущено.")
            continue
        if kind in _CLOSED_KINDS and coords[0] != coords[-1]:
            coords.append(coords[0])
        default_name = f"БРО {battalion.upper()}".strip() if kind == "bro" else ""
        name = line.get("name", default_name)
        uid = str(uuid.uuid5(_UUID_NAMESPACE, f"line|{kind}|{line.get('ref', '')}"))
        # Спершу лінії з тим самим знаком, далі - будь-які лінії того самого
        # набору знаків (штаб міг позначити межу іншим різновидом лінії).
        candidates = sorted(
            (o for o in delta_lines if id(o) not in used and o.sidc[4:6] == pair[0][4:6]),
            key=lambda o: o.sidc != pair[0],
        )
        refined = next((o for o in candidates if follows(coords, o.coords)), None)
        if refined is not None:
            # Штаб уточнив лінію на карті, і вона досі відповідає ПБД - лишається з карти.
            used.add(id(refined))
            polygon = refined.geometry == "POLYGON"
            lines.append(MapLine(refined.uid or uid, refined.sidc, pair[1], refined.name, refined.coords,
                                 row=refined.row, origin=ORIGIN_DELTA, area=polygon,
                                 wkt="POLYGON" if polygon else "LINESTRING"))
            continue
        same_name = next((o for o in candidates if name and name_key(o.name) == name_key(name)), None)
        row = None
        if same_name is not None:
            used.add(id(same_name))
            uid, row = same_name.uid or uid, same_name.row
            warnings.append(f"Лінію «{name}» на карті Дельти замінено на рубіж з ПБД - він змінився.")
        lines.append(MapLine(uid, pair[0], pair[1], name, coords, row=row))
    return lines


def _staff_objects(entries, config):
    """Штабні рішення (data_auto.json → staff) без копії на карті Дельти.
    Точка завжди отримує знак точки (STAFF_POINT_SIDC або DEFAULT_SIDC): знак
    лінії чи району для точки Дельта не приймає ("Невалідне значення 'Point'")."""
    points, lines = [], []
    staff_sidc = config.get("STAFF_SIDC") or {}
    line_sidc = config.get("LINE_SIDC") or {}
    point_pair = _sidc_pair(config.get("STAFF_POINT_SIDC")) or _sidc_pair(config.get("DEFAULT_SIDC")) or ("", "")
    for entry in entries:
        if not entry.on_map or entry.source == SOURCE_DELTA or not entry.mgrs:
            continue
        coords = [c for c in (mgrs_to_latlon(m) for m in entry.mgrs) if c is not None]
        if not coords:
            continue
        shape = entry.shape
        uid = str(uuid.uuid5(_UUID_NAMESPACE, f"staff|{entry.group}|{entry.key}"))
        name = entry.title or entry.key
        if shape == "point" or len(coords) == 1:
            points.append(MapPoint(uid, point_pair[0], point_pair[1], name, None, "", *coords[0], origin=ORIGIN_STAFF))
            continue
        pair = (_sidc_pair(staff_sidc.get(entry.group))
                or _sidc_pair(line_sidc.get(_STAFF_SHAPE_SIDC.get(shape, "")))
                or _sidc_pair(config.get("DEFAULT_SIDC")) or ("", ""))
        if shape == "area" and coords[0] != coords[-1]:
            coords.append(coords[0])
        lines.append(MapLine(uid, pair[0], pair[1], name, coords, origin=ORIGIN_STAFF, area=shape == "area"))
    return points, lines


def build_map_objects(pbd_cache, config, battalion="", delta=None, staff_entries=()):
    """staff_entries - штабні рішення з тексту бойового наказу.
    -> (точки, лінії, попередження, статистика)."""
    warnings = []
    objects = list(delta.objects) if delta else []
    delta_points = [o for o in objects if o.is_point]
    delta_lines = [o for o in objects if not o.is_point]
    learned = learn_type_sidcs(delta, POSITION_TYPES) if delta else {}
    used = set()

    points = _pbd_points(pbd_cache, config, learned, delta_points, used, warnings)
    lines = _pbd_lines(pbd_cache, config, battalion, delta_lines, used, warnings)
    staff_points, staff_lines = _staff_objects(staff_entries, config)
    points += staff_points
    lines += staff_lines

    # Позиція з карти вважається застарілою, лише якщо має той самий знак, що й
    # позиції з ПБД: точку зльоту чи ПУ, яку штаб позначив власним знаком, не чіпаємо.
    position_sidcs = {p.sidc for p in points if p.origin == ORIGIN_PBD}
    # З карти береться лише те, що є штабним рішенням бойового наказу (та сама
    # назва): екіпажів, противника й інших об'єктів, яких немає ні в ПБД, ні в
    # наказі, на новій карті не має бути.
    staff_titles = {name_key(entry.title) for entry in staff_entries
                    if entry.on_map and entry.title and entry.source == SOURCE_DELTA}
    stale, carried, skipped = [], 0, 0
    for obj in objects:
        if id(obj) in used:
            continue
        if obj.is_point and obj.type_prefix in POSITION_TYPES and obj.sidc in position_sidcs:
            stale.append(obj.name)
            continue
        if name_key(obj.name) not in staff_titles:
            skipped += 1
            continue
        carried += 1
        if obj.is_point:
            points.append(MapPoint(obj.uid, obj.sidc, "", obj.name, None, "", *obj.coords[0], row=obj.row, origin=ORIGIN_DELTA))
        else:
            polygon = obj.geometry == "POLYGON"
            lines.append(MapLine(obj.uid, obj.sidc, "", obj.name, obj.coords, row=obj.row, origin=ORIGIN_DELTA,
                                 area=polygon, wkt="POLYGON" if polygon else "LINESTRING"))
    if stale:
        warnings.append(
            f"На карті Дельти є {len(stale)} позицій, яких немає в ПБД - на нову карту не перенесено "
            f"(у Дельті видаліть вручну): " + ", ".join(stale)
        )
    if skipped:
        warnings.append(
            f"З карти Дельти не перенесено {skipped} об'єктів, яких немає ні в ПБД, ні в бойовому наказі "
            "(екіпажі БпЛА, противник, інші лінії) - на карту йде лише те, що є в ПБД і в наказі."
        )
    stats = {
        "pbd_points": sum(1 for p in points if p.origin == ORIGIN_PBD),
        "matched": sum(1 for p in points if p.origin == ORIGIN_PBD and p.row is not None),
        "staff": len(staff_points) + len(staff_lines),
        "carried": carried,
        "skipped": skipped,
        "stale": len(stale),
    }
    return points, lines, warnings, stats


def smooth_lines(lines, config):
    """Закруглює злами ліній і районів з ПБД і штабних рішень (а з
    map.SMOOTH_DELTA_LINES: true - і з карти Дельти). map.SMOOTH_LINES: false -
    вимкнути; map.SMOOTH_RADIUS_METERS - наскільки широко закруглювати злам.
    -> скільки ліній змінено."""
    if not config.get("SMOOTH_LINES", True):
        return 0
    radius = config.get("SMOOTH_RADIUS_METERS") or DEFAULT_RADIUS_METERS
    with_delta = bool(config.get("SMOOTH_DELTA_LINES", False))
    changed = 0
    for line in lines:
        if line.origin == ORIGIN_DELTA and not with_delta:
            continue
        smooth = smooth_polyline(line.coords, radius=radius)
        if smooth != [(c[0], c[1]) for c in line.coords]:
            line.coords = smooth
            changed += 1
    return changed


def load_icons(icons_dir, sidcs):
    icons = {}
    for sidc in sidcs:
        path = os.path.join(icons_dir, f"{sidc}.png")
        if os.path.isfile(path):
            with open(path, "rb") as f:
                icons[sidc] = f.read()
    return icons


def map_icons(sidcs, icons_dir, extra_icons=None):
    """Значок для кожного SIDC: resources/map_icons/<sidc>.png, далі - з
    KMZ-експорту Дельти, далі - повний знак, який програма малює сама (свої
    БпЛА / гвинтокрили), інакше - рамка зі «?» (знак невідомий).
    -> ({sidc: PNG}, [невідомі SIDC])."""
    icons = load_icons(icons_dir, sidcs)
    for sidc in sorted(set(sidcs) - set(icons)):
        if extra_icons and sidc in extra_icons:
            icons[sidc] = extra_icons[sidc]
    missing = sorted(set(sidcs) - set(icons))
    for sidc in missing:
        icons[sidc] = fallback_icon(sidc)
    return icons, [sidc for sidc in missing if not is_full_symbol(sidc)]


def load_name_icons(icons_dir):
    """Власні значки позицій: resources/map_icons/<назва точки>.png («ВП Альфа.png»).
    -> {назва (name_key): (ключ значка, PNG)}."""
    found = {}
    if not os.path.isdir(icons_dir):
        return found
    for file_name in sorted(os.listdir(icons_dir)):
        stem, extension = os.path.splitext(file_name)
        if extension.lower() != ".png" or stem.isdigit():
            continue
        key = name_key(stem)
        with open(os.path.join(icons_dir, file_name), "rb") as f:
            png = f.read()
        found[key] = ("name-" + hashlib.sha1(key.encode("utf-8")).hexdigest()[:12], png)
    return found


def assign_icons(points, icons_dir, extra_icons=None):
    """Значок кожної точки: власний значок позиції (за назвою), далі - значок її
    коду (map_icons). -> ({ключ значка: PNG}, {sidc: [точки з невідомим знаком]})."""
    by_name = load_name_icons(icons_dir)
    icons = {}
    for point in points:
        named = by_name.get(name_key(point.name))
        if named:
            point.icon, icons[named[0]] = named
    by_code = [point for point in points if not point.icon]
    code_icons, unknown_codes = map_icons({point.sidc for point in by_code}, icons_dir, extra_icons)
    icons.update(code_icons)
    unknown = {}
    for point in by_code:
        if point.sidc in unknown_codes:
            unknown.setdefault(point.sidc, []).append(point.name)
    return icons, unknown


def _sidc_descriptions(config):
    """{код: опис} з таблиць data_auto.json → map."""
    descriptions = {}
    for table in (config.get("SIDC_BY_TYPE") or {}, config.get("SIDC_BY_GROUP") or {}):
        for key, value in table.items():
            pair = _sidc_pair(value)
            if pair and pair[0] not in descriptions:
                descriptions[pair[0]] = f"{key} ({pair[1]})" if pair[1] else key
    for key in ("STAFF_POINT_SIDC", "DEFAULT_SIDC"):
        pair = _sidc_pair(config.get(key))
        if pair and pair[0] not in descriptions:
            descriptions[pair[0]] = pair[1] or key
    return descriptions


def write_icons_todo(path, unknown, config):
    """Точки, знак яких програма не знає (на карті - «?»), і як це виправити."""
    descriptions = _sidc_descriptions(config)
    lines = [
        "Програма не знає, як позначаються ці знаки, - на карті Кропиви вони з «?».",
        "Виправити можна будь-яким способом:",
        "  1) значок позиції: PNG (прозоре тло) у resources/map_icons/<назва точки>.png, напр. «ВП Альфа.png» -",
        "     лише для цієї точки (так можна взяти значок з Кропиви);",
        "  2) значок коду: resources/map_icons/<код>.png - для всіх точок з цим кодом, або KMZ-експорт карти",
        "     з Дельти в resources/Дельта/ - значки кодів програма візьме з нього;",
        "  3) інший код: data_auto.json → map.SIDC_BY_NAME {\"<назва точки>\": [\"<код>\", \"<опис>\"]}",
        "     (чи map.SIDC_BY_TYPE / SIDC_BY_GROUP) - знак зміниться і в Дельті.",
        "",
    ]
    for sidc in sorted(unknown):
        names = sorted(unknown[sidc])
        description = descriptions.get(sidc) or ", ".join(sorted({name.split()[0] for name in names if name}))
        lines.append(f"{sidc} - {description}:")
        lines += [f"  - {name}" for name in names]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def generate_map(data, pbd_cache, output_dir, icons_dir, delta=None, staff_entries=(), extra_icons=None):
    config = data.get("map") or {}
    battalion = data.get("unit", {}).get("SHORT_UNIT_BATTALION", "")
    points, lines, warnings, stats = build_map_objects(pbd_cache, config, battalion, delta, staff_entries)
    stats["smoothed"] = smooth_lines(lines, config)
    if delta is not None:
        warnings = list(delta.warnings) + warnings
    os.makedirs(output_dir, exist_ok=True)
    kmz_path = os.path.join(output_dir, kmz_file_name(config, battalion))
    csv_path = os.path.join(output_dir, DELTA_CSV_FILE_NAME)
    icons, unknown = assign_icons(points, icons_dir, extra_icons)
    stats["unknown_symbols"] = sum(len(names) for names in unknown.values())
    if unknown:
        write_icons_todo(os.path.join(output_dir, ICONS_TODO_FILE_NAME), unknown, config)
        warnings.insert(0, (
            f"Невідомі знаки: {stats['unknown_symbols']} точок ({len(unknown)} кодів) - на карті позначені «?»; "
            f"які саме і як виправити - output/map/{ICONS_TODO_FILE_NAME}."
        ))
    write_kmz(kmz_path, points, lines, icons, config.get("LINE_COLOR", FRIEND_COLOR),
              config.get("HOSTILE_LINE_COLOR", HOSTILE_COLOR), config.get("STAFF_LINE_COLOR", STAFF_COLOR))
    write_delta_csv(
        csv_path, points, lines,
        reliability=config.get("RELIABILITY", ""), staff_comment=config.get("STAFF_COMMENT", ""),
        platform_type=config.get("PLATFORM_TYPE", ""), extra_columns=delta.columns if delta else (),
    )
    return MapResult(kmz_path, csv_path, points, lines, warnings, stats)
