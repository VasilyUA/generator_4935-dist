"""resources/data.json: читання, безпечний запис (резервна копія + атомарна
заміна файлу) і кеш "pbd" - позиції, рубежі та блоки тексту з останніх ПБД і
РОП_ВОП. Кеш оновлюється лише на прохання користувача (index.py), тож
вручну виправлені в ньому значення (on_map, sidc) переживають оновлення."""
import json
import math
import os
import re
import shutil
import tempfile
from datetime import datetime

from utils.mgrs import mgrs_to_latlon

# Типи позицій, які з РОП_ВОП потрапляють на карту як точки (БРО/РОП/ВОП -
# райони, не точки).
_POINT_TYPES = {"КСП", "ВП", "ПВ", "СП", "ВЗ", "ТЗ", "ХАБ", "ПУ", "СПАР", "ЗКП"}
# Лінії, які за замовчуванням наносяться на карту (решта зберігається в кеші
# з on_map: false - можна ввімкнути вручну).
DEFAULT_LINES_ON_MAP = {"bro", "flot", "boundary_right", "boundary_left", "rop"}
_SAME_POSITION_METERS = 15
_PRESERVED_POINT_FIELDS = ("on_map", "sidc")


def load_data(path):
    if not os.path.isfile(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_data(path, data):
    """Попередня версія файлу копіюється в <path>.bak, новий вміст пишеться в
    тимчасовий файл поруч і лише тоді атомарно замінює data.json - збій під
    час запису не лишить напівзаписаний файл."""
    if os.path.isfile(path):
        shutil.copy2(path, path + ".bak")
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".data-", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_path, path)
    except BaseException:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def normalize_name(name):
    return re.sub(r"[^0-9A-ZА-ЯІЇЄҐ]", "", (name or "").upper().replace("Ё", "Е"))


def _names_match(a, b):
    a, b = normalize_name(a), normalize_name(b)
    if not a or not b:
        return False
    if a == b:
        return True
    # "АЛЬФА" / "АЛЬФИ" - однина/множина того самого позивного.
    shorter, longer = sorted((a, b), key=len)
    return len(shorter) >= 4 and longer.startswith(shorter) and len(longer) - len(shorter) <= 2


def distance_m(mgrs_a, mgrs_b):
    a, b = mgrs_to_latlon(mgrs_a), mgrs_to_latlon(mgrs_b)
    if a is None or b is None:
        return math.inf
    lat1, lon1, lat2, lon2 = map(math.radians, (*a, *b))
    h = math.sin((lat2 - lat1) / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 2 * 6371000 * math.asin(math.sqrt(h))


def _find_entry(point, entries, used):
    candidates = [(i, e) for i, e in enumerate(entries) if i not in used and e.type in _POINT_TYPES]
    for i, entry in candidates:
        if point.name and normalize_name(entry.name) == normalize_name(point.name):
            return i, entry
    for i, entry in candidates:
        if point.name and _names_match(entry.name, point.name):
            return i, entry
    for i, entry in candidates:
        if entry.mgrs and distance_m(entry.mgrs[0], point.mgrs) <= _SAME_POSITION_METERS:
            return i, entry
    return None, None


def _report_datetime(date_text, time_text):
    try:
        moment = datetime.strptime(f"{date_text} {time_text or '00.00'}", "%d.%m.%Y %H.%M")
    except ValueError:
        return ""
    return moment.strftime("%Y-%m-%dT%H:%M:%S")


def _label(type_, name):
    return f"{type_} «{name}»" if name else type_


def build_pbd_cache(report, table, previous=None, folder_name=""):
    """ПБД (+ необов'язково РОП_ВОП) -> словник для data.json["pbd"] і список
    попереджень. Тип і кількість о/с беруться з РОП_ВОП (там точніше: напр.
    "Вогнева засідка" замість загального "ВП"), координати - з ПБД."""
    warnings = list(report.warnings)
    entries = list(table.entries) if table else []
    if table:
        warnings.extend(table.warnings)
    previous = previous or {}
    previous_points = {(p.get("type"), p.get("name")): p for p in previous.get("points", [])}
    previous_lines = {(l.get("kind"), l.get("ref", "")): l for l in previous.get("lines", [])}

    used = set()
    points = []
    for point in report.points:
        index, entry = _find_entry(point, entries, used)
        type_, quantity = point.type, None
        if entry is not None:
            used.add(index)
            quantity = entry.quantity
            if entry.type != "ЗКП":
                type_ = entry.type
            if entry.name and normalize_name(entry.name) != normalize_name(point.name):
                warnings.append(
                    f"Назви розходяться: ПБД {_label(point.type, point.name)} / РОП_ВОП {_label(entry.type, entry.name)} - взято назву з ПБД."
                )
        elif table and point.name:
            warnings.append(f"{_label(point.type, point.name)} є в ПБД, але немає в РОП_ВОП - кількість о/с невідома.")
        points.append({
            "type": type_, "name": point.name, "mgrs": point.mgrs, "quantity": quantity,
            "group": point.group, "rop": point.rop, "on_map": True,
        })

    for i, entry in enumerate(entries):
        if i in used or entry.type not in _POINT_TYPES or not entry.mgrs:
            continue
        if entry.type != "ЗКП":
            warnings.append(f"{_label(entry.type, entry.name)} є лише в РОП_ВОП - додано з його координатами.")
        points.append({
            "type": entry.type, "name": entry.name, "mgrs": entry.mgrs[0], "quantity": entry.quantity,
            "group": "РОП_ВОП", "rop": "", "on_map": True,
        })

    for point in points:
        old = previous_points.get((point["type"], point["name"]))
        if old:
            for field_name in _PRESERVED_POINT_FIELDS:
                if field_name in old:
                    point[field_name] = old[field_name]

    lines = []
    for line in report.lines:
        old = previous_lines.get((line.kind, line.ref))
        entry = {"kind": line.kind, "ref": line.ref, "mgrs": list(line.mgrs), "on_map": line.kind in DEFAULT_LINES_ON_MAP}
        if old:
            for field_name in ("on_map", "name", "sidc"):
                if field_name in old:
                    entry[field_name] = old[field_name]
        lines.append(entry)

    cache = {
        "updated": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "folder": folder_name,
        "source": os.path.basename(report.path),
        "rop_vop_source": os.path.basename(table.path) if table else "",
        "report_datetime": _report_datetime(report.report_date, report.report_time),
        "ksp_place": report.ksp_place,
        "map_edition": report.map_edition,
        "ksp_mgrs": report.ksp_mgrs,
        "zkp_mgrs": table.zkp_mgrs if table else "",
        "reserve_areas": report.reserve_areas,
        "blocks": report.blocks,
        "points": points,
        "lines": lines,
    }
    return cache, warnings
