"""Штабні рішення з data_auto.json["staff"] - те, чого немає в БР/ПБД/РР: межі рот
і взводів, передні краї й опорні пункти після відходу, райони зосередження,
вогневі рубежі резервів, маршрути, власні зони розвідки, точки евакуації тощо.

    "staff": {
      "<група>": {
        "<ключ>": {
          "title": "Межа 2 рмп – дшр",               # назва; так само - об'єкт у Дельті
          "text": "лісосмуга ({1}), стик лісосмуг ({2})",  # шаблон; {N} - N-та координата
          "mgrs": ["<MGRS 1>", "<MGRS 2>"],
          "geometry": "line" | "area" | "point",       # необов'язково - за кількістю точок
          "on_map": true                               # необов'язково
        }
      }
    }

У тексті БН - плейсхолдер {staff:<група>:<ключ>}: підставляється текст
шаблону з координатами (без шаблону - "(<MGRS 1>), (<MGRS 2>)"). Якщо в
CSV-експорті Дельти є об'єкт з такою самою назвою (title), координати
беруться з карти (apply_delta) - штаб може вести рішення прямо на карті."""
import copy
import math
import re
from dataclasses import dataclass, field

from content.delta_reader import name_key
from utils.mgrs import find_mgrs, latlon_to_mgrs, mgrs_to_latlon

PLACEHOLDER_PREFIX = "staff"
SOURCE_JSON = "data_auto.json"
SOURCE_DELTA = "Дельта"
_SLOT_RE = re.compile(r"\{(\d+)\}")
_PLACEHOLDER_RE = re.compile(r"\{(" + PLACEHOLDER_PREFIX + r"(?::[\w\-]+)+)\}")
_GEOMETRIES = {"line", "area", "point"}


@dataclass
class StaffEntry:
    group: str
    key: str
    title: str = ""
    text: str = ""
    mgrs: list = field(default_factory=list)
    geometry: str = ""
    on_map: bool = True
    source: str = SOURCE_JSON

    @property
    def placeholder(self):
        return f"{PLACEHOLDER_PREFIX}:{self.group}:{self.key}"

    @property
    def shape(self):
        """line / area / point - явно заданий або за кількістю точок."""
        if self.geometry in _GEOMETRIES:
            return self.geometry
        return "point" if len(self.mgrs) == 1 else "line"

    def render(self):
        if self.text and _SLOT_RE.search(self.text):
            def slot(m):
                index = int(m.group(1)) - 1
                return self.mgrs[index] if 0 <= index < len(self.mgrs) else m.group(0)
            return _SLOT_RE.sub(slot, self.text)
        if self.text:
            return self.text
        return ", ".join(f"({m})" for m in self.mgrs)


def document_placeholders(data):
    """Плейсхолдери штабних рішень ({staff:група:ключ}), які є в тексті бойового
    наказу (data_auto.json → bn_document): на карту йдуть лише ці рішення."""
    found = set()

    def walk(node):
        if isinstance(node, str):
            found.update(_PLACEHOLDER_RE.findall(node))
        elif isinstance(node, dict):
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)

    walk(((data or {}).get("bn") or {}).get("document"))
    return found


def load_staff(data):
    """Склеєні дані (staff з data_auto.json) -> [StaffEntry] (групи й ключі - як у файлі)."""
    entries = []
    for group, items in ((data or {}).get("staff") or {}).items():
        if not isinstance(items, dict):
            continue
        for key, value in items.items():
            if not isinstance(value, dict):
                continue
            mgrs = value.get("mgrs") or []
            entries.append(StaffEntry(
                group=group, key=key, title=str(value.get("title", "")), text=str(value.get("text", "")),
                mgrs=[str(m) for m in (mgrs if isinstance(mgrs, list) else [mgrs]) if m],
                geometry=str(value.get("geometry", "")), on_map=value.get("on_map", True) is not False,
            ))
    return entries


def _vertices(obj, closed):
    coords = list(obj.coords)
    if closed and len(coords) > 2 and coords[0] == coords[-1]:
        coords = coords[:-1]
    return coords


def _near(a, b, meters):
    k = math.cos(math.radians(a[0]))
    return math.hypot((a[0] - b[0]) * 111320, (a[1] - b[1]) * 111320 * k) <= meters


def is_rounded_copy(vertices, coords, meters=2.0):
    """Чи лінія з карти (coords) - ті самі вершини, лише закруглені програмою
    (utils/geometry.py): усі вершини є серед її точок у тому самому порядку, а
    точок більше. Таку лінію штаб не змінював - після імпорту CSV у Дельту й
    нового експорту вона повертається саме такою."""
    if len(coords) <= len(vertices):
        return False
    remaining = iter(coords)
    return all(any(_near(vertex, point, meters) for point in remaining) for vertex in vertices)


def apply_delta(entries, export):
    """Координати штабних рішень з карти Дельти (об'єкт з тією самою назвою).
    Однакова кількість точок - текст шаблону лишається, змінюються лише
    координати; інша - текст будується з координат. Лінія, яку лише закруглила
    програма, нічого не змінює. -> [примітки для звіту]."""
    notes = []
    if export is None:
        return notes
    for entry in entries:
        if not entry.title:
            continue
        obj = export.by_name(entry.title)
        if obj is None:
            continue
        coords = _vertices(obj, entry.shape == "area")
        vertices = [c for c in (mgrs_to_latlon(m) for m in entry.mgrs) if c is not None]
        if vertices and is_rounded_copy(vertices, coords):
            entry.source = SOURCE_DELTA
            continue
        mgrs = [latlon_to_mgrs(lat, lon) for lat, lon in coords]
        if len(mgrs) != len(entry.mgrs):
            if entry.mgrs:
                notes.append(
                    f"«{entry.title}»: на карті Дельти {len(mgrs)} точок, у data_auto.json - {len(entry.mgrs)}; "
                    "текст складено з координат карти без описів місцевості."
                )
            entry.text = ""
        entry.mgrs = mgrs
        entry.source = SOURCE_DELTA
    return notes


def staff_to_data(entries, original=None):
    """[StaffEntry] -> словник для data_auto.json["staff"]: координати й текст
    (зокрема оновлені з карти Дельти) записуються в ті самі записи, решта їх
    полів (title, geometry, on_map …) лишається як була."""
    result = copy.deepcopy(original or {})
    for entry in entries:
        group = result.setdefault(entry.group, {})
        if not isinstance(group, dict):
            continue
        item = group.setdefault(entry.key, {"title": entry.title})
        item["mgrs"] = list(entry.mgrs)
        if entry.text or "text" in item:
            item["text"] = entry.text
    return result


def staff_values(entries):
    """{плейсхолдер без дужок: (текст, StaffEntry)} для DocumentBuilder."""
    return {entry.placeholder: (entry.render(), entry) for entry in entries}


def duplicate_coordinates(entries):
    """Різні записи (крім меж - вони спільні для сусідів) з однаковими
    координатами - імовірна описка. -> [[плейсхолдер, …], …]."""
    groups = {}
    for entry in entries:
        if entry.mgrs and entry.group != "межі":
            groups.setdefault((entry.group, tuple(entry.mgrs)), []).append(entry.placeholder)
    return [placeholders for placeholders in groups.values() if len(placeholders) > 1]


def invalid_coordinates(entries):
    """Записи з координатами, які не розпізнано як MGRS - для звіту."""
    problems = []
    for entry in entries:
        bad = [m for m in entry.mgrs if not any(match.valid for match in find_mgrs(m))]
        if bad:
            problems.append(f"{entry.placeholder}: не MGRS - {', '.join(bad)}")
    return problems
