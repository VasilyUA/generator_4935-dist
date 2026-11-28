"""CSV-експорт карти з Дельти (resources/Дельта/*.csv): об'єкти з sidc, id,
назвою і геометрією WKT (POINT / LINESTRING / POLYGON).

Експорт - "жива" карта штабу: звідси беруться коди SIDC, які штаб обрав для
позицій, координати штабних об'єктів з домовленими назвами (межі, вогневі
рубежі, райони - content/staff.py) і всі інші об'єкти (противник, екіпажі БпЛА
тощо), що переносяться в нову карту без змін (generators/generate_map.py)."""
import csv
import re
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field

_NUMBER_RE = re.compile(r"-?\d+(?:\.\d+)?")
_TYPE_PREFIX_RE = re.compile(r"^\s*([А-ЯІЇЄҐA-Z]{2,5})(?=\s|$)")


@dataclass
class DeltaObject:
    row: dict
    sidc: str
    uid: str
    name: str
    geometry: str           # POINT / LINESTRING / POLYGON
    coords: list            # [(lat, lon), …]

    @property
    def is_point(self):
        return self.geometry == "POINT"

    @property
    def type_prefix(self):
        """"ВП Литва" -> "ВП" (перше слово великими літерами, 2-5 літер)."""
        m = _TYPE_PREFIX_RE.match(self.name or "")
        return m.group(1) if m else ""


@dataclass
class DeltaExport:
    path: str
    columns: list = field(default_factory=list)
    objects: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    def by_name(self, name):
        key = name_key(name)
        return next((obj for obj in self.objects if key and name_key(obj.name) == key), None)


def name_key(name):
    """Назва для порівняння: без регістру, лапок, зайвих пробілів і різниці
    між дефісом і тире."""
    text = (name or "").upper().replace("–", "-").replace("—", "-")
    text = re.sub(r"[«»\"“”„‘’']", "", text)
    text = re.sub(r"\s*-\s*", "-", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_wkt(text):
    """"POINT (lon lat)" / "POINT Z(lon lat висота)" / "LINESTRING (…)" /
    "POLYGON ((…))" -> (тип, [(lat, lon)]); висота відкидається, з полігону
    береться зовнішній контур."""
    text = (text or "").strip()
    m = re.match(r"^([A-Za-z]+)(?:\s+(?:ZM|Z|M))?\s*\(", text)
    if not m:
        return "", []
    kind = m.group(1).upper()
    first_ring = text[m.end():].lstrip("( ").split(")", 1)[0]
    coords = []
    for vertex in first_ring.split(","):
        parts = vertex.split()
        if len(parts) < 2 or not (_NUMBER_RE.fullmatch(parts[0]) and _NUMBER_RE.fullmatch(parts[1])):
            continue
        coords.append((float(parts[1]), float(parts[0])))
    return kind, coords


def read_delta_export(path):
    export = DeltaExport(path)
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        export.columns = list(reader.fieldnames or [])
        for line_number, row in enumerate(reader, start=2):
            geometry, coords = parse_wkt(row.get("coordinates", ""))
            if not coords:
                export.warnings.append(f"Дельта, рядок {line_number}: не розпізнано координати - об'єкт пропущено.")
                continue
            export.objects.append(DeltaObject(
                row=dict(row), sidc=(row.get("sidc") or "").strip(), uid=(row.get("id") or "").strip(),
                name=(row.get("name") or "").strip(), geometry=geometry, coords=coords,
            ))
    return export


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _children(element, name):
    return [child for child in element.iter() if _local(child.tag) == name]


def _icon_href(style):
    hrefs = [href.text.strip() for href in _children(style, "href") if href.text]
    return hrefs[0] if hrefs else ""


def read_kmz_icons(path):
    """KMZ-експорт карти (напр. з Дельти) -> {sidc: PNG-байти}: файли
    "<20 цифр SIDC>.png" в архіві та іконки точок, у яких в ExtendedData є
    sidc. Пошкоджений архів - порожній словник."""
    icons = {}
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return icons
    with archive:
        names = archive.namelist()
        for name in names:
            m = re.search(r"(?<!\d)(\d{20})\.png$", name, re.IGNORECASE)
            if m:
                icons.setdefault(m.group(1), archive.read(name))
        kml_name = next((name for name in names if name.lower().endswith(".kml")), None)
        if kml_name is None:
            return icons
        try:
            root = ET.fromstring(archive.read(kml_name))
        except ET.ParseError:
            return icons
        styles = {style.get("id"): _icon_href(style) for style in _children(root, "Style") if style.get("id")}
        for style_map in _children(root, "StyleMap"):
            normal = next((pair for pair in _children(style_map, "Pair")
                           if any((key.text or "").strip() == "normal" for key in _children(pair, "key"))), None)
            url = next((u.text for u in _children(normal, "styleUrl")), "") if normal is not None else ""
            styles.setdefault(style_map.get("id"), styles.get((url or "").lstrip("#"), ""))
        for placemark in _children(root, "Placemark"):
            sidc = next(((value.text or "").strip() for data in _children(placemark, "Data") if data.get("name") == "sidc"
                         for value in _children(data, "value")), "")
            if not re.fullmatch(r"\d{20}", sidc) or sidc in icons:
                continue
            inline = [_icon_href(style) for style in _children(placemark, "Style")]
            url = next(((u.text or "").strip() for u in _children(placemark, "styleUrl")), "")
            href = next((h for h in inline if h), "") or styles.get(url.lstrip("#"), "")
            if href.lower().endswith(".png") and href in names:
                icons[sidc] = archive.read(href)
    return icons


def learn_type_sidcs(export, types):
    """{тип позиції: найчастіший SIDC точок Дельти, назва яких з нього
    починається} - коди, які штаб уже обрав для ТЗ, ХАБ, ПУ тощо."""
    counters = {}
    for obj in export.objects if export else []:
        prefix = obj.type_prefix
        if obj.is_point and prefix in types and obj.sidc:
            counters.setdefault(prefix, Counter())[obj.sidc] += 1
    return {prefix: counter.most_common(1)[0][0] for prefix, counter in counters.items()}
