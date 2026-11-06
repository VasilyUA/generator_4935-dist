"""KMZ для Кропиви: doc.kml (стандартні Point/LineString/Polygon, SIDC у
ExtendedData) + іконки files/<sidc>.png.

Кольори ліній за належністю з коду SIDC: свої - LINE_COLOR (стиль "line"),
противник - червоний ("line-hostile"), штабні рішення з data_auto.json -
окремий колір ("line-staff"); замкнені райони - полігони з напівпрозорою
заливкою того самого кольору ("area-…")."""
import xml.etree.ElementTree as ET
import zipfile

from utils.symbol_icons import HOSTILE, affiliation

KML_NS = "http://www.opengis.net/kml/2.2"
ICON_SCALE = "1.3"
LINE_WIDTH = "3"
# Своїх межі й рубежі Кропива малює чисто синім.
FRIEND_COLOR = "#0000FF"
HOSTILE_COLOR = "#FF0000"
STAFF_COLOR = "#FFD200"
_AREA_FILL_ALPHA = "40"


def format_coordinate(value):
    """7 знаків після коми (~1 см), без зайвих нулів - як в експорті Дельти."""
    text = f"{value:.7f}".rstrip("0")
    return text + "0" if text.endswith(".") else text


def kml_color(hex_rgb, alpha="ff"):
    """"#RRGGBB" -> KML "aabbggrr" (за замовчуванням непрозорий)."""
    value = (hex_rgb or "").lstrip("#")
    if len(value) != 6:
        value = FRIEND_COLOR.lstrip("#")
    return (alpha + value[4:6] + value[2:4] + value[0:2]).lower()


def _sub(parent, tag, text=None, **attrs):
    element = ET.SubElement(parent, tag, attrs)
    if text is not None:
        element.text = text
    return element


def _extended_data(placemark, sidc):
    extended = _sub(placemark, "ExtendedData")
    data = _sub(extended, "Data", name="sidc")
    _sub(data, "value", sidc)


def _coordinates(coords):
    return " ".join(f"{format_coordinate(lon)},{format_coordinate(lat)}" for lat, lon in coords)


def line_style_key(line):
    """"line" (свої), "line-hostile" (противник) чи "line-staff" (штабні рішення)."""
    if affiliation(line.sidc) == HOSTILE:
        return "line-hostile"
    if getattr(line, "origin", "") == "staff":
        return "line-staff"
    return "line"


def _add_styles(document, colors, used_keys, areas):
    for key in sorted(used_keys | {"line"}):
        style = _sub(document, "Style", id=key)
        line_style = _sub(style, "LineStyle")
        _sub(line_style, "color", kml_color(colors[key]))
        _sub(line_style, "width", LINE_WIDTH)
    for key in sorted(areas):
        style = _sub(document, "Style", id=key.replace("line", "area", 1))
        line_style = _sub(style, "LineStyle")
        _sub(line_style, "color", kml_color(colors[key]))
        _sub(line_style, "width", LINE_WIDTH)
        _sub(_sub(style, "PolyStyle"), "color", kml_color(colors[key], _AREA_FILL_ALPHA))


def icon_key(point):
    """Ключ значка точки: власний значок позиції (point.icon) або код SIDC."""
    return getattr(point, "icon", "") or point.sidc


def build_kml(points, lines, icon_sidcs, line_color=FRIEND_COLOR, hostile_color=HOSTILE_COLOR, staff_color=STAFF_COLOR):
    """icon_sidcs - ключі значків (коди SIDC або власні значки позицій), для яких є PNG."""
    kml = ET.Element("kml", xmlns=KML_NS)
    document = _sub(kml, "Document")
    colors = {"line": line_color, "line-hostile": hostile_color, "line-staff": staff_color}

    for sidc in sorted(icon_sidcs):
        style = _sub(document, "Style", id=f"sidc-{sidc}")
        icon_style = _sub(style, "IconStyle")
        _sub(icon_style, "scale", ICON_SCALE)
        icon = _sub(icon_style, "Icon")
        _sub(icon, "href", f"files/{sidc}.png")
    keys = {line_style_key(line) for line in lines}
    _add_styles(document, colors, keys, {line_style_key(line) for line in lines if getattr(line, "area", False)})

    for line in lines:
        placemark = _sub(document, "Placemark", id=line.uid)
        _sub(placemark, "name", line.name or "-")
        _sub(placemark, "visibility", "1")
        _sub(placemark, "description", line.description)
        key = line_style_key(line)
        area = getattr(line, "area", False)
        _sub(placemark, "styleUrl", "#" + (key.replace("line", "area", 1) if area else key))
        _extended_data(placemark, line.sidc)
        if area:
            ring = list(line.coords)
            if ring and ring[0] != ring[-1]:
                ring.append(ring[0])
            polygon = _sub(placemark, "Polygon")
            _sub(polygon, "tessellate", "1")
            boundary = _sub(_sub(polygon, "outerBoundaryIs"), "LinearRing")
            _sub(boundary, "coordinates", _coordinates(ring))
        else:
            geometry = _sub(placemark, "LineString")
            _sub(geometry, "tessellate", "1")
            _sub(geometry, "coordinates", _coordinates(line.coords))

    for point in points:
        placemark = _sub(document, "Placemark", id=point.uid)
        _sub(placemark, "name", point.name)
        _sub(placemark, "visibility", "1")
        _sub(placemark, "description", point.description)
        if icon_key(point) in icon_sidcs:
            _sub(placemark, "styleUrl", f"#sidc-{icon_key(point)}")
        _extended_data(placemark, point.sidc)
        geometry = _sub(placemark, "Point")
        _sub(geometry, "coordinates", f"{format_coordinate(point.lon)},{format_coordinate(point.lat)}")

    ET.indent(kml, space="    ")
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(kml, encoding="unicode")


def write_kmz(path, points, lines, icons, line_color=FRIEND_COLOR, hostile_color=HOSTILE_COLOR, staff_color=STAFF_COLOR):
    """icons - {ключ значка: PNG-байти}; у KMZ потрапляють лише іконки точок, які є на карті."""
    keys = {icon_key(p) for p in points}
    used_icons = {key: png for key, png in icons.items() if key in keys}
    kml = build_kml(points, lines, set(used_icons), line_color, hostile_color, staff_color)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("doc.kml", kml.encode("utf-8"))
        for sidc, png in sorted(used_icons.items()):
            archive.writestr(f"files/{sidc}.png", png)
