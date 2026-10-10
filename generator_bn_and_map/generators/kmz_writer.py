"""KMZ для Кропиви: doc.kml (стандартні Point/LineString, SIDC у
ExtendedData) + іконки files/<sidc>.png, якщо вони є."""
import xml.etree.ElementTree as ET
import zipfile

KML_NS = "http://www.opengis.net/kml/2.2"
ICON_SCALE = "1.3"
LINE_WIDTH = "3"


def format_coordinate(value):
    """7 знаків після коми (~1 см), без зайвих нулів - як в експорті Дельти."""
    text = f"{value:.7f}".rstrip("0")
    return text + "0" if text.endswith(".") else text


def kml_color(hex_rgb):
    """"#RRGGBB" -> KML "aabbggrr" (непрозорий)."""
    value = hex_rgb.lstrip("#")
    if len(value) != 6:
        value = "00C0FF"
    return ("ff" + value[4:6] + value[2:4] + value[0:2]).lower()


def _sub(parent, tag, text=None, **attrs):
    element = ET.SubElement(parent, tag, attrs)
    if text is not None:
        element.text = text
    return element


def _extended_data(placemark, sidc):
    extended = _sub(placemark, "ExtendedData")
    data = _sub(extended, "Data", name="sidc")
    _sub(data, "value", sidc)


def build_kml(points, lines, icon_sidcs, line_color="#00C0FF"):
    kml = ET.Element("kml", xmlns=KML_NS)
    document = _sub(kml, "Document")

    for sidc in sorted(icon_sidcs):
        style = _sub(document, "Style", id=f"sidc-{sidc}")
        icon_style = _sub(style, "IconStyle")
        _sub(icon_style, "scale", ICON_SCALE)
        icon = _sub(icon_style, "Icon")
        _sub(icon, "href", f"files/{sidc}.png")
    line_style = _sub(_sub(document, "Style", id="line"), "LineStyle")
    _sub(line_style, "color", kml_color(line_color))
    _sub(line_style, "width", LINE_WIDTH)

    for line in lines:
        placemark = _sub(document, "Placemark", id=line.uid)
        _sub(placemark, "name", line.name or "-")
        _sub(placemark, "visibility", "1")
        _sub(placemark, "description", line.description)
        _sub(placemark, "styleUrl", "#line")
        _extended_data(placemark, line.sidc)
        geometry = _sub(placemark, "LineString")
        _sub(geometry, "tessellate", "1")
        _sub(geometry, "coordinates", " ".join(f"{format_coordinate(lon)},{format_coordinate(lat)}" for lat, lon in line.coords))

    for point in points:
        placemark = _sub(document, "Placemark", id=point.uid)
        _sub(placemark, "name", point.name)
        _sub(placemark, "visibility", "1")
        _sub(placemark, "description", point.description)
        if point.sidc in icon_sidcs:
            _sub(placemark, "styleUrl", f"#sidc-{point.sidc}")
        _extended_data(placemark, point.sidc)
        geometry = _sub(placemark, "Point")
        _sub(geometry, "coordinates", f"{format_coordinate(point.lon)},{format_coordinate(point.lat)}")

    ET.indent(kml, space="    ")
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(kml, encoding="unicode")


def write_kmz(path, points, lines, icons, line_color="#00C0FF"):
    """icons - {sidc: PNG-байти}; у KMZ потрапляють лише іконки точок, які є на карті."""
    used_icons = {sidc: png for sidc, png in icons.items() if any(p.sidc == sidc for p in points)}
    kml = build_kml(points, lines, set(used_icons), line_color)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("doc.kml", kml.encode("utf-8"))
        for sidc, png in sorted(used_icons.items()):
            archive.writestr(f"files/{sidc}.png", png)
