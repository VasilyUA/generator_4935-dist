"""Значки для точок KMZ, для яких немає справжньої іконки.

Свої повітряні об'єкти (набір знаків 01 - ТЗ бомберів, мавіків, FPV) малюються
повністю, як у Кропиві: блакитний купол, значок і літери модифікаторів.

Решту програма намалювати не вміє, тож замість порожньої рамки (яка виглядала б
як справжній знак) малює рамку за належністю з коду SIDC (4-та цифра; кольори
APP-6: своє - блакитний прямокутник, своя техніка - коло, противник - червоний
ромб, нейтральне - зелений квадрат, невідоме - жовте коло) зі знаком «?»: на
карті одразу видно, що знак треба уточнити, а список таких точок програма
виводить окремо.

PNG пишеться без сторонніх бібліотек (zlib + struct)."""
import struct
import zlib

SIZE = 32
AIR_SIZE = 64
_BORDER = (0, 0, 0, 255)
_TRANSPARENT = (0, 0, 0, 0)
FRIEND, HOSTILE, NEUTRAL, UNKNOWN = "friend", "hostile", "neutral", "unknown"
_FILL = {
    FRIEND: (128, 224, 255, 255),
    HOSTILE: (255, 128, 128, 255),
    NEUTRAL: (170, 255, 170, 255),
    UNKNOWN: (255, 255, 128, 255),
}


def affiliation(sidc):
    """4-та цифра SIDC (standard identity) -> friend / hostile / neutral / unknown."""
    digit = (sidc or "")[3:4]
    if digit in ("2", "3"):
        return FRIEND
    if digit in ("5", "6"):
        return HOSTILE
    if digit == "4":
        return NEUTRAL
    return UNKNOWN


RECT, DIAMOND, SQUARE, CIRCLE, DOME = "rect", "diamond", "square", "circle", "dome"
_SHAPE = {FRIEND: RECT, HOSTILE: DIAMOND, NEUTRAL: SQUARE, UNKNOWN: CIRCLE}
_AIR_SYMBOL_SET = "01"
_EQUIPMENT_SYMBOL_SET = "15"

# APP-6D, набір знаків 01 (повітряні об'єкти): значки й літери модифікаторів.
_ROTARY_WING, _UNMANNED_AIRCRAFT = "110200", "110300"
_AIR_MODIFIER_1 = {"01": "A", "02": "B", "18": "R"}      # ударний, бомбардувальник, розвідник
_AIR_MODIFIER_2 = {"01": "H", "02": "M", "03": "L"}      # важкий, середній, легкий
_GLYPHS = {   # 5x7
    "A": ("01110", "10001", "10001", "11111", "10001", "10001", "10001"),
    "B": ("11110", "10001", "10001", "11110", "10001", "10001", "11110"),
    "H": ("10001", "10001", "10001", "11111", "10001", "10001", "10001"),
    "L": ("10000", "10000", "10000", "10000", "10000", "10000", "11111"),
    "M": ("10001", "11011", "10101", "10101", "10001", "10001", "10001"),
    "R": ("11110", "10001", "10001", "11110", "10100", "10010", "10001"),
    "?": ("01110", "10001", "00001", "00010", "00100", "00000", "00100"),
}


def frame_shape(sidc):
    """Форма рамки: свій повітряний об'єкт (набір знаків 01) - купол, своя
    техніка (набір 15 - екіпажі БпЛА, озброєння) - коло, як у Кропиві; решта -
    за належністю."""
    kind = affiliation(sidc)
    symbol_set = (sidc or "")[4:6]
    if kind == FRIEND and symbol_set == _AIR_SYMBOL_SET:
        return DOME
    if kind == FRIEND and symbol_set == _EQUIPMENT_SYMBOL_SET:
        return CIRCLE
    return _SHAPE[kind]


def _inside(shape, x, y, inset=0.0):
    """Чи піксель (x, y) усередині фігури shape, зменшеної на inset пікселів."""
    c = (SIZE - 1) / 2
    if shape == RECT:
        return 3 + inset <= x <= SIZE - 4 - inset and 8 + inset <= y <= SIZE - 9 - inset
    if shape == DIAMOND:
        return abs(x - c) + abs(y - c) <= 14 - inset * 1.4
    if shape == SQUARE:
        return 5 + inset <= x <= SIZE - 6 - inset and 5 + inset <= y <= SIZE - 6 - inset
    return (x - c) ** 2 + (y - c) ** 2 <= (13 - inset) ** 2


def _png(rows):
    def chunk(kind, data):
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    raw = b"".join(bytes([0]) + bytes(value for pixel in row for value in pixel) for row in rows)
    header = struct.pack(">IIBBBBB", len(rows[0]), len(rows), 8, 6, 0, 0, 0)
    signature = bytes([0x89]) + b"PNG" + bytes([13, 10, 26, 10])
    return signature + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


# --- свій повітряний об'єкт ------------------------------------------------------
_DOME_LEFT, _DOME_RIGHT, _DOME_TOP, _DOME_SHOULDER, _DOME_BOTTOM = 7, AIR_SIZE - 8, 5, 30, AIR_SIZE - 4


def _in_dome(x, y, inset):
    """Купол: півеліпс зверху, прямі боки, знизу відкритий (без лінії рамки)."""
    if y > _DOME_BOTTOM or not _DOME_LEFT + inset <= x <= _DOME_RIGHT - inset:
        return False
    if y >= _DOME_SHOULDER:
        return True
    rx = (_DOME_RIGHT - _DOME_LEFT) / 2 - inset
    ry = _DOME_SHOULDER - _DOME_TOP - inset
    dx = (x - (_DOME_LEFT + _DOME_RIGHT) / 2) / rx
    dy = (y - _DOME_SHOULDER) / ry
    return dx * dx + dy * dy <= 1.0


def _segment_distance(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return ((ax + t * dx - px) ** 2 + (ay + t * dy - py) ** 2) ** 0.5


def _in_triangle(px, py, a, b, c):
    def side(p, q):
        return (q[0] - p[0]) * (py - p[1]) - (q[1] - p[1]) * (px - p[0])
    s1, s2, s3 = side(a, b), side(b, c), side(c, a)
    return (s1 >= 0 and s2 >= 0 and s3 >= 0) or (s1 <= 0 and s2 <= 0 and s3 <= 0)


def _in_air_icon(entity, x, y):
    """БпЛА - товста «галочка»; гвинтокрил - «метелик» з двох трикутників."""
    cx, left, right = (_DOME_LEFT + _DOME_RIGHT) / 2, _DOME_LEFT + 4, _DOME_RIGHT - 4
    if entity == _UNMANNED_AIRCRAFT:
        return min(_segment_distance(x, y, left, 31, cx, 42), _segment_distance(x, y, cx, 42, right, 31)) <= 2.8
    if entity == _ROTARY_WING:
        middle = (cx, 37)
        return (_in_triangle(x, y, (left, 30), (left, 44), middle)
                or _in_triangle(x, y, (right, 30), (right, 44), middle))
    return False


def _draw_text(rows, text, top, scale=2):
    """Літери (5x7, збільшені в scale разів) по центру значка, від рядка top."""
    glyphs = [_GLYPHS[letter] for letter in text if letter in _GLYPHS]
    if not glyphs:
        return
    step = 6 * scale
    width = len(glyphs) * step - scale
    x0 = round((len(rows[0]) - 1) / 2 - width / 2)
    for index, glyph in enumerate(glyphs):
        for gy, line in enumerate(glyph):
            for gx, bit in enumerate(line):
                if bit == "1":
                    for dy in range(scale):
                        for dx in range(scale):
                            rows[top + scale * gy + dy][x0 + index * step + scale * gx + dx] = _BORDER


def _air_icon(sidc):
    entity = sidc[10:16]
    rows = []
    for y in range(AIR_SIZE):
        row = []
        for x in range(AIR_SIZE):
            if _in_dome(x, y, 3):
                row.append(_BORDER if _in_air_icon(entity, x, y) else _FILL[FRIEND])
            elif _in_dome(x, y, 0):
                row.append(_BORDER)
            else:
                row.append(_TRANSPARENT)
        rows.append(row)
    _draw_text(rows, _AIR_MODIFIER_1.get(sidc[16:18], ""), 10)
    _draw_text(rows, _AIR_MODIFIER_2.get(sidc[18:20], ""), 46)
    return _png(rows)


def is_full_symbol(sidc):
    """Чи fallback_icon малює повний знак (свій БпЛА / гвинтокрил), а не рамку зі «?»."""
    return frame_shape(sidc) == DOME and (sidc or "")[10:16] in (_ROTARY_WING, _UNMANNED_AIRCRAFT)


def _unknown_icon(sidc):
    """Рамка за належністю зі знаком «?» (64x64): знак невідомий, його треба уточнити."""
    kind, shape = affiliation(sidc), frame_shape(sidc)
    rows = []
    for y in range(AIR_SIZE):
        row = []
        for x in range(AIR_SIZE):
            if shape == DOME:
                inner, outer = _in_dome(x, y, 3), _in_dome(x, y, 0)
            else:
                inner, outer = _inside(shape, x / 2, y / 2, inset=1.5), _inside(shape, x / 2, y / 2)
            row.append(_FILL[kind] if inner else _BORDER if outer else _TRANSPARENT)
        rows.append(row)
    _draw_text(rows, "?", 21 if shape != DOME else 26, scale=3)
    return _png(rows)


def fallback_icon(sidc):
    """SIDC -> PNG 64x64: повний знак свого БпЛА / гвинтокрила або рамка зі «?»."""
    return _air_icon(sidc) if is_full_symbol(sidc) else _unknown_icon(sidc)
