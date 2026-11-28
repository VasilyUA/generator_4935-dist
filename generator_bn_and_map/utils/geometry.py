"""Закруглення ліній карти (KMZ для Кропиви і CSV для Дельти): злами ламаної
замінюються плавними дугами (кубічний сплайн Ерміта), які проходять ТОЧНО через
ті самі вершини - координати з ПБД, штабних рішень чи карти Дельти лишаються на
лінії, додаються лише проміжні точки дуг.

- Закруглюються лише злами, гостріші за SHARP_TURN_DEGREES.
- Наскільки широко закруглюється злам - radius (метри, але не більше за
  коротший із двох сусідніх відрізків). Лінія відходить від ламаної з документа
  не більше ніж приблизно на чверть radius (≈120 м при 500 м).
- Дуга ділиться так, щоб сусідні відрізки повертали не більше ніж на
  2 * STEP_DEGREES (< SHARP_TURN_DEGREES). Тому вже закруглена лінія (зокрема
  та, що повернулася з Дельти після імпорту) при наступному запуску лишається
  без змін - точки не множаться.
- Розворот майже назад (понад CUSP_DEGREES) і злам біля відрізка, коротшого за
  _MIN_REACH_METERS, лишаються гострими: дуга там намалювала б петлю, якої на
  карті не було, або була б непомітною."""
import math

SHARP_TURN_DEGREES = 20.0
STEP_DEGREES = 8.0
CUSP_DEGREES = 150.0
DEFAULT_RADIUS_METERS = 500.0
_MIN_REACH_METERS = 10.0
_MIN_PIECE_METERS = 1.0
_MAX_DEPTH = 16
_METERS_PER_DEGREE = 111320.0


def _dedupe(points):
    result = []
    for point in points:
        if not result or point != result[-1]:
            result.append(point)
    return result


def _unit(dx, dy):
    length = math.hypot(dx, dy)
    return (dx / length, dy / length) if length else (0.0, 0.0)


def _angle(a, b):
    """Кут між напрямками a і b (одиничні вектори), градуси 0…180."""
    return math.degrees(math.atan2(abs(a[0] * b[1] - a[1] * b[0]), a[0] * b[0] + a[1] * b[1]))


class _Span:
    """Кубічна крива Ерміта p→q: виходить з p у напрямку d0, входить у q у
    напрямку d1; m0, m1 - довжини дотичних (наскільки далеко сягає вигин)."""

    def __init__(self, p, q, d0, d1, m0, m1):
        self.p, self.q = p, q
        self.v0 = (d0[0] * m0, d0[1] * m0)
        self.v1 = (d1[0] * m1, d1[1] * m1)

    def point(self, t):
        t2 = t * t
        t3 = t2 * t
        h00, h10, h01, h11 = 2 * t3 - 3 * t2 + 1, t3 - 2 * t2 + t, 3 * t2 - 2 * t3, t3 - t2
        return tuple(h00 * self.p[k] + h10 * self.v0[k] + h01 * self.q[k] + h11 * self.v1[k] for k in (0, 1))

    def direction(self, t):
        t2 = t * t
        g00, g10, g01, g11 = 6 * t2 - 6 * t, 3 * t2 - 4 * t + 1, 6 * t - 6 * t2, 3 * t2 - 2 * t
        return _unit(*(g00 * self.p[k] + g10 * self.v0[k] + g01 * self.q[k] + g11 * self.v1[k] for k in (0, 1)))

    def inner_points(self):
        """Проміжні точки: інтервал ділиться навпіл, доки дотична на ньому
        повертає більше ніж на STEP_DEGREES (тоді сусідні відрізки ламаної
        повертають не більше ніж на 2 * STEP_DEGREES)."""
        result = []

        def split(t0, t1, depth):
            tm = (t0 + t1) / 2
            a, m, b = self.direction(t0), self.direction(tm), self.direction(t1)
            p0, p1 = self.point(t0), self.point(t1)
            if (depth >= _MAX_DEPTH or _angle(a, m) + _angle(m, b) <= STEP_DEGREES
                    or math.hypot(p1[0] - p0[0], p1[1] - p0[1]) < 2 * _MIN_PIECE_METERS):
                return
            split(t0, tm, depth + 1)
            result.append(self.point(tm))
            split(tm, t1, depth + 1)

        split(0.0, 1.0, 0)
        return result


def _lead(vertex, direction, nearest, reach):
    """Точка на відрізку біля вершини, кут у якій лишається гострим: перша
    частина дуги лягає точно на відрізок, тож кут не змінюється (і повторний
    запуск не прийме його за інший)."""
    distance = min(math.hypot(nearest[0] - vertex[0], nearest[1] - vertex[1]) / 2, reach / 10)
    return (vertex[0] + direction[0] * distance, vertex[1] + direction[1] * distance)


def smooth_polyline(coords, closed=None, radius=DEFAULT_RADIUS_METERS):
    """[(lat, lon), …] -> закруглена лінія: ті самі вершини (ті самі числа) плюс
    проміжні точки дуг біля зламів. closed=None - визначити за лінією (перша
    точка дорівнює останній); radius - ширина закруглення в метрах (0/None - аж
    до сусідніх вершин). Лінію без гострих зламів повертає без змін."""
    points = _dedupe([(c[0], c[1]) for c in coords])
    if closed is None:
        closed = len(points) > 3 and points[0] == points[-1]
    if closed and len(points) > 1 and points[0] == points[-1]:
        points = points[:-1]
    n = len(points)
    if n < 3:
        return list(coords)
    # Локальна площина в метрах: довгота множиться на cos(широти), щоб кути не спотворювались.
    lat0, lon0 = sum(p[0] for p in points) / n, sum(p[1] for p in points) / n
    kx, ky = _METERS_PER_DEGREE * math.cos(math.radians(lat0)), _METERS_PER_DEGREE
    xy = [((lon - lon0) * kx, (lat - lat0) * ky) for lat, lon in points]
    spans = n if closed else n - 1
    ends = [(i, (i + 1) % n) for i in range(spans)]
    lengths = [math.hypot(xy[j][0] - xy[i][0], xy[j][1] - xy[i][1]) for i, j in ends]
    chords = [_unit(xy[j][0] - xy[i][0], xy[j][1] - xy[i][1]) for i, j in ends]

    def sides(v):
        """(вхідний, вихідний) відрізок вершини v; None - кінець відкритої лінії."""
        if closed:
            return (v - 1) % n, v
        return (v - 1 if v > 0 else None), (v if v < n - 1 else None)

    turns, rounded = [], []
    for v in range(n):
        a, b = sides(v)
        if a is None or b is None:
            turns.append(0.0)
            rounded.append(False)
            continue
        turns.append(_angle(chords[a], chords[b]))
        rounded.append(SHARP_TURN_DEGREES < turns[v] <= CUSP_DEGREES and min(lengths[a], lengths[b]) >= _MIN_REACH_METERS)
    curved = [rounded[i] or rounded[j] for i, j in ends]
    if not any(curved):
        return list(coords)

    def tangent(v, span):
        """(напрямок, довжина, гострий кут лишається) дотичної відрізка span у його кінці v."""
        a, b = sides(v)
        reach = min(lengths[s] for s in (a, b) if s is not None)
        if radius:
            reach = min(reach, radius)
        if a is None or b is None:
            return chords[span], reach, False
        if turns[v] > SHARP_TURN_DEGREES and not rounded[v]:
            # Розворот назад чи злам біля дуже короткого відрізка - кут лишається як був.
            return chords[span], reach, True
        other = a if span == b else b
        if not rounded[v] and not curved[other]:
            # Майже прямий кут поруч із прямим відрізком: дуга плавно переходить у нього.
            return chords[other], reach, False
        return _unit(chords[a][0] + chords[b][0], chords[a][1] + chords[b][1]), reach, False

    result = []
    for span, (i, j) in enumerate(ends):
        result.append(points[i])
        if not curved[span]:
            continue
        (d0, m0, keep0), (d1, m1, keep1) = tangent(i, span), tangent(j, span)
        inner = _Span(xy[i], xy[j], d0, d1, m0, m1).inner_points()
        if inner and keep0:
            inner.insert(0, _lead(xy[i], d0, inner[0], m0))
        if inner and keep1:
            inner.append(_lead(xy[j], (-d1[0], -d1[1]), inner[-1], m1))
        for x, y in inner:
            result.append((lat0 + y / ky, lon0 + x / kx))
    result.append(points[0] if closed else points[-1])
    return result
