"""Читання resources/addons/book61.xlsm - окремої, значно повнішої особової
книги підрозділу (21 аркуш, ведеться незалежно й вручну, НЕ через щоденні
рапорти/information_unit) - джерело "очікуваних періодів" для звірки ЩОЙНО
згенерованого ОБЛІК.xlsx (content/addon_mismatch_checker.py). Детальна
інструкція по кожному статусу - resources/ОБЛІК.xlsx, аркуш "map"
(чернетка користувача, підтверджена й доповнена в розмові 2026-10-18 -
конкретні уточнення позначені в докстрінгах нижче).

За прямою вказівкою користувача - статуси, яких book61 ЩЕ не охоплює
(аркуш "map": "не має поки що файлу та даних" чи "не підається
автоматизації") - Задув, Аршт, БЗВП, Адап, НОВ, ЗВФН, ЛМР, Інт, КФАХ - тут
НЕ читаються взагалі, НЕ через недогляд.

КОЖЕН аркуш нижче (Алф2/ВП/ВД/ШП/ВЛК/ППД/СЗЧ/Втрати/Прибулі у ВД) має
рядок 1 = текстові заголовки колонок, рядок 2 = суто числові індекси
колонок (внутрішня службова нумерація самої книги, не дані) - дані
завжди починаються з рядка 3 (_DATA_START_ROW)."""

import re
from datetime import date, datetime, timedelta

from constants import TRANSFER_TO_BRIGADE_STATUS, TRANSFER_TO_OTHER_UNIT_STATUS
from content.excel_compat import load_workbook
from content.oblik_timesheet import normalize_name

_DATA_START_ROW = 3

# Книга book61 використовує "01.01.1990" як типову заглушку "дата ще не
# встановлена" (підтверджено користувачем, map рядок 11 - "але не 1990") -
# лише для аркуша "Прибулі у ВД" (read_rrez_tvo_periods нижче); інші
# аркуші цієї заглушки не використовують.
_PLACEHOLDER_DATE = date(1990, 1, 1)

# Очікуваний статус "будь-який" - статус бере Табель (content/
# addon_mismatch_checker.py вважає збігом що завгодно в Табелі).
ANY_STATUS = "*"


def load_addon_workbook(path):
    """Будь-яка книга resources/addons (book61 чи ЦВС - content/cvs_reader.py).
    data_only=True (значення формул, не самі формули) - той самий підхід,
    що й information_unit_reader/payment_mismatch_checker для сторонніх
    книг Excel цього проєкту."""
    return load_workbook(path, data_only=True)


# Аркуші, які відкривають читачі нижче - книга без будь-якого з них не
# book61 (generators/check_addon_mismatches.py розпізнає тип книги за цим).
BOOK61_SHEETS = ("ВП", "ВД", "ШП", "ВЛК", "ППД", "Алф2", "СЗЧ", "Втрати", "Прибулі у ВД")


def is_book61_workbook(wb):
    return all(name in wb.sheetnames for name in BOOK61_SHEETS)


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


# Колонка "кінця" period-запису (ВД/ВП "Дата прибуття"/"Дата 08 ранку", ШП
# "Кінець (виписка)", ППД "Дата вибуття з ППД") - реальні випадки на
# book61.xlsm: ЧАСТО вільний текст із вбудованою датою ("у ППД з
# 18.07.2026", "у СЗЧ з 09.03.2026", "надходження 27.06.2026, виписка
# 15.07.2026"), а не чиста дата-комірка. БЕЗ цього - _as_date повертає
# None на такий текст, кінець періоду лишається НЕВИЗНАЧЕНИМ, і "відкритий"
# період (end=None, content/addon_mismatch_checker.py) тягнеться аж до
# ОСТАННЬОЇ колонки-дати ПОТОЧНОГО Табеля - реальний випадок, підтверджений
# користувачем: ВД "12.07.2026-18.07.2026" ("Дата прибуття" - саме такий
# вільний текст) хибно перевірявся аж до ЖОВТНЯ, хоча сама поїздка
# завершилась у ЛИПНІ.
#
# ОСТАННЯ (а не перша) дата в тексті - за прямою вказівкою задуму: "виписка"
# (ШП) завжди йде ПІСЛЯ "надходження" в тому самому тексті, а для "у ВЛК з
# ..., ... прибув" - друга дата - це фактичне ПОВЕРНЕННЯ, новіша інформація,
# ніж перша. Текст без жодної дати ("продовжено", "до окремого
# розпорядження", "ТВО", "-") - period лишається ВІДКРИТИМ (None), як і
# раніше - це НАВМИСНО (дійсно немає конкретної дати завершення в book61).
_EMBEDDED_DATE_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})")


def _end_date_from_cell(value):
    direct = _as_date(value)
    if direct is not None:
        return direct
    if not isinstance(value, str):
        return None
    matches = _EMBEDDED_DATE_RE.findall(value)
    if not matches:
        return None
    try:
        return datetime.strptime(matches[-1], "%d.%m.%Y").date()
    except ValueError:
        return None


# "у <статус> з <дата>" ("у вдр з 09.10.2026", "у СЗЧ з 12.03.2026") - дата
# ПОЧАТКУ наступного статусу, тож сам період закінчився напередодні.
# Перевірено на реальних даних: у ці дати Табель уже показує новий статус.
_TRANSITION_TEXT_RE = re.compile(r"^\s*у\s+\S+\s+(?:з\s+)?\d{2}\.\d{2}\.\d{4}", re.IGNORECASE)


def _is_transition_text(value):
    return isinstance(value, str) and _TRANSITION_TEXT_RE.match(value) is not None


def _header_map(ws):
    """{назва колонки з рядка 1: номер колонки} - лише непорожні заголовки."""
    headers = {}
    for c in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=c).value
        if value:
            headers[value] = c
    return headers


def _read_period_records(wb, sheet_name, *, pib_header, start_header, status_fn,
                          end_header=None, fallback_end_header=None, exclusive_end=False):
    """Один рядок аркаша (з рядка 3) -> period-запис {"pib_raw", "status",
    "start_date", "end_date", "source_sheet"} - спільна основа для ВСІХ
    "період = просто дати" сімей (ВП/ВПС/ВПСЗ, ВД/Навч, ШП/БПШП, ВЛК, ППД).
    end_date (_end_date_from_cell - чиста дата-комірка АБО дата, вбудована у
    вільний текст, напр. "у ППД з 18.07.2026") - None означає ВІДКРИТИЙ
    період (ще триває чи дійсно немає даних про завершення - content/
    addon_mismatch_checker.py перевіряє аж до останньої колонки-дати
    Табеля).

    status_fn(row) -> статус ОБЛІК для ЦЬОГО рядка (row - {назва колонки:
    значення клітинки}), або None, щоб рядок ПРОПУСТИТИ - дозволяє ОДНІЙ й
    тій самій функції обслуговувати і "увесь аркуш - один статус" (ВЛК/ППД:
    lambda row: "ВЛК"), і "статус залежить від колонки" (ВП: "Тип
    відпустки"; ВД: "Виплата"; ШП: "БЧС") випадки без дублювання циклу.

    exclusive_end=True - дата кінця в book61 - це ПЕРШИЙ день ПІСЛЯ періоду
    (людина вже повернулась чи перейшла в інший статус), тож останній день
    самого періоду - на день раніше. Для решти аркушів так само - лише коли
    кінець записано текстом "у <статус> з <дата>" (_is_transition_text).

    "arrival_date" - день прибуття (повернення): кінець періоду, записаний
    датою БЕЗ такого зсуву; None - якщо кінець зсунуто чи період відкритий
    (content/addon_mismatch_checker.py, ADDON_ARRIVAL_DAY_STATUSES)."""
    ws = wb[sheet_name]
    headers = _header_map(ws)
    pib_col = headers[pib_header]
    start_col = headers[start_header]
    end_col = headers.get(end_header) if end_header else None
    fallback_col = headers.get(fallback_end_header) if fallback_end_header else None

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        row = {name: ws.cell(row=row_idx, column=c).value for name, c in headers.items()}
        status = status_fn(row)
        if status is None:
            continue
        start = _as_date(row.get(start_header))
        if start is None:
            continue
        end_raw = row.get(end_header) if end_col else None
        end = _end_date_from_cell(end_raw)
        if end is None and fallback_col:
            end_raw = row.get(fallback_end_header)
            end = _end_date_from_cell(end_raw)
        shifted = end is not None and (exclusive_end or _is_transition_text(end_raw))
        if shifted:
            end -= timedelta(days=1)
        records.append({
            "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
            "source_sheet": sheet_name, "arrival_date": None if shifted else end,
        })
    return records


# "Тип відпустки" на аркуші "ВП" - значення, підтверджені напряму з
# реальних даних book61: "Перша/Друга частина щорічної відпустки" і дві
# значення, яких НЕМАЄ в аркуші map ("По догляду за дитиною до досягнення
# нею трирічного віку", "Соціальна відпустка") - за прямою вказівкою
# користувача, УСІ ці чотири різновиди зводяться до звичайного "ВП" (мапа
# нижче НЕ згадує їх явно - вони просто не підпадають під жоден із двох
# спеціальніших випадків і лишаються на дефолтному "ВП").
_VP_TYPE_TO_STATUS = {
    "За станом здоров'я": "ВПСЗ",
    "За сімейними обставинами": "ВПС",
}


def read_vp_family_periods(wb):
    """ВП/ВПС/ВПСЗ (map рядки 7-9) - аркуш "ВП": період - "Дата вибуття" ->
    "Дата прибуття", якщо відсутня - "Дата 08 ранку" (map, дослівно), МІНУС
    один день (exclusive_end): за прямою вказівкою користувача "Дата 08
    ранку" - це ранок, коли людина вже на місці; на реальних даних "Дата
    прибуття" здебільшого та сама дата (чи дата переходу в інший статус,
    "у лік з ..."), і Табель у цей день теж уже не "ВП" - без зсуву
    останній день КОЖНОЇ відпустки хибно рахувався розбіжністю.
    Статус - за колонкою "Тип відпустки" (_VP_TYPE_TO_STATUS)."""
    return _read_period_records(
        wb, "ВП", pib_header="П.І.Б.", start_header="Дата вибуття",
        end_header="Дата прибуття", fallback_end_header="Дата 08 ранку", exclusive_end=True,
        status_fn=lambda row: _VP_TYPE_TO_STATUS.get(row.get("Тип відпустки"), "ВП"),
    )


_NOT_DEPARTED_RE = re.compile(r"не\s+вибу(в|ла|ли)\b")


def read_vd_periods(wb):
    """ВД/Навч (map рядки 12+15) - аркуш "ВД": той самий період, що й ВП
    (вибуття -> прибуття, інакше "08 ранку"), АЛЕ статус залежить від
    колонки "Виплата" - "10" (чи порожньо) означає "ВД" дійсно
    підтверджено; текст "навчання" означає, що ОБЛІК МАЄ казати "Навч", а
    не "ВД" (та сама логіка, що й content/daily_report_reader.py::
    _is_training_destination, лише джерело істини тут - book61, а не сам
    текст рапорту). "не вибув"/"не вибула" - за прямою вказівкою
    користувача людина фактично не вибула, тож статус бере Табель
    (ANY_STATUS) - розбіжності на ці дати не буває."""
    def _status(row):
        payment = row.get("Виплата")
        if isinstance(payment, str):
            text = " ".join(payment.split()).lower()
            if text == "навчання":
                return "Навч"
            if _NOT_DEPARTED_RE.match(text):
                return ANY_STATUS
        return "ВД"

    return _read_period_records(
        wb, "ВД", pib_header="П.І.Б.", start_header="Дата вибуття",
        end_header="Дата прибуття", fallback_end_header="Дата 08 ранку", status_fn=_status,
    )


# "БПШП"/"ВПБП" (map рядки 22 і 10) - ОБИДВА за підтвердженням користувача
# йдуть із ТОГО САМОГО аркуша "ШП", фільтр БЧС="ЛП" (рядок 10 ("ВП АРКУШ")
# - описка в map, підтверджено користувачем: book61 НЕ розрізняє ці два
# статуси на рівні даних) - тут виробляється лише "БПШП" (та сама, що й
# реальні рядки book61), а "ВПБП" трактується як рівнозначний варіант під
# час самої звірки (content/addon_mismatch_checker.py), а не як окремий,
# ЩЕ ОДИН period-запис на ТІ САМІ рядки/дати.
_SHP_BCHS_TO_STATUS = {"лік": "ШП", "ЛП": "БПШП"}


def read_shp_family_periods(wb):
    """ШП/БПШП (map рядки 20+22) - аркуш "ШП": період - "Дата вибуття" ->
    "Кінець (виписка)" (відкритий, якщо відсутня). Статус - за колонкою
    "БЧС" (_SHP_BCHS_TO_STATUS) - рядки з БУДЬ-ЯКИМ іншим значенням БЧС не
    стосуються жодного з цих двох статусів, пропускаються."""
    return _read_period_records(
        wb, "ШП", pib_header="П.І.Б.", start_header="Дата вибуття", end_header="Кінець (виписка)",
        status_fn=lambda row: _SHP_BCHS_TO_STATUS.get(row.get("БЧС")),
    )


def read_vlk_periods(wb):
    """ВЛК (map рядок 3) - аркуш "ВЛК": період - "Дата вибуття" -> "Кінець
    (ВЛК)" (відкритий, якщо відсутня)."""
    return _read_period_records(
        wb, "ВЛК", pib_header="П.І.Б.", start_header="Дата вибуття", end_header="Кінець (ВЛК)",
        status_fn=lambda row: "ВЛК",
    )


def read_ppd_periods(wb):
    """ППД (map рядок 14) - аркуш "ППД": період - "Дата прибуття на ППД" ->
    день ПЕРЕД "Дата вибуття з ППД" (відкритий, якщо відсутня - людина й
    далі на ППД). День вибуття - уже перший день наступного статусу: так
    само, як "Дата вибуття" на аркушах ВП/ВД/ШП/ВЛК - це ПЕРШИЙ день
    відпустки/відрядження/лікування; перевірено на реальних даних - Табель
    у цей день уже показує новий статус."""
    return _read_period_records(
        wb, "ППД", pib_header="П.І.Б.", start_header="Дата прибуття на ППД", end_header="Дата вибуття з ППД",
        status_fn=lambda row: "ППД", exclusive_end=True,
    )


# Дата, вбудована у ВІЛЬНИЙ текст колонки "Дата, наказ розпорядження,
# виключення" аркуша "Алф2" - реальний випадок, підтверджений напряму з
# даних: "Виключений із списків 31.10.2023" (а не чиста дата-комірка).
_ORDER_DATE_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})")


def _order_date_from_text(value):
    if not isinstance(value, str):
        return None
    match = _ORDER_DATE_RE.search(value)
    if match is None:
        return None
    try:
        return datetime.strptime(match.group(1), "%d.%m.%Y").date()
    except ValueError:
        return None


# Колонка "СТАТУС" аркуша "Алф2" -> ОБЛІК-статус (map рядки 6, 24, 29) -
# лише значення, для яких є ОДНОЗНАЧНИЙ відповідник серед 30 рядків map.
# "переведений в середині бригади" - "ПРВДБ", "переведений в іншу в/ч" -
# "ПРВДВЧ" (за прямою вказівкою користувача - замість колишнього "ПРВД").
# "самостріл" - РЕАЛЬНЕ значення цієї колонки, але ЖОДЕН із 30 рядків
# map не називає, якому статусу ОБЛІК воно відповідає - за прямою
# вказівкою користувача ("поправ якщо є уточнення" - не вгадувати) сюди
# НЕ додається; _read_unmapped_alf2_statuses нижче окремо позначає таких
# людей попередженням у звіті, а не мовчки ігнорує чи хибно вгадує.
_ALF2_STATUS_TO_OBLIK = {
    "звільнений": "Звіль",
    "переведений в середині бригади": TRANSFER_TO_BRIGADE_STATUS,
    "переведений в іншу в/ч": TRANSFER_TO_OTHER_UNIT_STATUS,
    "полон": "полон",
}

# Значення колонки "СТАТУС" на Алф2, яким НЕ відповідає жоден із 30
# рядків аркуша map - "самостріл" підтверджено напряму з реальних даних;
# "ВД"/"ВП"/"СЗЧ"/"ШП" тут НЕ зайві (вони вже є - ці чотири статуси мають
# своє ВЛАСНЕ, набагато детальніше джерело в однойменних аркушах book61,
# тож Алф2 для них навмисно НЕ використовується, а не просто "забуто").
_ALF2_UNMAPPED_STATUSES = ("самостріл",)


def read_alf2_order_periods(wb):
    """Звіль/ПРВДБ/ПРВДВЧ/полон (map рядки 6, 29, 24) - аркуш "Алф2": колонка
    "СТАТУС" визначає ОБЛІК-статус (_ALF2_STATUS_TO_OBLIK), дата - з
    вільного тексту колонки "Дата, наказ розпорядження, виключення"
    (_order_date_from_text). Період - ЗАВЖДИ ВІДКРИТИЙ (адміністративний
    статус, "з моменту цієї дати включно" - map, дослівно - без природного
    завершення в самій книзі)."""
    ws = wb["Алф2"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б"]
    status_col = headers["СТАТУС"]
    order_col = headers["Дата, наказ розпорядження, виключення"]

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        oblik_status = _ALF2_STATUS_TO_OBLIK.get(ws.cell(row=row_idx, column=status_col).value)
        if oblik_status is None:
            continue
        start = _order_date_from_text(ws.cell(row=row_idx, column=order_col).value)
        if start is None:
            continue
        records.append({
            "pib_raw": pib_raw, "status": oblik_status, "start_date": start, "end_date": None,
            "source_sheet": "Алф2",
        })
    return records


def read_szch_to_rozp_periods(wb):
    """СЗЧ -> Розп (map рядки 2+5) - ДВА незалежні, послідовні періоди для
    однієї людини:
    (1) аркуш "СЗЧ", "Дата події" -> ОЧІКУЄТЬСЯ "СЗЧ" до дня перед "Дата
    повернення з СЗЧ" (у день повернення людина вже на місці - той самий
    принцип, що й "Дата прибуття" на ВП); без дати повернення - відкритий.
    Реальний випадок, підтверджений користувачем: СЗЧ 28.02-01.03.2026 без
    урахування дати повернення "тягнувся" аж у жовтень;
    (2) ЯКЩО ТА САМА людина ТЕЖ має рядок на "Алф2" з непорожньою "Дата,
    наказ розпорядження, виключення" (сам факт заповненої дати наказу -
    сигнал, незалежно від того, що саме каже колонка СТАТУС там) ->
    ОЧІКУЄТЬСЯ "Розп" від ЦІЄЇ дати.
    Два записи природно "складаються" під час звірки по датах
    (content/addon_mismatch_checker.py) - період (2), якщо є, фактично
    перекриває "хвіст" відкритого періоду (1) починаючи з дати наказу.

    "Розп"-запис навмисно має ІНШИЙ source_sheet ("Алф2 (Розп)", а не
    просто "Алф2") - інакше під час групування (content/addon_mismatch_
    checker.py) він змішався б із "Звіль"/"ПРВДБ"/"полон" цієї ж людини в
    ОДИН рядок звіту."""
    szch_ws = wb["СЗЧ"]
    szch_headers = _header_map(szch_ws)
    pib_col = szch_headers["П.І.Б."]
    event_col = szch_headers["Дата події"]
    return_col = szch_headers["Дата повернення з СЗЧ"]

    records = []
    szch_normalized = set()
    for row_idx in range(_DATA_START_ROW, szch_ws.max_row + 1):
        pib_raw = szch_ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        start = _as_date(szch_ws.cell(row=row_idx, column=event_col).value)
        if start is None:
            continue
        returned = _end_date_from_cell(szch_ws.cell(row=row_idx, column=return_col).value)
        records.append({
            "pib_raw": pib_raw, "status": "СЗЧ", "start_date": start,
            "end_date": returned - timedelta(days=1) if returned else None,
            "source_sheet": "СЗЧ",
        })
        szch_normalized.add(normalize_name(pib_raw))

    alf2_ws = wb["Алф2"]
    alf2_headers = _header_map(alf2_ws)
    alf2_pib_col = alf2_headers["П.І.Б"]
    order_col = alf2_headers["Дата, наказ розпорядження, виключення"]
    for row_idx in range(_DATA_START_ROW, alf2_ws.max_row + 1):
        pib_raw = alf2_ws.cell(row=row_idx, column=alf2_pib_col).value
        if not pib_raw or normalize_name(pib_raw) not in szch_normalized:
            continue
        order_date = _order_date_from_text(alf2_ws.cell(row=row_idx, column=order_col).value)
        if order_date is None:
            continue
        records.append({
            "pib_raw": pib_raw, "status": "Розп", "start_date": order_date, "end_date": None,
            "source_sheet": "Алф2 (Розп)",
        })
    return records


def read_unmapped_alf2_people(wb):
    """Люди з Алф2, чий СТАТУС - одне зі значень _ALF2_UNMAPPED_STATUSES
    ("самостріл") - жоден із 30 рядків аркуша map не називає, якому
    статусу ОБЛІК це відповідає. За прямою вказівкою користувача ("поправ
    якщо є уточнення") - НЕ вгадується мовчки: повертає [{"pib_raw",
    "alf2_status"}] для окремого попередження в звіті
    (content/addon_mismatch_checker.py), а не для порівняння по датах."""
    ws = wb["Алф2"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б"]
    status_col = headers["СТАТУС"]

    people = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        status = ws.cell(row=row_idx, column=status_col).value
        if status in _ALF2_UNMAPPED_STATUSES:
            people.append({"pib_raw": pib_raw, "alf2_status": status})
    return people


# "БЧС" аркуша "Втрати" -> ОБЛІК-статус (map рядки 23, 26, 27). "смерть" -
# підтверджено логікою map, хоча НАРАЗІ немає жодного реального рядка з
# таким значенням у book61 (лише "загибель"/"безвісті" станом на
# 2026-10-18) - готово спрацювати, щойно з'явиться.
_VTRATY_BCHS_TO_OBLIK = {
    "загибель": "Загиб",
    "смерть": "Смерть",
    "безвісті": "БЗ",
}


def read_vtraty_periods(wb):
    """Загиб/Смерть/БЗ (map рядки 26, 27, 23) - аркуш "Втрати": період -
    "Дата події", ЗАВЖДИ ВІДКРИТИЙ (термінальний статус, не "завершується"
    в book61)."""
    ws = wb["Втрати"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б."]
    bchs_col = headers["БЧС"]
    event_col = headers["Дата події"]

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        oblik_status = _VTRATY_BCHS_TO_OBLIK.get(ws.cell(row=row_idx, column=bchs_col).value)
        if oblik_status is None:
            continue
        start = _as_date(ws.cell(row=row_idx, column=event_col).value)
        if start is None:
            continue
        records.append({
            "pib_raw": pib_raw, "status": oblik_status, "start_date": start, "end_date": None,
            "source_sheet": "Втрати",
        })
    return records


def read_rrez_tvo_periods(wb):
    """Аркуш "Прибулі у ВД" - за прямою вказівкою користувача, процес
    прибуття людини, три послідовні кроки:
    - "Дата допущення до ТВО" .. день перед "Дата зарахування в списки" -
      "РРез" (ТВО, не оплачується);
    - "Дата зарахування в списки" .. день перед "Дата прийому посади" -
      лише "ППД" (сам день зарахування - ЗАВЖДИ "ППД");
    - "Дата прийому посади" - будь-який статус (ANY_STATUS).
    Без "Дата зарахування в списки" - "РРез" аж до дня перед прийомом
    посади. Рядок без "Дата прийому посади" (чи з заглушкою
    _PLACEHOLDER_DATE) пропускається - посаду ще не прийнято, судити зарано.

    На свої дні цей аркуш ПЕРЕВАЖАЄ будь-який паралельний book61-період
    (реальний випадок: людина ОДНОЧАСНО є й на аркуші "ППД" з дня допущення
    до ТВО) - content/addon_mismatch_checker.py (_PRIORITY_SOURCE_SHEETS).

    Попередня редакція (read_rrez_people) трактувала ВЕСЬ аркуш як
    ПЕРЕВІРКУ ІСНУВАННЯ ("ця людина не має фігурувати в Табелі взагалі") -
    за прямою вказівкою користувача ВИДАЛЕНО: це ЛЕГІТИМНІ, щойно прибулі
    військовослужбовці."""
    ws = wb["Прибулі у ВД"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б."]

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        row = {name: ws.cell(row=row_idx, column=c).value for name, c in headers.items()}
        tvo = _as_date(row.get("Дата допущення до ТВО"))
        position = _as_date(row.get("Дата прийому посади"))
        if tvo is None or position is None or position == _PLACEHOLDER_DATE:
            continue
        enrollment = _as_date(row.get("Дата зарахування в списки"))
        if enrollment == _PLACEHOLDER_DATE:
            enrollment = None

        steps = [("РРез", tvo, (enrollment or position) - timedelta(days=1))]
        if enrollment is not None:
            steps.append(("ППД", enrollment, max(enrollment, position - timedelta(days=1))))
        steps.append((ANY_STATUS, position, position))
        for status, start, end in steps:
            if start <= end:
                records.append({
                    "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
                    "source_sheet": "Прибулі у ВД",
                })
    return records


def read_all_period_records(wb, source_file):
    """Усі "період -> очікуваний статус" записи РАЗОМ (усі сім'ї вище,
    КРІМ read_unmapped_alf2_people - та має ІНШУ форму запису, звіряється
    окремо, content/addon_mismatch_checker.py). Записи можуть
    перекриватись у часі (кожен аркуш book61 ведеться окремо) - ОДНУ
    послідовну шкалу на людину з них будує content/addon_mismatch_checker.py
    (_paint_timeline).

    source_file ("book61" - ім'я файлу без розширення) - на КОЖНОМУ записі:
    за прямою вказівкою користувача в resources/addons з часом будуть й інші
    файли, і їхні аркуші мають іти в ТОЙ САМИЙ рядок звіту людини, кожен зі
    своїм префіксом ("book61: ППД -> <інший файл>: ...")."""
    records = []
    for reader in (
        read_vp_family_periods, read_vd_periods, read_shp_family_periods, read_vlk_periods,
        read_ppd_periods, read_alf2_order_periods, read_szch_to_rozp_periods, read_vtraty_periods,
        read_rrez_tvo_periods,
    ):
        records.extend(reader(wb))
    for record in records:
        record["source_file"] = source_file
    return records
