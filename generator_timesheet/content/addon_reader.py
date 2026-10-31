"""Читання resources/addons/book61.xlsm - окремої й значно повнішої особової
книги підрозділу (21 аркуш, ведеться вручну, незалежно від щоденних рапортів
та information_unit). Це джерело очікуваних періодів для звірки щойно
згенерованого ОБЛІК.xlsx (content/addon_mismatch_checker.py). Докладна
інструкція по кожному статусу - аркуш "map" у resources/ОБЛІК.xlsx (чернетка
користувача, уточнена разом з ним 2026-10-18; уточнення - в docstring-ах
нижче).

Статуси, яких book61 ще не охоплює (у "map": "не має поки що файлу та даних"
чи "не підається автоматизації") - Задув, Аршт, БЗВП, Адап, НОВ, ЗВФН, ЛМР,
Інт, КФАХ, - тут навмисно не читаються.

На кожному аркуші нижче (Алф2/ВП/ВД/ШП/ВЛК/ППД/СЗЧ/Втрати/Прибулі у ВД)
рядок 1 - заголовки колонок, рядок 2 - службова нумерація колонок самої
книги, а дані починаються з рядка 3 (_DATA_START_ROW)."""

import re
from datetime import date, datetime, timedelta

from constants import (
    ADDON_IMPLAUSIBLE_FUTURE_DAYS,
    ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS,
    OWN_MILITARY_UNIT_CODES,
    TRANSFER_TO_BRIGADE_STATUS,
    TRANSFER_TO_OTHER_UNIT_STATUS,
)
from content.excel_compat import load_workbook
from content.oblik_timesheet import normalize_name

_DATA_START_ROW = 3

# book61 пише "01.01.1990" як заглушку "дату ще не встановлено" ("map", рядок
# 11: "але не 1990") - лише на аркуші "Прибулі у ВД" (read_rrez_tvo_periods).
_PLACEHOLDER_DATE = date(1990, 1, 1)

# Очікуваний статус "будь-який": збігом вважається що завгодно в Табелі
# (content/addon_mismatch_checker.py).
ANY_STATUS = "*"


def load_addon_workbook(path):
    """Будь-яка книга resources/addons (book61 чи ЦВС, content/cvs_reader.py).
    data_only=True - значення формул, а не самі формули, як і для інших
    сторонніх книг у проєкті."""
    return load_workbook(path, data_only=True)


# Аркуші, які відкривають читачі нижче: книга без будь-якого з них - не book61
# (за цим generators/check_addon_mismatches.py визначає тип книги).
BOOK61_SHEETS = ("ВП", "ВД", "ШП", "ВЛК", "ППД", "Алф2", "СЗЧ", "Втрати", "Прибулі у ВД")


def is_book61_workbook(wb):
    return all(name in wb.sheetnames for name in BOOK61_SHEETS)


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


_UNIT_CODE_RE = re.compile(r"[АA]\s*(\d{4})")


def _is_own_unit(place):
    """Чи є в тексті (лікувальний заклад, місце навчання) код власної в/ч
    (OWN_MILITARY_UNIT_CODES з data.json). Спільне для всіх книг addons."""
    return bool(set(_UNIT_CODE_RE.findall(str(place or ""))) & OWN_MILITARY_UNIT_CODES)


def is_implausible_future_date(value, today=None):
    """Дата пізніше за сьогодні більш ніж на ADDON_IMPLAUSIBLE_FUTURE_DAYS -
    найімовірніше, помилка в році."""
    return value is not None and value > (today or date.today()) + timedelta(days=ADDON_IMPLAUSIBLE_FUTURE_DAYS)


def implausible_date_records(period_records, today=None):
    """Записи будь-якої книги addons з такою датою початку чи кінця."""
    return [
        record for record in period_records
        if is_implausible_future_date(record["start_date"], today)
        or is_implausible_future_date(record.get("end_date"), today)
    ]


# Колонка кінця запису (ВД/ВП "Дата прибуття"/"Дата 08 ранку", ШП "Кінець
# (виписка)", ППД "Дата вибуття з ППД") у book61 часто містить не дату, а
# вільний текст з датою: "у ППД з 18.07.2026", "у СЗЧ з 09.03.2026",
# "надходження 27.06.2026, виписка 15.07.2026". _as_date на такому тексті дає
# None, і період лишався б відкритим аж до останньої колонки Табеля. Реальний
# випадок: ВД "12.07.2026-18.07.2026" (саме такий текст у "Дата прибуття")
# хибно перевірявся до жовтня, хоча поїздка скінчилася в липні.
#
# Береться остання дата в тексті: "виписка" (ШП) завжди після
# "надходження", а в "у ВЛК з ..., ... прибув" друга дата - фактичне
# повернення. Текст без дати ("продовжено", "до окремого розпорядження",
# "ТВО", "-") лишає період відкритим - кінця в книзі справді немає.
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
# початку наступного статусу, тож сам період закінчився напередодні.
# Перевірено на реальних даних: цього дня Табель уже показує новий статус.
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
                          end_header=None, fallback_end_header=None, exclusive_end=False, extra_fn=None):
    """Рядки аркуша (з рядка 3) -> записи {"pib_raw", "status", "start_date",
    "end_date", "source_sheet", "arrival_date"}. Спільна основа для всіх
    аркушів, де період - це просто дві дати (ВП/ВПС/ВПСЗ, ВД/Навч,
    ШП/БПШП, ВЛК, ППД). end_date - дата з клітинки або з вільного тексту
    (_end_date_from_cell); None - період відкритий (ще триває чи кінця в
    книзі немає), і звірка йде до останньої колонки-дати Табеля.

    status_fn(row) повертає статус ОБЛІК для рядка (row - {назва колонки:
    значення}) або None, щоб рядок пропустити. Так одна функція обслуговує і
    аркуші з одним статусом (ВЛК, ППД), і ті, де статус залежить від колонки
    (ВП - "Тип відпустки", ВД - "Виплата", ШП - "БЧС").

    exclusive_end=True: дата кінця в book61 - це перший день після періоду
    (людина вже повернулась чи має інший статус), тож останній день періоду
    - на день раніше. Для інших аркушів так само, але лише коли кінець
    записано текстом "у <статус> з <дата>" (_is_transition_text).

    "arrival_date" - день прибуття (кінець періоду, записаний датою без
    зсуву); None, якщо кінець зсунуто чи період відкритий
    (ADDON_ARRIVAL_DAY_STATUSES у content/addon_mismatch_checker.py)."""
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
        shifted = False
        if end is not None and (exclusive_end or _is_transition_text(end_raw)):
            end -= timedelta(days=1)
            shifted = True
        record = {
            "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
            "source_sheet": sheet_name, "arrival_date": None if shifted else end,
        }
        if extra_fn is not None:
            record.update(extra_fn(row))  # додаткові ознаки запису з інших колонок рядка
        records.append(record)
    return records


# "Тип відпустки" на аркуші "ВП" у реальних даних: "Перша/Друга частина
# щорічної відпустки", "По догляду за дитиною до досягнення нею трирічного
# віку" і "Соціальна відпустка" (останніх двох у "map" немає). Усі вони -
# звичайна "ВП" (рішення користувача): у мапі нижче їх немає, тож вони
# просто лишаються на типовому "ВП".
_VP_TYPE_TO_STATUS = {
    "За станом здоров'я": "ВПСЗ",
    "За сімейними обставинами": "ВПС",
}


def read_vp_family_periods(wb):
    """ВП/ВПС/ВПСЗ ("map", рядки 7-9) - аркуш "ВП": період від "Дата вибуття"
    до "Дата прибуття" (а якщо її немає - до "Дата 08 ранку", як у "map")
    мінус один день (exclusive_end). "Дата 08 ранку" - ранок, коли людина вже
    на місці, а "Дата прибуття" на реальних даних здебільшого та сама (чи
    дата переходу в інший статус, "у лік з ..."). Табель цього дня вже не
    "ВП", тож без зсуву останній день кожної відпустки був би розбіжністю.
    Статус - за "Тип відпустки" (_VP_TYPE_TO_STATUS); порожнє чи "?" у
    "Куди вибув" дає ознаку destination_missing (_vp_destination_flags)."""
    return _read_period_records(
        wb, "ВП", pib_header="П.І.Б.", start_header="Дата вибуття",
        end_header="Дата прибуття", fallback_end_header="Дата 08 ранку", exclusive_end=True,
        status_fn=lambda row: _VP_TYPE_TO_STATUS.get(row.get("Тип відпустки"), "ВП"),
        extra_fn=_vp_destination_flags,
    )


# Порожнє чи лише "?" у "Куди вибув" (аркуш "ВП") - розбіжність (рішення
# користувача): звірка з addons фарбує дні такої відпустки в місяці Табеля
# жовтим. Аркуш без цієї колонки (старіша книга) не перевіряється.
_VP_DESTINATION_HEADER = "Куди вибув"
_UNKNOWN_DESTINATION_RE = re.compile(r"[\s?]*")


def _vp_destination_flags(row):
    if _VP_DESTINATION_HEADER not in row:
        return {}
    value = row[_VP_DESTINATION_HEADER]
    return {"destination_missing": value is None or bool(_UNKNOWN_DESTINATION_RE.fullmatch(str(value)))}


_NOT_DEPARTED_RE = re.compile(r"не\s+вибу(в|ла|ли)\b")


def read_vd_periods(wb):
    """ВД/Навч ("map", рядки 12 і 15) - аркуш "ВД": період як у ВП (вибуття -
    прибуття, інакше "08 ранку"), але статус залежить від колонки "Виплата".
    "10" чи порожньо - "ВД" підтверджено; текст "навчання" - ОБЛІК має казати
    "Навч", а не "ВД" (та сама логіка, що й _is_training_destination у
    content/daily_report_reader.py, лише джерело тут - book61). "не
    вибув"/"не вибула" - людина фактично не вибула, тож статус бере Табель
    (ANY_STATUS), і розбіжності на ці дати не буває."""
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


# "БПШП" і "ВПБП" ("map", рядки 22 і 10) обидва беруться з аркуша "ШП" з
# фільтром БЧС="ЛП" (у рядку 10 "ВП АРКУШ" - описка в "map", як підтвердив
# користувач). book61 цих статусів не розрізняє, тож тут виходить лише
# "БПШП", а "ВПБП" вважається рівнозначним під час звірки
# (content/addon_mismatch_checker.py), а не окремим записом на ті самі дні.
_SHP_BCHS_TO_STATUS = {"лік": "ШП", "ЛП": "БПШП"}
_SHP_FACILITY_HEADER = "Назва госпіталя чи в/ч"
_OWN_UNIT_TREATMENT_STATUS = "ЛМР"


def _shp_status(row):
    """Лікування у власній в/ч (код у "Назва госпіталя чи в/ч", у реальних
    даних - "Медична рота військової частини ...") - "ЛМР", як і в книзі ЦВС:
    без цього звірка книг між собою давала хибні "ШП / ЛМР". Інакше - за
    колонкою "БЧС" (_SHP_BCHS_TO_STATUS)."""
    status = _SHP_BCHS_TO_STATUS.get(row.get("БЧС"))
    if status is not None and _is_own_unit(row.get(_SHP_FACILITY_HEADER)):
        return _OWN_UNIT_TREATMENT_STATUS
    return status


def read_shp_family_periods(wb):
    """ШП/БПШП/ЛМР ("map", рядки 20 і 22) - аркуш "ШП": період від "Дата
    вибуття" до "Кінець (виписка)" (відкритий, якщо кінця немає). Статус -
    _shp_status; рядки з будь-яким іншим БЧС до цих статусів не належать і
    пропускаються."""
    return _read_period_records(
        wb, "ШП", pib_header="П.І.Б.", start_header="Дата вибуття", end_header="Кінець (виписка)",
        status_fn=_shp_status,
    )


def read_vlk_periods(wb):
    """ВЛК ("map", рядок 3) - аркуш "ВЛК": період від "Дата вибуття" до
    "Кінець (ВЛК)" (відкритий, якщо кінця немає)."""
    return _read_period_records(
        wb, "ВЛК", pib_header="П.І.Б.", start_header="Дата вибуття", end_header="Кінець (ВЛК)",
        status_fn=lambda row: "ВЛК",
    )


def read_ppd_periods(wb):
    """ППД ("map", рядок 14) - аркуш "ППД": від "Дата прибуття на ППД" до дня
    перед "Дата вибуття з ППД" (відкритий, якщо дати вибуття немає - людина
    ще на ППД). День вибуття - вже перший день наступного статусу, як і
    "Дата вибуття" на аркушах ВП/ВД/ШП/ВЛК; на реальних даних Табель цього
    дня вже показує новий статус."""
    return _read_period_records(
        wb, "ППД", pib_header="П.І.Б.", start_header="Дата прибуття на ППД", end_header="Дата вибуття з ППД",
        status_fn=lambda row: "ППД", exclusive_end=True,
    )


# Дата у вільному тексті колонки "Дата, наказ розпорядження, виключення"
# аркуша "Алф2", напр. "Виключений із списків 31.10.2023".
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


# Колонка "СТАТУС" аркуша "Алф2" -> статус ОБЛІК ("map", рядки 6, 24, 29) -
# лише значення з однозначним відповідником у "map". "переведений в середині
# бригади" - "ПРВДБ", "переведений в іншу в/ч" - "ПРВДВЧ" (замість
# колишнього "ПРВД"). "самостріл" у даних трапляється, але "map" не каже,
# якому статусу він відповідає, тож користувач попросив не вгадувати:
# read_unmapped_alf2_people показує таких людей окремим попередженням у
# звіті.
_ALF2_STATUS_TO_OBLIK = {
    "звільнений": "Звіль",
    "переведений в середині бригади": TRANSFER_TO_BRIGADE_STATUS,
    "переведений в іншу в/ч": TRANSFER_TO_OTHER_UNIT_STATUS,
}

# Події (не накази): "безвісті" - БЗ, "полон" - полон, "загибель" - БЗ або
# Загиб (ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS: на Алф2 немає лікарського
# свідоцтва, а "Загиб" рахується лише з його дати - правило книги ЦВС).
# Дата - з колонки "з" (початок самої події), а якщо її немає - дата наказу.
# Реальний випадок: 74 з 77 людей з "БЗ" у Табелі не мали жодного запису в
# addons - вони є лише на "Алф2", де "з" заповнена завжди, а наказ -
# здебільшого ні (через це й "полон" з Алф2 раніше не читався).
_ALF2_EVENT_STATUS_TO_OBLIK = {
    "безвісті": "БЗ", "загибель": ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS, "полон": "полон",
}

# Значення "СТАТУС" на Алф2, для яких у "map" немає відповідника
# ("самостріл" - з реальних даних). "ВД"/"ВП"/"СЗЧ"/"ШП" сюди навмисно не
# входять: для них є власні, набагато докладніші аркуші book61, тож Алф2 для
# них не використовується.
_ALF2_UNMAPPED_STATUSES = ("самостріл",)


def read_alf2_order_periods(wb):
    """Звіль/ПРВДБ/ПРВДВЧ ("map", рядки 6 і 29) - аркуш "Алф2": статус ОБЛІК
    визначає колонка "СТАТУС" (_ALF2_STATUS_TO_OBLIK), дата - з вільного
    тексту "Дата, наказ розпорядження, виключення" (_order_date_from_text).
    Період завжди відкритий: це адміністративний статус "з моменту цієї дати
    включно" ("map"), без кінця в самій книзі. БЗ/Загиб/полон
    (_ALF2_EVENT_STATUS_TO_OBLIK) - з дати в колонці "з" (немає - з дати
    наказу), кінець - "по", якщо заповнена."""
    ws = wb["Алф2"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б"]
    status_col = headers["СТАТУС"]
    order_col = headers["Дата, наказ розпорядження, виключення"]
    from_col, to_col = headers.get("з"), headers.get("по")

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        status_value = ws.cell(row=row_idx, column=status_col).value
        order_date = _order_date_from_text(ws.cell(row=row_idx, column=order_col).value)
        if status_value in _ALF2_EVENT_STATUS_TO_OBLIK:
            oblik_status = _ALF2_EVENT_STATUS_TO_OBLIK[status_value]
            start = (_as_date(ws.cell(row=row_idx, column=from_col).value) if from_col else None) or order_date
            end = _as_date(ws.cell(row=row_idx, column=to_col).value) if to_col else None
        else:
            oblik_status = _ALF2_STATUS_TO_OBLIK.get(status_value)
            start, end = order_date, None
        if oblik_status is None or start is None:
            continue
        records.append({
            "pib_raw": pib_raw, "status": oblik_status, "start_date": start, "end_date": end,
            "source_sheet": "Алф2",
        })
    return records


# "Дата повернення з СЗЧ" = "у вдр з <дата>" (у відрядження; у реальних
# даних - "Повернувся у в/ч ..." іншої частини) - не кінець СЗЧ для нашої
# частини: СЗЧ триває, доки немає Розп (дата наказу на Алф2) чи виключення зі
# списків (рішення користувача). "у загибель"/"у смерть з" - кінець, як і
# раніше.
_SZCH_CONTINUES_RE = re.compile(r"^\s*у\s+вдр\b", re.IGNORECASE)

# Повернення з СЗЧ у власну в/ч (код з OWN_MILITARY_UNIT_CODES, data.json):
# "Примітка" чи "Куди повернувся з СЗЧ" з кодом власної частини і словом
# повернення - з дати повернення статус "Повер" (відкритий). Сам код без
# слова повернення - ще не повернення (реальний випадок: "перевели з ...
# в/ч <власна>, відразу втік").
RETURNED_FROM_SZCH_STATUS = "Повер"
_SZCH_RETURN_NOTE_HEADERS = ("Куди повернувся з СЗЧ", "Примітка")
_RETURN_WORD_RE = re.compile(r"поверн|прибув", re.IGNORECASE)


def _returned_to_own_unit(note):
    return isinstance(note, str) and bool(_RETURN_WORD_RE.search(note)) and _is_own_unit(note)


def read_szch_to_rozp_periods(wb):
    """СЗЧ -> Розп ("map", рядки 2 і 5) - два послідовні періоди однієї людини:
    1. Аркуш "СЗЧ": від "Дата події" - "СЗЧ" до дня перед "Дата повернення з
       СЗЧ" (у день повернення людина вже на місці, як і "Дата прибуття" у
       ВП); без дати повернення - відкритий. Повернення у власну в/ч
       (_returned_to_own_unit) - далі "Повер" з дати повернення. Реальний
       випадок: СЗЧ 28.02-01.03.2026 без урахування дати повернення тягнувся
       аж до жовтня.
    2. Якщо в тієї самої людини є рядок на "Алф2" із заповненою "Дата, наказ
       розпорядження, виключення" (важить сам факт дати, хоч би що казав
       "СТАТУС"), з цієї дати очікується "Розп".
    Під час звірки за датами записи накладаються: період 2, якщо є,
    перекриває "хвіст" відкритого періоду 1 з дати наказу.

    У запису "Розп" окремий source_sheet ("Алф2 (Розп)", а не "Алф2"), інакше
    під час групування він злився б зі "Звіль"/"ПРВДБ"/"полон" тієї ж людини
    в один рядок звіту."""
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
        return_value = szch_ws.cell(row=row_idx, column=return_col).value
        notes = [szch_ws.cell(row=row_idx, column=szch_headers[name]).value
                 for name in _SZCH_RETURN_NOTE_HEADERS if name in szch_headers]
        own_unit_notes = [note for note in notes if _returned_to_own_unit(note)]
        if own_unit_notes:
            returned = _end_date_from_cell(return_value) or next(
                (date_value for date_value in map(_end_date_from_cell, own_unit_notes) if date_value), None,
            )
        else:
            continues = isinstance(return_value, str) and _SZCH_CONTINUES_RE.match(return_value)
            returned = None if continues else _end_date_from_cell(return_value)
        records.append({
            "pib_raw": pib_raw, "status": "СЗЧ", "start_date": start,
            "end_date": returned - timedelta(days=1) if returned else None,
            "source_sheet": "СЗЧ",
        })
        if own_unit_notes and returned:
            records.append({
                "pib_raw": pib_raw, "status": RETURNED_FROM_SZCH_STATUS, "start_date": returned,
                "end_date": None, "source_sheet": "СЗЧ",
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
    """Люди з Алф2, чий "СТАТУС" - одне з _ALF2_UNMAPPED_STATUSES ("самостріл"),
    для яких "map" не називає статусу ОБЛІК. Не вгадуємо (рішення
    користувача): повертає [{"pib_raw", "alf2_status"}] для окремого
    попередження у звіті, а не для звірки за датами."""
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


# "БЧС" аркуша "Втрати" -> статус ОБЛІК ("map", рядки 23, 26, 27). "смерть"
# випливає з логіки "map", хоча в book61 таких рядків поки немає (станом на
# 2026-10-18 лише "загибель" і "безвісті") - спрацює, щойно з'являться.
_VTRATY_BCHS_TO_OBLIK = {
    "загибель": "Загиб",
    "смерть": "Смерть",
    "безвісті": "БЗ",
}


def read_vtraty_periods(wb):
    """Загиб/Смерть/БЗ ("map", рядки 26, 27, 23) - аркуш "Втрати": з "Дата
    події", період завжди відкритий (кінцевий статус, у book61 не
    "завершується")."""
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


def _arrival_date(value):
    """Дата з аркуша "Прибулі у ВД" або None, якщо її ще не встановлено:
    порожня клітинка, 0 (Excel показує порожню дату, openpyxl читає час
    00:00) чи заглушка _PLACEHOLDER_DATE."""
    day = _as_date(value)
    return None if day == _PLACEHOLDER_DATE else day


def read_rrez_tvo_periods(wb):
    """Аркуш "Прибулі у ВД" - прибуття людини в три послідовні кроки (рішення
    користувача):
    - від "Дата допущення до ТВО" до дня перед "Дата зарахування в списки" -
      "РРез" (ТВО, не оплачується);
    - від "Дата зарахування в списки" до дня перед "Дата прийому посади" -
      лише "ППД" (день зарахування - завжди "ППД");
    - "Дата прийому посади" - будь-який статус (ANY_STATUS).
    Без "Дата зарахування в списки" - "РРез" аж до дня перед прийомом посади.
    Посаду ще не прийнято (дата порожня, 0 чи заглушка _PLACEHOLDER_DATE) -
    останній крок відкритий і триває до наступної події людини в addons
    (рішення користувача). Реальний випадок: ТВО 07.10, зарахування 10.10,
    посади ще немає - "РРез" 07.10-09.10, "ППД" з 10.10; такий рядок колись
    пропускався, і "РРез" у Табелі 10.10 ніхто не звіряв.

    На ці дні аркуш переважає будь-який паралельний період book61 (реальний
    випадок: людина водночас і на аркуші "ППД" з дня допущення до ТВО) -
    _PRIORITY_SOURCE_SHEETS у content/addon_mismatch_checker.py. Колись цей
    аркуш вважався перевіркою "людини не має бути в Табелі", але це звичайні
    щойно прибулі військовослужбовці."""
    ws = wb["Прибулі у ВД"]
    headers = _header_map(ws)
    pib_col = headers["П.І.Б."]

    records = []
    for row_idx in range(_DATA_START_ROW, ws.max_row + 1):
        pib_raw = ws.cell(row=row_idx, column=pib_col).value
        if not pib_raw:
            continue
        row = {name: ws.cell(row=row_idx, column=c).value for name, c in headers.items()}
        tvo = _arrival_date(row.get("Дата допущення до ТВО"))
        if tvo is None:
            continue
        enrollment = _arrival_date(row.get("Дата зарахування в списки"))
        position = _arrival_date(row.get("Дата прийому посади"))

        # Кінець None - відкритий крок (посаду ще не прийнято).
        next_step = enrollment or position
        steps = [("РРез", tvo, next_step - timedelta(days=1) if next_step else None)]
        if enrollment is not None:
            steps.append(("ППД", enrollment, max(enrollment, position - timedelta(days=1)) if position else None))
        if position is not None:
            steps.append((ANY_STATUS, position, position))
        for status, start, end in steps:
            if end is None or start <= end:
                records.append({
                    "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
                    "source_sheet": "Прибулі у ВД",
                })
    return records


def read_all_period_records(wb, source_file):
    """Усі записи "період -> очікуваний статус" разом (усі читачі вище, крім
    read_unmapped_alf2_people - у того інша форма, і він звіряється окремо).
    Записи можуть перекриватися в часі (аркуші book61 ведуться окремо), тож
    одну послідовну шкалу на людину з них будує _paint_timeline у
    content/addon_mismatch_checker.py.

    source_file ("book61" - ім'я файлу без розширення) ставиться на кожен
    запис: у resources/addons можуть бути й інші файли, і звіт показує, з
    якого саме файлу й аркуша кожен запис ("book61: ППД -> <інший файл>: ...")."""
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
