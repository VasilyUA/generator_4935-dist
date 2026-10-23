"""Читання книги ЦВС з resources/addons - ще одне джерело "очікуваних
періодів" для звірки Табеля (content/addon_mismatch_checker.py), записи того
самого вигляду, що й у content/addon_reader.py (book61). Що звіряти - за
прямою вказівкою користувача, аркуш "map" самої книги (колонка B - потрібні
колонки, C - умова):

- ВЛК - "Дата проходження ВЛК" .. дата з тексту "Номер ВЛК/Номер
  протоколу/Свідоцтво про хворобу" ("... від 16.01.2026" - завершення); лише
  одна з дат - один день (на реальних даних колонку "Дата проходження ВЛК"
  майже ніколи не заповнюють);
- ЗВФН - "Дата початку" .. "Дата завершення";
- ШП/ЛМР/БПШП - аркуш лікування: "Дата Надходження в мед заклад" .. "Дата
  виписки" (порожня - період відкритий). Лікувальний заклад власної
  військової частини (OWN_MILITARY_UNIT_CODES) -> "ЛМР"; інакше поранення
  (чи бойове травмування) -> "БПШП", соматика/травмування -> "ШП"; інших
  типів ("після полону", "реабілітація") map не називає - не звіряються;
- БЗ - "Дата зникнення", відкритий; рядки зі статусом СЗЧ не враховуються
  (map: "не враховуючи сзч"). map називає "Дата зарахування в частини", але
  це дата зарахування на службу (роки до зникнення) - з неї "БЗ" перекрили б
  будь-які пізніші події ще до самого зникнення;
- полон - "Дата в полоні" .. день перед "Дата звільнення з полону";
- Загиб - з дати "лікарське свідоцтво про смерть (номер, дата)" (за прямою
  вказівкою користувача - лише після медичного документа), без неї - "Дата
  втрат (смерті) особового складу"; відкритий.

Аркуш шукається за заголовками колонок, а не за назвою (назви аркушів і
файлу можуть містити позначення підрозділів, а рядок заголовків - не завжди
перший); назва аркуша без зайвих пробілів іде в колонку "Джерело" звіту."""

import re
from datetime import datetime, timedelta

from constants import OWN_MILITARY_UNIT_CODES
from content.addon_reader import _as_date

_HEADER_SEARCH_ROWS = 5
_DATE_TEXT_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")
_DATED_FROM_RE = re.compile(r"від\s+(\d{2}\.\d{2}\.\d{4})", re.IGNORECASE)
_UNIT_CODE_RE = re.compile(r"[АA]\s*(\d{4})")

_VLK_HEADERS = ("ПІБ", "Дата проходження ВЛК")
_VLK_DOCUMENT = "Номер ВЛК/Номер протоколу/Свідоцтво про хворобу"
_ZVFN_HEADERS = ("ПІБ", "Дата початку", "Дата завершення")
_MEDICAL_HEADERS = (
    "Прізвище ім'я по батькові", "Дата Надходження в мед заклад", "Дата виписки",
    "соматика / травмування / поранення", "Лікувальний заклад",
)
_MISSING_HEADERS = ("ПІБ", "Дата зникнення")
_MISSING_STATUS = "Статус"
_CAPTIVITY_HEADERS = ("ПІБ", "Дата в полоні")
_CAPTIVITY_END = "Дата звільнення з полону"
_KILLED_HEADERS = ("ПІБ", "Дата втрат (смерті) особового складу")
_KILLED_CERTIFICATE = "лікарське свідоцтво про смерть (номер, дата)"

_ALL_HEADER_SETS = (_VLK_HEADERS, _ZVFN_HEADERS, _MEDICAL_HEADERS, _MISSING_HEADERS, _CAPTIVITY_HEADERS, _KILLED_HEADERS)


def _norm(value):
    if value is None:
        return ""
    return " ".join(str(value).replace("’", "'").replace("ʼ", "'").split()).lower()


def _date(value, pattern=_DATE_TEXT_RE):
    """Дата-клітинка або дата, вбудована в текст (pattern - де саме її шукати)."""
    direct = _as_date(value)
    if direct is not None or not isinstance(value, str):
        return direct
    match = pattern.search(value)
    if match is None:
        return None
    try:
        return datetime.strptime(match.group(match.lastindex or 0), "%d.%m.%Y").date()
    except ValueError:
        return None


def _find_header(ws, headers):
    """(рядок заголовків, {нормалізований заголовок: колонка}) - якщо в
    одному з перших _HEADER_SEARCH_ROWS рядків ws є ВСІ headers (без
    урахування регістру й зайвих пробілів); None - немає."""
    wanted = [_norm(header) for header in headers]
    for row in ws.iter_rows(max_row=_HEADER_SEARCH_ROWS):
        columns = {}
        for cell in row:
            if cell.value is not None:
                columns.setdefault(_norm(cell.value), cell.column)
        if all(header in columns for header in wanted):
            return row[0].row, columns
    return None


def _find_table(wb, headers):
    """(аркуш, рядок заголовків, колонки) - перший аркуш книги, де є ВСІ
    headers (_find_header); None - немає."""
    for ws in wb.worksheets:
        found = _find_header(ws, headers)
        if found is not None:
            return (ws, *found)
    return None


def _rows(table, headers):
    """(ПІБ, {заголовок: значення}) для кожного рядка даних - рядки без
    ПІБ-тексту (номери колонок, підзаголовки) пропускаються; заголовок,
    якого на аркуші немає, дає None."""
    ws, header_row, columns = table
    found = {header: columns.get(_norm(header)) for header in headers}
    pib_col = found[headers[0]]
    max_col = max(col for col in found.values() if col is not None)
    for values in ws.iter_rows(min_row=header_row + 1, max_col=max_col, values_only=True):
        pib_raw = values[pib_col - 1]
        if not isinstance(pib_raw, str) or not any(ch.isalpha() for ch in pib_raw):
            continue
        yield pib_raw, {header: values[col - 1] if col else None for header, col in found.items()}


def _record(ws, pib_raw, status, start, end, arrival_date=None):
    return {
        "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
        "source_sheet": ws.title.strip(), "arrival_date": arrival_date,
    }


def _read(wb, headers, optional, build):
    """Записи з аркуша, знайденого за headers; build(ws, pib_raw, row) ->
    запис або None (рядок пропускається)."""
    table = _find_table(wb, headers)
    if table is None:
        return []
    records = []
    for pib_raw, row in _rows(table, headers + optional):
        record = build(table[0], pib_raw, row)
        if record is not None:
            records.append(record)
    return records


def _vlk_record(ws, pib_raw, row):
    """"Дата проходження ВЛК" - початок, дата документа ("від 20.10.2026") -
    завершення ВЛК (реальний випадок, підтверджений користувачем; на
    реальних даних документ - через 0-6 днів після початку). Лише одна з
    дат (чи документ, старіший за початок) - один день."""
    document = row[_VLK_DOCUMENT]
    passed = _date(row["Дата проходження ВЛК"])
    concluded = _date(document, _DATED_FROM_RE) or _date(document)
    start = passed or concluded
    if start is None:
        return None
    end = concluded if concluded is not None and concluded >= start else start
    return _record(ws, pib_raw, "ВЛК", start, end, arrival_date=end)


def _zvfn_record(ws, pib_raw, row):
    start = _date(row["Дата початку"])
    return _record(ws, pib_raw, "ЗВФН", start, _date(row["Дата завершення"])) if start else None


def _is_own_unit(facility):
    return bool(set(_UNIT_CODE_RE.findall(str(facility or ""))) & OWN_MILITARY_UNIT_CODES)


def _medical_status(facility, kind):
    """map: заклад власної частини - "ЛМР"; інакше поранення чи бойове
    травмування - "БПШП", соматика/травмування - "ШП"; решта - None."""
    if _is_own_unit(facility):
        return "ЛМР"
    kind = _norm(kind)
    if "поранен" in kind or ("травм" in kind and "бойов" in kind):
        return "БПШП"
    if "травм" in kind or "соматик" in kind:
        return "ШП"
    return None


def _medical_record(ws, pib_raw, row):
    status = _medical_status(row["Лікувальний заклад"], row["соматика / травмування / поранення"])
    start = _date(row["Дата Надходження в мед заклад"])
    if status is None or start is None:
        return None
    end = _date(row["Дата виписки"])
    return _record(ws, pib_raw, status, start, end, arrival_date=end)


def _missing_record(ws, pib_raw, row):
    start = _date(row["Дата зникнення"])
    if start is None or "сзч" in _norm(row[_MISSING_STATUS]):
        return None
    return _record(ws, pib_raw, "БЗ", start, None)


def _captivity_record(ws, pib_raw, row):
    start = _date(row["Дата в полоні"])
    if start is None:
        return None
    released = _date(row[_CAPTIVITY_END])
    return _record(ws, pib_raw, "полон", start, released - timedelta(days=1) if released else None)


def _killed_record(ws, pib_raw, row):
    """За прямою вказівкою користувача "Загиб" - лише з лікарського
    свідоцтва про смерть: дата з колонки свідоцтва ("№... від 08.10.2026");
    "Дата втрат (смерті) особового складу" - лише коли дати свідоцтва немає."""
    certificate = row[_KILLED_CERTIFICATE]
    start = (
        _date(certificate, _DATED_FROM_RE) or _date(certificate)
        or _date(row["Дата втрат (смерті) особового складу"])
    )
    return _record(ws, pib_raw, "Загиб", start, None) if start else None


def is_cvs_workbook(wb):
    """Книга ЦВС - є хоча б один аркуш із заголовками зі списку вище."""
    return any(_find_table(wb, headers) is not None for headers in _ALL_HEADER_SETS)


def read_all_period_records(wb, source_file):
    """Усі period-записи книги ЦВС; source_file - ім'я файлу без розширення
    (префікс колонки "Джерело", як і в content/addon_reader.py)."""
    records = (
        _read(wb, _VLK_HEADERS, (_VLK_DOCUMENT,), _vlk_record)
        + _read(wb, _ZVFN_HEADERS, (), _zvfn_record)
        + _read(wb, _MEDICAL_HEADERS, (), _medical_record)
        + _read(wb, _MISSING_HEADERS, (_MISSING_STATUS,), _missing_record)
        + _read(wb, _CAPTIVITY_HEADERS, (_CAPTIVITY_END,), _captivity_record)
        + _read(wb, _KILLED_HEADERS, (_KILLED_CERTIFICATE,), _killed_record)
    )
    for record in records:
        record["source_file"] = source_file
    return records
