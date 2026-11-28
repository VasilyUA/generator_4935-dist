"""Журнал обліку військовослужбовців на лікуванні поза межами військової
частини з resources/addons - ще одне джерело очікуваних періодів для звірки
Табеля (content/addon_mismatch_checker.py). Записи такого самого вигляду, як у
content/addon_reader.py (book61) і content/cvs_reader.py (ЦВС).

Аркуші - за роками, кожен рядок - одне лікування: "Дата госпіталізації" ..
"Дата виписки", порожня "Дата виписки" - період відкритий. Статусу в журналі
немає, тож очікується будь-який з ADDON_TREATMENT_JOURNAL_STATUS. Книга
впізнається за назвами колонок, а не за іменем файлу (у ньому позначення
підрозділу)."""

import re
from datetime import date

from constants import ADDON_TREATMENT_JOURNAL_STATUS
from content.cvs_reader import _date, _find_header, _rows

_PIB = "Прізвище, ім'я та по батькові"
_START = "Дата госпіталізації"
_END = "Дата виписки"
_HEADERS = (_PIB, _START, _END)
# Необов'язкова колонка: "19.11.26 госпіталізований у ВМКЦ ПР" - людину
# перевели до іншого закладу, і "Дата виписки" тоді лише день переводу.
_TRANSFER = "До якого закладу було переведено та дата переводу"
# Рік у тексті буває й двома цифрами ("19.11.26").
_TRANSFER_DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4}|\d{2})(?!\d)")


def _transfer_date(value):
    """Дата переводу з колонки _TRANSFER: дата Excel чи перша дата в тексті."""
    if not isinstance(value, str):
        return _date(value)
    match = _TRANSFER_DATE_RE.search(value)
    if match is None:
        return None
    day, month, year = match.groups()
    try:
        return date(int(year) + (2000 if len(year) == 2 else 0), int(month), int(day))
    except ValueError:
        return None


def _journal_tables(wb):
    """[(аркуш, рядок заголовків, колонки)] - аркуші з колонками _HEADERS."""
    tables = []
    for ws in wb.worksheets:
        found = _find_header(ws, _HEADERS)
        if found is not None:
            tables.append((ws, *found))
    return tables


def is_treatment_journal_workbook(wb):
    return bool(_journal_tables(wb))


def read_all_period_records(wb, source_file):
    """Усі записи журналу. source_file - ім'я файлу без розширення (назва
    аркуша mis_match.xlsx). Рядок без дати госпіталізації не звіряється.
    Виписка раніше за госпіталізацію (реальний випадок: рік 2025 замість 2026)
    мовчки не виправляється: такий запис нічого не покриває, а
    end_before_start показує його серед підозрілих дат.

    Переведення до іншого закладу (_TRANSFER) з датою, не ранішою за
    госпіталізацію, і "Дата виписки" не пізніша за цю дату (виписали саме для
    переводу): лікування триває, тож запис відкритий, як і без виписки.
    Реальний випадок: госпіталізація 15.10, "19.11.26 госпіталізований у ВМКЦ
    ПР", виписка 19.11 - звірка писала "немає запису в журналі" з 20.11, хоч
    людина й далі на лікуванні (БПШП). Пізніша виписка - вже з нового закладу,
    вона й закінчує запис."""
    records = []
    for table in _journal_tables(wb):
        for pib_raw, row in _rows(table, (*_HEADERS, _TRANSFER)):
            start = _date(row[_START])
            if start is None:
                continue
            end = _date(row[_END])
            transferred = _transfer_date(row[_TRANSFER])
            if transferred is not None and transferred >= start and (end is None or end <= transferred):
                end = None
            records.append({
                "pib_raw": pib_raw, "status": ADDON_TREATMENT_JOURNAL_STATUS, "start_date": start, "end_date": end,
                "source_sheet": table[0].title.strip(), "arrival_date": end, "source_file": source_file,
                "end_before_start": end is not None and end < start,
            })
    return records
