"""Журнал обліку військовослужбовців на лікуванні поза межами військової
частини з resources/addons - ще одне джерело очікуваних періодів для звірки
Табеля (content/addon_mismatch_checker.py). Записи такого самого вигляду, як у
content/addon_reader.py (book61) і content/cvs_reader.py (ЦВС).

Аркуші - за роками, кожен рядок - одне лікування: "Дата госпіталізації" ..
"Дата виписки", порожня "Дата виписки" - період відкритий. Статусу в журналі
немає, тож очікується будь-який з ADDON_TREATMENT_JOURNAL_STATUS. Книга
впізнається за назвами колонок, а не за іменем файлу (у ньому позначення
підрозділу)."""

from constants import ADDON_TREATMENT_JOURNAL_STATUS
from content.cvs_reader import _date, _find_header, _rows

_PIB = "Прізвище, ім'я та по батькові"
_START = "Дата госпіталізації"
_END = "Дата виписки"
_HEADERS = (_PIB, _START, _END)


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
    end_before_start показує його серед підозрілих дат."""
    records = []
    for table in _journal_tables(wb):
        for pib_raw, row in _rows(table, _HEADERS):
            start = _date(row[_START])
            if start is None:
                continue
            end = _date(row[_END])
            records.append({
                "pib_raw": pib_raw, "status": ADDON_TREATMENT_JOURNAL_STATUS, "start_date": start, "end_date": end,
                "source_sheet": table[0].title.strip(), "arrival_date": end, "source_file": source_file,
                "end_before_start": end is not None and end < start,
            })
    return records
