"""Журнал обліку бойової підготовки з resources/addons - ще одне джерело
очікуваних періодів для звірки Табеля (content/addon_mismatch_checker.py).
Записи такого самого вигляду, як у content/addon_reader.py (book61) і
content/cvs_reader.py (ЦВС).

Аркуш журналу визначає статус ОБЛІК (ADDON_TRAINING_JOURNAL_SHEETS: статус
за замовчуванням і уточнення за колонкою "НЦ"). Рядок - "ПІБ" і період
"З" .. "До": порожнє "До" - період відкритий, а "До" - це день повернення
(ADDON_ARRIVAL_DAY_SHEETS). Книга впізнається за назвами аркушів і колонок,
а не за іменем файлу (у ньому позначення підрозділу)."""

from constants import ADDON_JOURNAL_OWN_UNIT, ADDON_TRAINING_JOURNAL_SHEETS
from content.addon_reader import _is_own_unit
from content.cvs_reader import _date, _find_header, _norm, _rows

_HEADERS = ("ПІБ", "З", "До")
_PLACE = "НЦ"


def _journal_tables(wb):
    """[(аркуш, (статус за замовчуванням, {НЦ: статус}), рядок заголовків,
    колонки)] - аркуші журналу з колонками _HEADERS."""
    rules_by_sheet = {_norm(name): rules for name, rules in ADDON_TRAINING_JOURNAL_SHEETS.items()}
    tables = []
    for ws in wb.worksheets:
        rules = rules_by_sheet.get(_norm(ws.title))
        if rules is None:
            continue
        found = _find_header(ws, _HEADERS)
        if found is not None:
            tables.append((ws, rules, *found))
    return tables


def _status(rules, place):
    """Статус рядка за колонкою "НЦ" (ADDON_TRAINING_JOURNAL_SHEETS)."""
    default, by_place = rules
    for key, status in by_place.items():
        if key == ADDON_JOURNAL_OWN_UNIT:
            if _is_own_unit(place):
                return status
        elif _norm(key) == _norm(place):
            return status
    return default


def is_training_journal_workbook(wb):
    return bool(_journal_tables(wb))


def read_all_period_records(wb, source_file):
    """Усі записи журналу. source_file - ім'я файлу без розширення (префікс
    колонки "Джерело" й назва аркуша mis_match.xlsx)."""
    records = []
    for ws, rules, header_row, columns in _journal_tables(wb):
        for pib_raw, row in _rows((ws, header_row, columns), _HEADERS + (_PLACE,)):
            start = _date(row["З"])
            if start is None:
                continue
            end = _date(row["До"])
            records.append({
                "pib_raw": pib_raw, "status": _status(rules, row[_PLACE]), "start_date": start, "end_date": end,
                "source_sheet": ws.title.strip(), "arrival_date": end, "source_file": source_file,
            })
    return records
