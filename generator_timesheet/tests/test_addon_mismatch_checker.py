import os
import unicodedata
from collections import Counter
from datetime import date, datetime

import pytest
from book61_fixtures import write_full_book61
from cvs_fixtures import MEDICAL_SHEET, VLK_SHEET, write_cvs
from openpyxl import Workbook, load_workbook
from treatment_fixtures import write_treatment_journal

from constants import (
    ADDON_BOOK_KINDS,
    ADDON_NON_COMBAT_TREATMENT_STATUS,
    ADDON_TREATMENT_JOURNAL_KIND,
    ADDON_TREATMENT_JOURNAL_STATUS,
)
import content.addon_mismatch_checker as amc
from content.oblik_timesheet import Timesheet, normalize_name
from content.report_pib_column import pib_column_width
from generators.check_addon_mismatches import addon_files, check_addon_mismatches_file, cross_check_addons

_HEADER = ["ПОСАДА", "ЗВАННЯ", "ПІБ", datetime(2026, 7, 1), datetime(2026, 7, 2), datetime(2026, 7, 3)]
_WEEK_HEADER = ["ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 10, day) for day in range(1, 8)]


def _write_roster(path, rows, header=_HEADER):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Табель"
    for col_idx, value in enumerate(header, start=1):
        ws.cell(row=1, column=col_idx, value=value)
    for row_idx, row_values in enumerate(rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(str(path))
    return str(path)


def _period(pib_raw, status, start, end, source_sheet="ВП", source_file="book61"):
    return {
        "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
        "source_sheet": source_sheet, "source_file": source_file,
    }


# --- _statuses_match ---------------------------------------------------------

def test_statuses_match_exact():
    assert amc._statuses_match("ВП", "ВП") is True
    assert amc._statuses_match("ВП", "ВЛК") is False


def test_statuses_match_treats_bpshp_and_vpbp_as_equivalent():
    """Рішення користувача: book61 НЕ розрізняє "БПШП"/"ВПБП"
    на рівні даних (content/addon_reader.py::read_shp_family_periods)."""
    assert amc._statuses_match("БПШП", "ВПБП") is True
    assert amc._statuses_match("ВПБП", "БПШП") is True


def test_statuses_match_treats_ppd_and_nov_as_equivalent():
    """Рішення користувача: "ППД" і "НОВ" ідентичні статуси,
    це не розбіжність (незалежно від регістру написання в Табелі)."""
    assert amc._statuses_match("ППД", "НОВ") is True
    assert amc._statuses_match("НОВ", "ППД") is True
    assert amc._statuses_match("ППД", "Нов") is True
    assert amc._statuses_match("СЗЧ", "НОВ") is False


def test_statuses_match_treats_zagyb_and_smert_as_equivalent():
    """Рішення користувача: "Смерть" і "Загиб" не дають
    розбіжності між собою (реальний випадок: book61 "Загиб", Табель
    "Смерть")."""
    assert amc._statuses_match("Загиб", "Смерть") is True
    assert amc._statuses_match("Смерть", "Загиб") is True
    assert amc._statuses_match("Загиб", "РВЗ") is False


def test_statuses_match_treats_ppd_and_bzvp_as_equivalent():
    """Рішення користувача: "ППД" і "БЗВП" ідентичні статуси,
    це не розбіжність."""
    assert amc._statuses_match("ППД", "БЗВП") is True
    assert amc._statuses_match("БЗВП", "ППД") is True
    assert amc._statuses_match("ППД", "Бзвп") is True


def test_statuses_match_prefers_tabel_10_med_over_book61_vpsz():
    """Рішення користувача: book61 "ВПСЗ", Табель "10_МЕД":
    пріоритетніший 10_МЕД (за нього платиться), це не розбіжність. Пріоритет
    односторонній - це не еквівалентність статусів."""
    assert amc._statuses_match("ВПСЗ", "10_МЕД") is True
    assert amc._statuses_match("ВПСЗ", "10_мед") is True
    assert amc._statuses_match("10_МЕД", "ВПСЗ") is False
    assert amc._statuses_match("ВПСЗ", "РВЗ") is False


@pytest.mark.parametrize("expected", ["БЗВП", "Адап"])
def test_statuses_match_prefers_tabel_rrez_over_an_expected_bzvp_or_adaptation(expected):
    """Рішення користувача: "РРез" пріоритетніший за "БЗВП" і
    "Адап": журнал очікує БЗВП/Адап, Табель - РРез (будь-яке написання) - не
    розбіжність; навпаки (очікується РРез, у Табелі БЗВП/Адап) - розбіжність."""
    assert amc._statuses_match(expected, "РРез") is True
    assert amc._statuses_match(expected, "Ррез") is True
    assert amc._statuses_match("РРез", expected) is False


def test_find_period_mismatches_skips_vpsz_when_tabel_has_10_med(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "10_МЕД", "10_МЕД", "10_МЕД"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ВПСЗ", date(2026, 7, 1), None, source_sheet="ВП")]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_skips_a_person_whose_only_difference_is_ppd_vs_nov(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "НОВ", "НОВ", "ППД"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 7, 1), None, source_sheet="ППД")]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_write_addon_mismatch_report_shows_the_tabel_value_for_an_equivalent_status(tmp_path):
    """Рішення користувача: пріоритет Табеля: для ППД/НОВ
    клітинка - ОДНА (об'єднана), зелена, зі значенням із Табеля."""
    period_mismatches = [{
        "source_file": "book61",
        "sources": ["book61: ППД", "book61: СЗЧ"], "roster_pidrozdil": "Підрозділ1",
        "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
        "dates": {date(2026, 10, 1): ("НОВ", "ППД"), date(2026, 10, 2): ("РВЗ", "СЗЧ")},
    }]
    output_path = str(tmp_path / "mis_match.xlsx")

    amc.write_addon_mismatch_report(period_mismatches, [], output_path)

    ws = load_workbook(output_path)["book61"]
    assert ws.cell(row=2, column=4).value == "НОВ"
    assert "D2:D3" in {str(merged) for merged in ws.merged_cells.ranges}
    assert ws.cell(row=2, column=4).fill.start_color.rgb == amc._MATCH_FILL.start_color.rgb


def test_statuses_match_treats_none_as_distinct_from_any_status():
    assert amc._statuses_match("ВП", None) is False


def test_statuses_match_treats_payment_converted_forms_as_equivalent():
    """Реальні випадки на output/ОБЛІК.xlsx: apply_payment_values (content/
    oblik_timesheet.py) може замінити ПОДІЄВИЙ статус на фіксовану
    категорію виплати (constants.py: PAYMENT_STATUS_CONVERSIONS/
    PAYMENT_STATUS_OVERRIDES) - book61 звіряється зі своїм ВЛАСНИМ,
    незалежним визначенням факту (до будь-якої конвертації), тож це й
    далі ТА САМА ситуація, а не розбіжність ("БПШП"/"100_БПШП",
    "ВД"/10)."""
    assert amc._statuses_match("БПШП", "100_БПШП") is True
    assert amc._statuses_match("ВД", 10) is True
    assert amc._statuses_match("ВПСЗ", "100_ВПБП") is True


def test_statuses_match_does_not_treat_navch_as_equivalent_to_10():
    """Застереження користувача: "10" еквівалентний "ВД", АЛЕ якщо book61
    очікує саме "Навч" (не "ВД") - число 10 в Табелі й далі МАЄ лишатись
    розбіжністю, а не гаситись через ширше правило "10 ~ ВД" вище."""
    assert amc._statuses_match("Навч", 10) is False


# --- find_period_mismatches ---------------------------------------------------

def test_find_period_mismatches_skips_a_matching_period(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", "ВП", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 7, 1), date(2026, 7, 2))]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_reports_a_genuine_mismatch(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 7, 1), date(2026, 7, 2))]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert results[0]["sources"] == ["book61: ВП"]
    assert results[0]["roster_pib_raw"] == "ПЕРШИЙ Перший Перший"
    assert results[0]["dates"] == {
        date(2026, 7, 1): ("РВЗ", "ВП"),
        date(2026, 7, 2): ("РВЗ", "ВП"),
    }


def test_find_period_mismatches_ignores_not_yet_filled_in_cells(tmp_path):
    """Реальний випадок (виявлено на book61.xlsm): дата ПІСЛЯ останнього
    наявного рапорту - apply_events (content/oblik_timesheet.py) навмисно
    лишає такі колонки порожніми - це НЕ розбіжність, ще просто немає
    даних, щоб судити."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", None, None],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 7, 1), date(2026, 7, 3))]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_never_includes_a_person_missing_from_the_roster(tmp_path):
    """Рішення користувача: БЕЗ винятків: кого немає в Табелі,
    той не звіряється з жодним аркушем book61 (реальні випадки з аркушів
    "ВД" і "СЗЧ" хибно потрапляли в mis_match.xlsx)."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ІНШИЙ Інший Інший", "РВЗ", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "Навч", date(2026, 7, 1), date(2026, 7, 2), source_sheet="ВД"),
        _period("ДРУГИЙ Другий Другий", "ППД", date(2026, 7, 1), None, source_sheet="ППД"),
        _period("ТРЕТІЙ Третій Третій", "СЗЧ", date(2026, 7, 1), None, source_sheet="СЗЧ"),
        _period("ЧЕТВЕРТИЙ Четвертий Четвертий", "ПРВДБ", date(2023, 10, 31), None, source_sheet="Алф2"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_builds_one_sequential_timeline_per_person(tmp_path):
    """Рішення користувача (реальний випадок): book61 "ППД" до
    04.10, "ШП" 03.10-06.10, "СЗЧ" з 06.10 - ОДИН рядок звіту, кожен період
    триває до дня перед початком наступного: ППД до 02.10 -> ШП
    03.10-05.10 -> СЗЧ з 06.10."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", "ППД", "ШП", "ШП", "ШП", "СЗЧ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 10, 6), None, source_sheet="СЗЧ"),
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 8, 13), date(2026, 10, 4), source_sheet="ППД"),
        _period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 10, 3), date(2026, 10, 6), source_sheet="ШП"),
    ]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert results[0]["sources"] == ["book61: ППД", "book61: ШП", "book61: СЗЧ"]
    assert {d.day: expected for d, (_actual, expected) in results[0]["dates"].items()} == {
        1: "ППД", 2: "ППД", 3: "ШП", 4: "ШП", 5: "ШП", 6: "СЗЧ", 7: "СЗЧ",
    }


def test_find_period_mismatches_does_not_revive_an_old_open_status_after_newer_periods(tmp_path):
    """Реальний випадок, підтверджений користувачем: "Звіль" з Алф2 за
    2025 рік (відкритий) не може діяти в 2026, якщо після нього вже є
    пізніші періоди (ВЛК, що закінчився 03.10) - новіша подія закриває
    давню, і після кінця ВЛК "Звіль" НЕ повертається."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВЛК", "ВЛК", "ВЛК", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "Звіль", date(2025, 11, 15), None, source_sheet="Алф2"),
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 9, 28), date(2026, 10, 3), source_sheet="ВЛК"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_does_not_revive_an_open_szch_after_a_later_period(tmp_path):
    """Реальний випадок, підтверджений користувачем: СЗЧ з лютого (без дати
    повернення в book61) "повертався" після ВП, що закінчився 03.10, - і
    04.10 хибно очікувалось СЗЧ. Пізніша подія закриває СЗЧ назавжди."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 2, 28), None, source_sheet="СЗЧ"),
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 9, 17), date(2026, 10, 3), source_sheet="ВП"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_ends_a_final_status_at_the_next_event_even_if_it_is_later_resumable(tmp_path):
    """"Розп"/"Звіль"/... - остаточні статуси: пізніший ВП закриває їх
    назавжди, навіть коли сам ВП короткий і після нього нічого немає."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "Розп", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "Розп", date(2026, 9, 1), None, source_sheet="Алф2 (Розп)"),
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 10, 2), date(2026, 10, 3), source_sheet="ВП"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_ends_an_open_period_at_the_next_event_except_ppd(tmp_path):
    """Реальний випадок (книга ЦВС): лікування без "Дата виписки" не
    "повертається" після пізнішої ВЛК до кінця місяця; відкритий ППД -
    повертається (людина знову на ППД після лікування)."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ШП", "ВЛК", "100", "100", "100", "100", "100"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ППД", "ШП", "ППД", "ППД", "ППД", "ППД", "ППД"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 9, 20), None, source_sheet="СОМАТИКА-ПОРАНЕНІ"),
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 2), date(2026, 10, 2), source_sheet="ВЛК"),
        _period("ДРУГИЙ Другий Другий", "ППД", date(2026, 9, 1), None, source_sheet="ППД"),
        _period("ДРУГИЙ Другий Другий", "ШП", date(2026, 10, 2), date(2026, 10, 2), source_sheet="ШП"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_ends_an_open_period_where_another_book_ends_the_same_status(tmp_path):
    """Реальний випадок, підтверджений користувачем: book61 "ВЛК" з 14.10
    без "Кінець (ВЛК)", ЦВС - ВЛК 14.10-20.10 (той самий день початку) -
    з 21.10 ВЛК не очікується. Так само, коли запис іншої книги почався
    раніше, але закінчився вже після початку відкритого періоду."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "ВЛК", "ВЛК", "ВЛК", "100", "100", "100"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ШП", "ШП", "ШП", "ШП", "100", "100", "100"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 2), None, source_sheet="ВЛК"),
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 2), date(2026, 10, 4), source_sheet="ВЛК", source_file="cvs"),
        _period("ДРУГИЙ Другий Другий", "ШП", date(2026, 10, 2), None, source_sheet="ШП"),
        _period("ДРУГИЙ Другий Другий", "ШП", date(2026, 10, 1), date(2026, 10, 4), source_sheet="СОМАТИКА-ПОРАНЕНІ",
                source_file="cvs"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_resumes_a_regular_period_after_an_interruption(tmp_path):
    """Звичайний (не остаточний) період, перерваний коротшим пізнішим, -
    продовжується після нього до власного кінця за book61."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", "ВП", "ВЛК", "ВЛК", "ВП", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 10, 1), date(2026, 10, 6), source_sheet="ВП"),
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 3), date(2026, 10, 4), source_sheet="ВЛК"),
    ]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert {d.day: expected for d, (_actual, expected) in results[0]["dates"].items()} == {
        1: "ВП", 2: "ВП", 3: "ВЛК", 4: "ВЛК", 5: "ВП", 6: "ВП",
    }


def test_find_period_mismatches_accepts_the_whole_arrival_sequence_without_errors(tmp_path):
    """Рішення користувача ("Прибулі у ВД"): до зарахування в
    списки - "РРез", від зарахування - лише "ППД" (у Табелі - 10), у день
    прийому посади - будь-який статус (навіть коли паралельний аркуш "ППД"
    каже інакше), далі - знову аркуш "ППД". Жодної розбіжності."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "Ррез", "Ррез", "Ррез", 10, 10, 30, 10],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 9, 22), None, source_sheet="ППД"),
        _period("ПЕРШИЙ Перший Перший", "РРез", date(2026, 9, 22), date(2026, 10, 3), source_sheet="Прибулі у ВД"),
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 4), date(2026, 10, 5), source_sheet="Прибулі у ВД"),
        _period("ПЕРШИЙ Перший Перший", amc.ANY_STATUS, date(2026, 10, 6), date(2026, 10, 6),
                source_sheet="Прибулі у ВД"),
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_prefers_rvz_on_the_arrival_day(tmp_path):
    """Рішення користувача: у день прибуття з ШП/ВЛК/ВД
    пріоритетніший "РВЗ" (у Табелі для виплат - 30, 100): це НЕ
    розбіжність; сам статус аркуша в цей день - теж ні."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ШП", "ШП", 30, "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ВЛК", "ВЛК", "Рвз", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "ВД", 100, "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада4", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "ШП", "ШП", "ШП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        {**_period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 9, 28), date(2026, 10, 3), source_sheet="ШП"),
         "arrival_date": date(2026, 10, 3)},
        {**_period("ДРУГИЙ Другий Другий", "ВЛК", date(2026, 9, 28), date(2026, 10, 3), source_sheet="ВЛК"),
         "arrival_date": date(2026, 10, 3)},
        {**_period("ТРЕТІЙ Третій Третій", "ВД", date(2026, 9, 28), date(2026, 10, 2), source_sheet="ВД"),
         "arrival_date": date(2026, 10, 2)},
        {**_period("ЧЕТВЕРТИЙ Четвертий Четвертий", "ШП", date(2026, 9, 28), date(2026, 10, 3), source_sheet="ШП"),
         "arrival_date": date(2026, 10, 3)},
    ]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_still_flags_another_status_on_the_arrival_day(tmp_path):
    """Не "РВЗ"/30/100 і не сам статус аркуша (напр. СЗЧ) у день прибуття -
    і далі розбіжність; так само "РВЗ" ДО дня прибуття."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ШП", 30, "СЗЧ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [{**_period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 9, 28), date(2026, 10, 3), source_sheet="ШП"),
                "arrival_date": date(2026, 10, 3)}]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    flagged = {d.day for d, (actual, expected) in results[0]["dates"].items() if not amc._statuses_match(expected, actual)}
    assert flagged == {2, 3}


def test_find_period_mismatches_accepts_any_tabel_status_when_vd_says_not_departed(tmp_path):
    """book61 "ВД", "Виплата" - "не вибув": статус бере Табель, розбіжності
    немає, хоч би що там стояло."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", 30, 100, "ППД", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", amc.ANY_STATUS, date(2026, 10, 1), date(2026, 10, 4), source_sheet="ВД")]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_statuses_match_accepts_anything_for_the_any_status():
    assert amc._statuses_match(amc.ANY_STATUS, "30") is True
    assert amc._statuses_match(amc.ANY_STATUS, 100) is True


def test_open_arrival_period_without_a_position_is_checked_and_ends_at_the_next_event(tmp_path):
    """Рішення користувача (реальний випадок): ТВО, зарахування, посаду ще
    не прийнято - "РРез" до зарахування, далі відкритий "ППД"; "РРез" у
    Табелі в день зарахування - розбіжність. Відкритий "ППД" накладається
    поверх решти, тож закінчується перед пізнішою подією (відпустка)."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "РРез", "РРез", "РРез", "РРез", "РВЗ", "РВЗ", None],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ППД", "ППД", "ППД", "ВПС", "ППД", "ППД", "ППД"],
    ], header=_WEEK_HEADER)
    records = [
        _period("ПЕРШИЙ Перший Перший", "РРез", date(2026, 10, 1), date(2026, 10, 3), source_sheet="Прибулі у ВД"),
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 4), None, source_sheet="Прибулі у ВД"),
        _period("ДРУГИЙ Другий Другий", "ППД", date(2026, 10, 1), None, source_sheet="Прибулі у ВД"),
        _period("ДРУГИЙ Другий Другий", "ППД", date(2026, 10, 1), date(2026, 10, 3), source_sheet="ППД"),
        _period("ДРУГИЙ Другий Другий", "ВПС", date(2026, 10, 4), date(2026, 10, 4), source_sheet="ВП"),
        _period("ДРУГИЙ Другий Другий", "ППД", date(2026, 10, 5), None, source_sheet="ППД"),
    ]

    results = amc.find_period_mismatches(Timesheet(path, "Табель"), records)

    assert [r["roster_pib_raw"] for r in results] == ["ПЕРШИЙ Перший Перший"]
    flagged = {d.day: (actual, expected) for d, (actual, expected) in results[0]["dates"].items()
               if not amc._statuses_match(expected, actual)}
    assert flagged == {4: ("РРез", "ППД"), 5: ("РВЗ", "ППД"), 6: ("РВЗ", "ППД")}


def test_find_period_mismatches_gives_rrez_priority_and_resumes_the_parallel_period(tmp_path):
    """Рішення користувача: "РРез" (ТВО, не оплачується)
    переважає паралельний "ППД" на весь свій проміжок, навіть якщо ППД
    почався пізніше; ППД продовжується лише після кінця ТВО."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "Ррез", "Ррез", "Ррез", "ППД", "ППД", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "РРез", date(2026, 10, 2), date(2026, 10, 4), source_sheet="Прибулі у ВД"),
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 3), None, source_sheet="ППД"),
    ]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert results[0]["sources"] == ["book61: Прибулі у ВД", "book61: ППД"]
    assert {d.day: expected for d, (_actual, expected) in results[0]["dates"].items()} == {
        2: "РРез", 3: "РРез", 4: "РРез", 5: "ППД", 6: "ППД", 7: "ППД",
    }


def test_find_period_mismatches_never_duplicates_a_person_within_a_file(tmp_path):
    """Рішення користувача: ПІБ у mis_match.xlsx не
    дублюються, а звіт розділено за addon-файлами: записи однієї людини з
    КІЛЬКОХ аркушів і з різним написанням ПІБ (подвійний пробіл, регістр) -
    ОДИН запис на кожен файл, кожен зі своїм ланцюжком аркушів і лише своїми
    датами."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", "ППД", "ШП", "ВП", "СЗЧ", "СЗЧ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 1), date(2026, 10, 2), source_sheet="ППД"),
        _period("ПЕРШИЙ  Перший Перший", "ШП", date(2026, 10, 3), date(2026, 10, 4), source_sheet="ШП"),
        _period("перший перший перший", "СЗЧ", date(2026, 10, 5), None, source_sheet="СЗЧ", source_file="addon2"),
    ]

    results = amc.find_period_mismatches(timesheet, records)

    assert [(r["source_file"], r["sources"], sorted(d.day for d in r["dates"])) for r in results] == [
        ("book61", ["book61: ППД", "book61: ШП"], [1, 2, 3, 4]),
        ("addon2", ["addon2: СЗЧ"], [5, 6, 7]),
    ]


def test_check_addon_mismatches_file_lists_every_person_only_once(tmp_path):
    """Наскрізна перевірка самого mis_match.xlsx (рішення користувача: ПІБ не мають дублюватись у файлі): людина з
    ППД/ШП/СЗЧ-періодами, що перекриваються, й іще одним ШП-записом з іншим
    написанням ПІБ - ОДИН рядок "book61: ППД -> book61: ШП -> book61: СЗЧ";
    людина з двома рядками Алф2 "самостріл" - ОДИН раз у "Застереження".
    В обох аркушах - ПІДРОЗДІЛ людини з Табеля."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Підрозділ1", "Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", "ППД", "ШП", "ШП", "ШП", "СЗЧ", "РВЗ"],
        ["Підрозділ2", "Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=["ПІДРОЗДІЛ"] + _WEEK_HEADER)
    book61 = write_full_book61(tmp_path / "book61.xlsm", {
        "ППД": [[1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 8, 13), "у лік з 04.10.2026"]],
        "ШП": [
            [1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "лік", date(2026, 10, 3), date(2026, 10, 6)],
            [2, "Підрозділ1", "сержант", "ПЕРШИЙ  Перший Перший", "лік", date(2026, 9, 1), date(2026, 9, 5)],
        ],
        "СЗЧ": [[1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 10, 6)]],
        "ВП": [[1, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "від", date(2026, 10, 2), None, date(2026, 10, 5),
                "Перша частина щорічної відпустки"]],
        "Алф2": [
            [1, "ДРУГИЙ Другий Другий", "самостріл", None, None, None],
            [2, "ДРУГИЙ Другий Другий", "самостріл", None, None, None],
        ],
    })
    report = str(tmp_path / "mis_match.xlsx")

    result = check_addon_mismatches_file(oblik, [book61], report_file_name=report)

    assert result.report_path == report
    assert result.checked_files == [book61] and result.unrecognized_files == []
    wb = load_workbook(report)
    assert wb.sheetnames == ["book61", "Застереження"]
    ws = wb["book61"]
    assert [ws.cell(row=1, column=c).value for c in (1, 2, 3)] == ["Джерело", "Підрозділ", "ПІБ"]
    pib_rows = {
        normalize_name(ws.cell(row=r, column=3).value): r
        for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=3).value not in (None, "—")
    }
    pib_counts = Counter(
        normalize_name(ws.cell(row=r, column=3).value)
        for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=3).value not in (None, "—")
    )
    assert pib_counts == Counter({"ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ": 1, "ДРУГИЙ ДРУГИЙ ДРУГИЙ": 1})
    first_row = pib_rows["ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ"]
    assert ws.cell(row=first_row + 1, column=1).value == "book61: ППД -> book61: ШП -> book61: СЗЧ"
    assert ws.cell(row=first_row, column=2).value == "Підрозділ1"
    assert ws.cell(row=pib_rows["ДРУГИЙ ДРУГИЙ ДРУГИЙ"], column=2).value == "Підрозділ2"

    warn_ws = wb["Застереження"]
    warn_rows = [
        (warn_ws.cell(row=r, column=1).value, warn_ws.cell(row=r, column=2).value)
        for r in range(2, warn_ws.max_row + 1)
    ]
    assert warn_rows == [("Підрозділ2", "ДРУГИЙ Другий Другий")]


def test_find_period_mismatches_merges_multiple_records_for_the_same_person_and_sheet(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "ВП", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 7, 1), date(2026, 7, 1)),
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 7, 3), date(2026, 7, 3)),
    ]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert results[0]["dates"] == {
        date(2026, 7, 1): ("РВЗ", "ВП"),
        date(2026, 7, 3): ("РВЗ", "ВП"),
    }


def test_find_period_mismatches_treats_bpshp_vpbp_as_a_non_mismatch(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВПБП", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "БПШП", date(2026, 7, 1), date(2026, 7, 1), source_sheet="ШП")]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_treats_payment_converted_value_as_a_non_mismatch(tmp_path):
    """Реальний випадок: output/ОБЛІК.xlsx вже "для виплат" - клітинка
    каже "100_БПШП" (число 10 - аналогічно для "ВД") замість буквального
    book61-статусу "БПШП"."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "100_БПШП", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "БПШП", date(2026, 7, 1), date(2026, 7, 1), source_sheet="ШП")]

    assert amc.find_period_mismatches(timesheet, records) == []


def test_find_period_mismatches_still_flags_10_vs_navch_as_a_genuine_mismatch(tmp_path):
    """Застереження користувача: "10" еквівалентний "ВД", АЛЕ якщо book61
    очікує саме "Навч" - число 10 в Табелі МАЄ лишатись розбіжністю."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", 10, "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "Навч", date(2026, 7, 1), date(2026, 7, 1), source_sheet="ВД")]

    results = amc.find_period_mismatches(timesheet, records)

    assert len(results) == 1
    assert results[0]["dates"] == {date(2026, 7, 1): (10, "Навч")}


def test_find_period_mismatches_without_any_date_columns_overlapping_is_empty(tmp_path):
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    records = [_period("ПЕРШИЙ Перший Перший", "ВП", date(2020, 1, 1), date(2020, 1, 2))]

    assert amc.find_period_mismatches(timesheet, records) == []


# --- unmapped_people_in_roster -------------------------------------------------

def test_unmapped_people_in_roster_keeps_only_people_present_in_the_tabel(tmp_path):
    """Те саме правило, що й для періодів: кого немає в Табелі, той не
    потрапляє в mis_match.xlsx навіть як застереження."""
    path = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ"],
    ])
    timesheet = Timesheet(path, "Табель")
    unmapped_people = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "alf2_status": "самостріл"},
        {"pib_raw": "ВІДСУТНІЙ Відсутній Відсутній", "alf2_status": "самостріл"},
    ]

    kept = amc.unmapped_people_in_roster(timesheet, unmapped_people)

    assert [p["pib_raw"] for p in kept] == ["ПЕРШИЙ Перший Перший"]


def test_unmapped_people_in_roster_lists_each_person_once():
    """Алф2 може мати кілька рядків з тим самим ПІБ - у "Застереження"
    людина МАЄ бути один раз."""
    class _Roster:
        label_columns = {}

        def person_rows(self):
            return {"ПЕРШИЙ ПЕРШИЙ ПЕРШИЙ": 2}

    unmapped_people = [
        {"pib_raw": "ПЕРШИЙ Перший Перший", "alf2_status": "самостріл"},
        {"pib_raw": "ПЕРШИЙ  Перший Перший", "alf2_status": "самостріл"},
    ]

    assert len(amc.unmapped_people_in_roster(_Roster(), unmapped_people)) == 1


# --- write_addon_mismatch_report ----------------------------------------------

def test_write_addon_mismatch_report_returns_none_when_nothing_to_report(tmp_path):
    assert amc.write_addon_mismatch_report([], [], str(tmp_path / "mis_match.xlsx")) is None


def test_write_addon_mismatch_report_writes_period_mismatches_and_warnings(tmp_path):
    period_mismatches = [{
        "source_file": "book61",
        "sources": ["book61: ВП"], "roster_pidrozdil": "Підрозділ1",
        "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
        "dates": {date(2026, 7, 1): ("РВЗ", "ВП")},
    }]
    unmapped_people = [{"pib_raw": "ТРЕТІЙ Третій Третій", "alf2_status": "самостріл", "roster_pidrozdil": "Підрозділ3"}]
    output_path = str(tmp_path / "mis_match.xlsx")

    result = amc.write_addon_mismatch_report(period_mismatches, unmapped_people, output_path)

    assert result == output_path
    wb = load_workbook(output_path)
    assert wb.sheetnames == ["book61", "Застереження"]
    ws = wb["book61"]
    assert [ws.cell(row=1, column=c).value for c in (1, 2, 3)] == ["Джерело", "Підрозділ", "ПІБ"]
    assert ws.cell(row=2, column=1).value == "Табель (ОБЛІК.xlsx)"
    assert ws.cell(row=3, column=1).value == "book61: ВП"
    assert ws.cell(row=2, column=2).value == "Підрозділ1"
    assert "B2:B3" in {str(merged) for merged in ws.merged_cells.ranges}
    assert ws.cell(row=2, column=3).value == "ПЕРШИЙ Перший Перший"

    warn_ws = wb["Застереження"]
    assert [warn_ws.cell(row=1, column=c).value for c in (1, 2)] == ["Підрозділ", "ПІБ (book61)"]
    assert (warn_ws.cell(row=2, column=1).value, warn_ws.cell(row=2, column=2).value) == (
        "Підрозділ3", "ТРЕТІЙ Третій Третій",
    )


def test_write_addon_mismatch_report_shows_a_placeholder_when_the_tabel_has_no_pidrozdil(tmp_path):
    period_mismatches = [{
        "source_file": "book61",
        "sources": ["book61: ВП"], "roster_pidrozdil": None,
        "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
        "dates": {date(2026, 7, 1): ("РВЗ", "ВП")},
    }]
    output_path = str(tmp_path / "mis_match.xlsx")

    amc.write_addon_mismatch_report(period_mismatches, [], output_path)

    assert load_workbook(output_path)["book61"].cell(row=2, column=2).value == "—"


def test_write_addon_mismatch_report_writes_only_warnings_when_no_period_mismatches(tmp_path):
    unmapped_people = [{"pib_raw": "ДРУГИЙ Другий Другий", "alf2_status": "самостріл", "roster_pidrozdil": None}]
    output_path = str(tmp_path / "mis_match.xlsx")

    result = amc.write_addon_mismatch_report([], unmapped_people, output_path)

    assert result == output_path
    wb = load_workbook(output_path)
    assert "Застереження" in wb.sheetnames


def test_write_addon_mismatch_report_renders_the_sheet_chain_as_the_source_label(tmp_path):
    """Рішення користувача: колонка "Джерело" показує ЛАНЦЮЖОК
    аркушів book61 у тому порядку, в якому вони змінюють одне одного на
    шкалі людини: "book61: ППД -> book61: ШП -> book61: СЗЧ"."""
    period_mismatches = [{
        "source_file": "book61",
        "sources": ["book61: ППД", "book61: ШП", "book61: СЗЧ"], "roster_pidrozdil": "Підрозділ1",
        "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
        "dates": {date(2026, 10, 8): ("РВЗ", "СЗЧ")},
    }]
    output_path = str(tmp_path / "mis_match.xlsx")

    amc.write_addon_mismatch_report(period_mismatches, [], output_path)

    wb = load_workbook(output_path)
    ws = wb["book61"]
    assert ws.cell(row=3, column=1).value == "book61: ППД -> book61: ШП -> book61: СЗЧ"


def test_addon_files_lists_excel_books_without_lock_files(tmp_path):
    folder = tmp_path / "addons"
    folder.mkdir()
    for name in ("b.xlsx", "a.xlsm", "~$b.xlsx", "notes.txt"):
        (folder / name).write_text("x")
    assert [os.path.basename(path) for path in addon_files(str(folder))] == ["a.xlsm", "b.xlsx"]


def test_check_addon_mismatches_file_puts_each_addon_file_on_its_own_sheet(tmp_path):
    """Рішення користувача: mis_match.xlsx розділено на аркуші з
    назвами addon-файлів: шкала людини одна (book61 + ЦВС), але кожна дата -
    на аркуші файлу, чий запис її визначив, ланцюжок "Джерело" - лише цього
    файлу; звірений файл без розбіжностей - аркуш із написом. День виписки
    (ЦВС, аркуш лікування) з РВЗ у Табелі - не розбіжність. Книга, тип якої не
    розпізнано (чи це взагалі не Excel), пропускається й повертається окремо."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Підрозділ1", "Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", "РВЗ", "ШП", "ВП", "РВЗ", "РВЗ", "РВЗ"],
        ["Підрозділ2", "Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=["ПІДРОЗДІЛ"] + _WEEK_HEADER)
    book61 = write_full_book61(tmp_path / "book61.xlsm", {
        "ППД": [[1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 8, 13), "у лік з 03.10.2026"]],
    })
    cvs = write_cvs(tmp_path / "cvs.xlsx", {
        MEDICAL_SHEET: [[None, 1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "соматика", datetime(2026, 10, 3),
                         "Лікарня", "№1", datetime(2026, 10, 5)]],
    })
    cvs2 = write_cvs(tmp_path / "cvs2.xlsx", {
        VLK_SHEET: [[1, "ДРУГИЙ Другий Другий", "матрос", None, datetime(2026, 10, 2), None, "придатний"]],
    })
    junk = tmp_path / "junk.xlsx"
    junk_wb = Workbook()
    junk_wb.worksheets[0].append(["ПІБ", "Щось інше"])
    junk_wb.save(str(junk))
    broken = tmp_path / "broken.xlsx"
    broken.write_text("не книга")
    report = str(tmp_path / "mis_match.xlsx")

    result = check_addon_mismatches_file(oblik, [book61, cvs, cvs2, str(junk), str(broken)], report_file_name=report)

    assert result.checked_files == [book61, cvs, cvs2]
    assert result.unrecognized_files == [str(junk), str(broken)]
    # 04.10 "ВП" у Табелі - ще й на аркуші book61: у book61 немає запису "ВП"
    # (ADDON_REQUIRED_BOOK61_SHEETS).
    assert [(r["source_file"], r["roster_pib_raw"], r["sources"]) for r in result.period_mismatches] == [
        ("book61", "ПЕРШИЙ Перший Перший", ["book61: ППД", "book61: ВП"]),
        ("cvs", "ПЕРШИЙ Перший Перший", ["cvs: СОМАТИКА-ПОРАНЕНІ"]),
    ]
    bad_days = [
        [d.day for d, (actual, expected) in record["dates"].items() if not amc._statuses_match(expected, actual)]
        for record in result.period_mismatches
    ]
    assert bad_days == [[2, 4], [4]]

    wb = load_workbook(report)
    assert wb.sheetnames == ["book61", "cvs", "cvs2"]
    assert wb["book61"].cell(row=3, column=1).value == "book61: ППД -> book61: ВП"
    assert wb["cvs"].cell(row=3, column=1).value == "cvs: СОМАТИКА-ПОРАНЕНІ"
    header_dates = [wb["cvs"].cell(row=1, column=c).value for c in range(4, wb["cvs"].max_column + 1)]
    assert [value.day for value in header_dates if isinstance(value, datetime)] == [3, 4, 5]
    assert len(header_dates) == 3
    assert wb["cvs2"].cell(row=2, column=1).value == "Розбіжностей з цим файлом немає."


def test_sheet_titles_follow_excel_rules_and_stay_unique():
    used = set()
    assert amc._sheet_title("a/b:c*?[x]\\y", used) == "a_b_c___x__y"
    assert len(amc._sheet_title("Б" * 40, used)) == 31
    assert amc._sheet_title("book61", used) == "book61"
    assert amc._sheet_title("BOOK61", used) == "BOOK61 (2)"


def test_non_combat_treatment_accepts_any_of_its_statuses():
    """Рішення користувача: не бойове лікування з книги ЦВС - у
    Табелі ШП, ЛМР або ВПСЗ (і їхні форми в ОБЛІК для виплат, напр. 10_МЕД);
    бойове (БПШП) чи РВЗ - розбіжність; в обидва боки (звірка книг між собою)."""
    from constants import ADDON_NON_COMBAT_TREATMENT_STATUS as non_combat
    for actual in ("ШП", "ЛМР", "ВПСЗ"):
        assert amc._statuses_match(non_combat, actual) and amc._statuses_match(actual, non_combat)
    assert amc._statuses_match(non_combat, "10_МЕД")
    for actual in ("БПШП", "РВЗ", 100):
        assert not amc._statuses_match(non_combat, actual)
    assert not amc._statuses_match("БПШП", non_combat)


def test_open_non_combat_treatment_takes_the_end_from_another_books_hospital_stay():
    """Відкритий запис ЦВС "ШП/ЛМР/ВПСЗ" бере кінець того самого епізоду з
    іншої книги (book61 "ШП"), як раніше відкритий "ШП"."""
    from constants import ADDON_NON_COMBAT_TREATMENT_STATUS as non_combat
    cvs = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": non_combat, "start_date": date(2026, 10, 10),
           "end_date": None, "source_sheet": "СОМАТИКА", "arrival_date": None}
    book61 = {"pib_raw": "ПЕРШИЙ Перший Перший", "status": "ШП", "start_date": date(2026, 10, 10),
              "end_date": date(2026, 10, 15), "source_sheet": "ШП", "arrival_date": date(2026, 10, 15)}
    assert amc._known_end_of_open_period(cvs, [cvs, book61]) == date(2026, 10, 15)


def test_leave_in_tabel_must_be_covered_by_a_book61_vp_record(tmp_path):
    """Рішення користувача (реальний випадок): у Табелі ВПСЗ до
    кінця тижня, а запис "ВП" book61 закінчився раніше - дні після нього -
    розбіжність на аркуші book61 ("немає запису на аркуші ВП"), хоча жодна
    книга в ці дні нічого не очікує. Людина з відпусткою в Табелі й без
    жодного запису в addons - теж. Без book61 серед книг - не перевіряється."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВПСЗ", "ВПСЗ", "ВПСЗ", "ВПСЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ВПСЗ", date(2026, 9, 13), date(2026, 10, 2)),
        _period("ТРЕТІЙ Третій Третій", "ВП", date(2026, 9, 25), date(2026, 10, 2)),
    ]

    results = amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"})

    no_record = "немає запису на аркуші ВП"
    assert {r["roster_pib_raw"]: {d.day: pair for d, pair in r["dates"].items() if pair[1] == no_record}
            for r in results} == {
        "ПЕРШИЙ Перший Перший": {3: ("ВПСЗ", no_record), 4: ("ВПСЗ", no_record)},
        "ДРУГИЙ Другий Другий": {2: ("ВП", no_record), 3: ("ВП", no_record)},
    }
    assert all(r["source_file"] == "book61" and "book61: ВП" in r["sources"] for r in results)
    assert amc.find_period_mismatches(timesheet, records) == []  # без book61 - не перевіряється


def test_tracked_statuses_without_any_addon_record_are_shown(tmp_path):
    """Рішення користувача: "теж їх відображати, враховуючи
    попередні правила": статус Табеля, який веде книга addons, у день, коли
    ЖОДНА книга нічого не каже, - розбіжність "немає запису на аркуші ..." на
    аркуші відповідної книги (ШП - book61 "ШП", Адап - журнал; форма для
    виплат 100_БЗ - book61 "Алф2"); запис будь-якої книги цього дня -
    звичайна звірка (ЦВС "ШП/ЛМР/ВПСЗ" і ШП у Табелі - збіг); статус, якого не
    веде жодна книга (РТГр), і книга, якої немає серед звірених, - не
    перевіряються."""
    from constants import ADDON_NON_COMBAT_TREATMENT_STATUS
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ШП", "ШП", "ШП", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "Адап", "Адап", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада3", "матрос", "ТРЕТІЙ Третій Третій", "100_БЗ", "100_БЗ", "100_БЗ", "100_БЗ", "100_БЗ", "100_БЗ", "100_БЗ"],
        ["Посада4", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "РТГр", "РТГр", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    records = [
        # ПЕРШИЙ: 01-02 - ЦВС (не бойове) - збіг; 03 - жодного запису.
        _period("ПЕРШИЙ Перший Перший", ADDON_NON_COMBAT_TREATMENT_STATUS, date(2026, 10, 1), date(2026, 10, 2),
                "СОМАТИКА", "cvs"),
    ]
    no_record = lambda sheet: f"немає запису на аркуші {sheet}"

    results = amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61", "journal": "journal"})

    assert sorted((r["source_file"], r["roster_pib_raw"], sorted((d.day, e) for d, (a, e) in r["dates"].items()))
                  for r in results) == [
        ("book61", "ПЕРШИЙ Перший Перший", [(3, no_record("ШП"))]),
        ("book61", "ТРЕТІЙ Третій Третій", [(day, no_record("Алф2")) for day in range(1, 8)]),
        ("journal", "ДРУГИЙ Другий Другий", [(1, no_record("Адаптація")), (2, no_record("Адаптація"))]),
    ]
    only_book61 = amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"})
    assert "ДРУГИЙ Другий Другий" not in {r["roster_pib_raw"] for r in only_book61}  # журналу немає


def _check_transfer_on_alf2(tmp_path, order_day):
    """Табель: ПРВДБ лише 07.10. Книга: той самий статус на Алф2, дата наказу -
    справжня дата Excel. Повертає розбіжні дні."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "ПРВДБ"],
    ], header=_WEEK_HEADER)
    book61 = write_full_book61(tmp_path / "book61.xlsm", {"Алф2": [
        [1, "ПЕРШИЙ Перший Перший", "переведений в середині бригади", date(2026, 9, 11), None, datetime(2026, 10, order_day)],
    ]})
    result = check_addon_mismatches_file(oblik, [book61], report_file_name=str(tmp_path / "mis_match.xlsx"))
    return [
        (day.day, actual, expected) for record in result.period_mismatches
        for day, (actual, expected) in record["dates"].items() if not amc._statuses_match(expected, actual)
    ]


def test_transfer_with_a_real_date_cell_on_alf2_is_not_reported_as_missing(tmp_path):
    """Реальний випадок: у Табелі ПРВДБ 31.10, на Алф2 "переведений в середині
    бригади" і дата 31.10.2026, а звіт писав "немає запису на аркуші Алф2"."""
    assert _check_transfer_on_alf2(tmp_path, order_day=7) == []


def test_transfer_date_on_alf2_is_compared_with_the_timesheet(tmp_path):
    """Запис справді звіряється, а не просто гасить зворотну перевірку."""
    assert _check_transfer_on_alf2(tmp_path, order_day=6) == [(6, "РВЗ", "ПРВДБ")]


def test_killed_on_alf2_allows_bz_until_the_death_certificate(tmp_path):
    """Правило книги ЦВС (рішення користувача): "Загиб" - лише з
    дати лікарського свідоцтва. Реальний випадок: Алф2 - "загибель" з 2024
    (без свідоцтва), ЦВС - Загиб з 05.10 (свідоцтво), у Табелі БЗ до 04.10,
    далі Загиб - не розбіжність. Число в Табелі без відомого статусу за ним
    (10 - і ВД, і ППД...) у зворотну перевірку не йде."""
    from constants import ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "БЗ", "БЗ", "БЗ", "БЗ", "Загиб", "Загиб", "Загиб"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", 10, 10, "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [
        _period("ПЕРШИЙ Перший Перший", ADDON_KILLED_WITHOUT_CERTIFICATE_STATUS, date(2024, 8, 6), None, "Алф2"),
        _period("ПЕРШИЙ Перший Перший", "Загиб", date(2026, 10, 5), None, "Загиблі-200", "cvs"),
    ]

    assert amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind={"book61": "book61"}) == []


def test_report_has_no_gaps_dates_compared_elsewhere_show_the_tabel_value_in_grey(tmp_path):
    """Рішення користувача (реальний випадок: ВПСЗ | порожньо |
    ВПСЗ): дата, яку визначила інша книга (її аркуш) чи жодна, - не порожні
    клітинки, а значення Табеля в сірій об'єднаній клітинці."""
    d1, d2, d3 = date(2026, 10, 8), date(2026, 10, 9), date(2026, 10, 10)
    tabel = {d1: "ВПСЗ", d2: "10_МЕД", d3: "ВПСЗ"}
    records = [
        {"source_file": "book61", "sources": ["book61: ВП"], "roster_pidrozdil": "Підрозділ1",
         "roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
         "dates": {d1: ("ВПСЗ", "ВПСЗ"), d3: ("ВПСЗ", "ВП")}, "tabel_values": tabel},
        {"source_file": "book61", "sources": ["book61: ВП"], "roster_pidrozdil": "Підрозділ1",
         "roster_pib_raw": "ДРУГИЙ Другий Другий", "addon_pib_raw": "ДРУГИЙ Другий Другий",
         "dates": {d2: ("РВЗ", "ВП")}, "tabel_values": {d1: None, d2: "РВЗ", d3: "РВЗ"}},
    ]
    path = str(tmp_path / "mis_match.xlsx")

    amc.write_addon_mismatch_report(records, [], path, source_files=["book61"])

    ws = load_workbook(path)["book61"]
    assert [ws.cell(row=1, column=c).value for c in (4, 5, 6)] == [datetime(2026, 10, d) for d in (8, 9, 10)]
    gap = ws.cell(row=2, column=5)
    assert (gap.value, gap.fill.fgColor.rgb) == ("10_МЕД", "FFF2F2F2")
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert "E2:E3" in merged
    assert ws.cell(row=4, column=4).value == "—" and "D4:D5" in merged  # Табель порожній
    assert all(ws.cell(row=row, column=col).value is not None for row in (2, 4) for col in (4, 5, 6))


def _episode(status, start, end, source_file, **extra):
    return {"pib_raw": "ПЕРШИЙ Перший Перший", "status": status, "start_date": start, "end_date": end,
            "source_sheet": "аркуш", "source_file": source_file, **extra}


def test_open_period_takes_the_end_of_the_same_episode_only():
    """Рішення користувача ("продовження статусів"): відкритий
    період бере кінець того самого епізоду з іншої книги (закінчився пізніше
    чи того самого дня з документом), але НЕ попереднього епізоду, що
    закінчився в день його початку (переведення з лікарні до лікарні 16.10), і
    НЕ одноденного запису без документа (ЦВС ВЛК лише з датою проходження)."""
    open_bpshp = _episode("БПШП", date(2026, 10, 16), None, "book61")
    previous = _episode("БПШП", date(2026, 9, 24), date(2026, 10, 16), "cvs")
    assert amc._known_end_of_open_period(open_bpshp, [open_bpshp, previous]) is None

    open_vlk = _episode("ВЛК", date(2026, 10, 13), None, "book61")
    one_day = _episode("ВЛК", date(2026, 10, 13), date(2026, 10, 13), "cvs", end_assumed=True)
    assert amc._known_end_of_open_period(open_vlk, [open_vlk, one_day]) is None

    same_day_document = _episode("ВЛК", date(2026, 10, 13), date(2026, 10, 13), "cvs", end_assumed=False)
    later_document = _episode("ВЛК", date(2026, 10, 13), date(2026, 10, 20), "cvs", end_assumed=False)
    assert amc._known_end_of_open_period(open_vlk, [open_vlk, same_day_document]) == date(2026, 10, 13)
    assert amc._known_end_of_open_period(open_vlk, [open_vlk, later_document]) == date(2026, 10, 20)


def _names_table(ws):
    """Рядки таблиці "Невідповідності ПІБ" аркуша звіту (без назви й шапки)."""
    title_row = next(row for row in range(1, ws.max_row + 1) if ws.cell(row=row, column=1).value == "Невідповідності ПІБ")
    assert [ws.cell(row=title_row + 1, column=col).value for col in range(1, 6)] == [
        "Джерело", "Підрозділ", "ПІБ в ОБЛІК", "ПІБ у файлі", "Період у файлі",
    ]
    return [[ws.cell(row=row, column=col).value for col in range(1, 6)] for row in range(title_row + 2, ws.max_row + 1)]


def test_a_name_spelled_differently_in_an_addon_is_matched_and_reported(tmp_path):
    """Рішення користувача (реальний випадок: в ОБЛІК "В’...", у
    book61 - без апострофа; book61-записи цієї людини раніше не звірялись
    зовсім): людину зіставлено попри апостроф (дати звіряються), а різниця
    написання - у таблиці ПІБ навіть без розбіжних дат (ПІБ файлу червоним)."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "П'ЯТИЙ П'ятий П'ятий", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [_period("ПЯТИЙ Пятий Пятий", "ВП", date(2026, 9, 25), date(2026, 10, 2))]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records)

    assert [(r["source_file"], r["roster_pib_raw"], r["addon_pib_raw"]) for r in results] == [
        ("book61", "П'ЯТИЙ П'ятий П'ятий", "ПЯТИЙ Пятий Пятий"),
    ]
    assert all(amc._statuses_match(e, a) for a, e in results[0]["dates"].values())  # дати звіряються й збігаються
    report = str(tmp_path / "mis_match.xlsx")
    amc.write_addon_mismatch_report(results, [], report, source_files=["book61"])
    ws = load_workbook(report)["book61"]
    assert _names_table(ws) == [
        ["book61: ВП", "—", "П'ЯТИЙ П'ятий П'ятий", "ПЯТИЙ Пятий Пятий", "25.09.2026 - 02.10.2026"],
    ]
    assert ws.cell(row=3, column=4).fill.fgColor.rgb == "FFFFC7CE"


def test_vacation_without_a_destination_is_a_yellow_mismatch_in_the_tabel_month(tmp_path):
    """Рішення користувача: відпустка book61 без "Куди вибув"
    (порожньо чи "?") - розбіжність, навіть коли статуси збігаються; у
    mis_match.xlsx - дні цієї відпустки в місяці Табеля просто жовтим
    (об'єднана клітинка зі значенням Табеля, без тексту), а в "Джерело" -
    примітка. Та сама відпустка з "Куди вибув" - не розбіжність."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "ПЕРШИЙ Перший Перший", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ВП", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [
        {**_period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 9, 25), date(2026, 10, 2)), "destination_missing": True},
        {**_period("ДРУГИЙ Другий Другий", "ВП", date(2026, 9, 25), date(2026, 10, 2)), "destination_missing": False},
    ]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records)

    assert [(r["roster_pib_raw"], r["no_destination"]) for r in results] == [
        ("ПЕРШИЙ Перший Перший", [date(2026, 10, 1), date(2026, 10, 2)]),
    ]
    report = str(tmp_path / "mis_match.xlsx")
    amc.write_addon_mismatch_report(results, [], report, source_files=["book61"])
    ws = load_workbook(report)["book61"]
    assert ws.cell(row=3, column=1).value == "book61: ВП; ВП: не вказано «Куди вибув»"
    merged = {str(rng) for rng in ws.merged_cells.ranges}
    assert [ws.cell(row=2, column=col).value for col in (4, 5)] == ["ВП", "ВП"]
    assert {"D2:D3", "E2:E3"} <= merged  # одна клітинка, без тексту внизу
    assert {ws.cell(row=2, column=col).fill.fgColor.rgb for col in (4, 5)} == {"FFFFEB9C"}


def test_report_pib_column_is_wide_enough_and_left_aligned(tmp_path):
    """Рішення користувача: ПІБ у mis_match.xlsx - повністю."""
    from content.report_pib_column import pib_column_width
    name = "ВІСІМНАДЦЯТИЙ Вісімнадцятий Вісімнадцятий"
    records = [{"source_file": "book61", "sources": ["book61: ВП"], "roster_pidrozdil": "Підрозділ1",
                "roster_pib_raw": name, "addon_pib_raw": name, "dates": {date(2026, 10, 1): ("РВЗ", "ВП")},
                "tabel_values": {date(2026, 10, 1): "РВЗ"}}]
    path = str(tmp_path / "mis_match.xlsx")

    amc.write_addon_mismatch_report(records, [], path, source_files=["book61"])

    ws = load_workbook(path)["book61"]
    assert ws.column_dimensions["C"].width == pib_column_width([name]) > 30
    assert ws.cell(row=2, column=3).alignment.horizontal == "left"


def test_vacation_without_a_destination_alone_is_only_a_warning():
    """Рішення користувача: у терміналі жовтим, це не помилка:
    лише відпустка без "Куди вибув", а статуси збігаються й ПІБ те саме.
    Будь-яка розбіжність статусу чи інше написання ПІБ - вже помилка."""
    base = {"roster_pib_raw": "ПЕРШИЙ Перший Перший", "addon_pib_raw": "ПЕРШИЙ Перший Перший",
            "dates": {date(2026, 10, 1): ("ВП", "ВП")}, "no_destination": [date(2026, 10, 1)]}
    assert amc.is_warning_only(base)
    assert not amc.is_warning_only({**base, "dates": {date(2026, 10, 1): ("РВЗ", "ВП")}})
    assert not amc.is_warning_only({**base, "addon_pib_raw": "ПЕРШИЙ Перший Перши"})
    assert not amc.is_warning_only({**base, "no_destination": []})


def test_a_name_spelled_differently_is_reported_even_without_dates_in_the_tabel(tmp_path):
    """Запис книги цілком поза датами Табеля (лише вересень) - дат звіряти
    нема, але інше написання ПІБ однаково показується."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "сержант", "П'ЯТИЙ П'ятий П'ятий", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [_period("ПЯТИЙ Пятий Пятий", "ВП", date(2026, 9, 1), date(2026, 9, 5))]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records)

    assert [(r["addon_pib_raw"], r["sources"], r["dates"]) for r in results] == [("ПЯТИЙ Пятий Пятий", ["book61: ВП"], {})]


def test_a_day_without_a_document_does_not_end_an_open_treatment(tmp_path):
    """Реальний випадок: БПШП з 01.10 без виписки, ЦВС - ВЛК 05.10 лише з
    датою проходження (end_assumed). Лікування після ВЛК триває, а не
    лишається без жодного запису ("немає запису на аркуші ШП")."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "БПШП", "БПШП", "БПШП", "БПШП", "ВЛК", "БПШП", "БПШП"],
    ], header=_WEEK_HEADER)
    records = [
        _period("ПЕРШИЙ Перший Перший", "БПШП", date(2026, 10, 1), None, "СОМАТИКА-ПОРАНЕНІ", "cvs"),
        {**_period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 5), date(2026, 10, 5), "ВЛК", "cvs"), "end_assumed": True},
    ]

    assert amc.find_period_mismatches(Timesheet(oblik, "Табель"), records,
                                      files_by_kind={"book61": "book61", "cvs": "cvs"}) == []


def test_returned_from_szch_accepts_any_status_of_someone_back_on_duty():
    """Рішення користувача: "Повер" (book61, повернувся з СЗЧ) - збіг з
    будь-яким статусом Табеля, крім ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES
    (і їхніх форм для виплат, напр. "100_БЗ")."""
    from constants import ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES

    for actual in ("РВЗ", 100, 30, "ВП", "ШП", "Повер"):
        assert amc._statuses_match("Повер", actual), actual
    for actual in (*ADDON_NOT_RETURNED_AFTER_SZCH_STATUSES, *amc._payment_converted_forms("БЗ")):
        assert not amc._statuses_match("Повер", actual), actual


def test_a_misspelled_name_in_a_book_is_matched_and_shown_as_a_spelling_difference(tmp_path):
    """Реальний випадок: у Табелі ім'я з опискою, у book61 - інакше (прізвище
    й по батькові ті самі). Записи книги приєднуються до людини Табеля, дати
    звіряються як звичайно, а звіт показує обидва написання ПІБ. Інше ім'я
    (не описка) і людина, яку книга вже знає під табельним ПІБ, - ні."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ДРУГИЙ Другий Другий", "СЗЧ", "СЗЧ", "СЗЧ", "СЗЧ", "СЗЧ", "СЗЧ", "СЗЧ"],
        ["Посада2", "матрос", "ТРЕТІЙ Третій Третій", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада3", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [
        _period("ДРУГИЙ Дугий Другий", "СЗЧ", date(2026, 10, 1), None, "СЗЧ"),
        _period("ТРЕТІЙ Сьомий Третій", "СЗЧ", date(2026, 10, 1), None, "СЗЧ"),  # інша людина
        _period("ЧЕТВЕРТИЙ Четвертий Четвертий", "ВП", date(2026, 9, 1), date(2026, 9, 5)),
        _period("ЧЕТВЕРТИЙ Четвртий Четвертий", "СЗЧ", date(2026, 10, 1), None, "СЗЧ"),
    ]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records)

    assert [(r["roster_pib_raw"], r["addon_pib_raw"], r["dates"]) for r in results] == [
        ("ДРУГИЙ Другий Другий", "ДРУГИЙ Дугий Другий", {
            date(2026, 10, day): ("СЗЧ", "СЗЧ") for day in range(1, 8)
        }),
    ]


def _trip(pib_raw, start, planned_end):
    """Рядок аркуша "ВД" без фактичного кінця (content/addon_reader.py)."""
    return {**_period(pib_raw, "ВД", start, None, "ВД"), "planned_end": planned_end, "arrival_date": planned_end}


def test_open_trip_is_confirmed_only_until_its_planned_end(tmp_path):
    """Реальний випадок: у book61 ВД з "Дата кінця" 04.10 без дати прибуття,
    далі жодного запису, а в Табелі й далі ВД. Дні після плану мають бути
    розбіжністю, а не збігом з відкритим періодом."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВД", "ВД", "ВД", "ВД", "ВД", "ВД", "ВД"],
    ], header=_WEEK_HEADER)
    records = [_trip("ПЕРШИЙ Перший Перший", date(2026, 9, 14), date(2026, 10, 4))]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind={"book61": "book61"})

    no_record = "немає запису на аркуші ВД"
    assert [(r["source_file"], r["sources"], r["dates"]) for r in results] == [("book61", ["book61: ВД"], {
        **{date(2026, 10, day): ("ВД", "ВД") for day in range(1, 5)},
        **{date(2026, 10, day): ("ВД", no_record) for day in range(5, 8)},
    })]


def test_trip_extended_on_the_next_row_has_no_gap(tmp_path):
    """Реальний випадок: "продовжено" - наступний рядок ВД іде впритул, і
    розбіжності немає."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВД", "ВД", "ВД", "ВД", "ВД", "ВД", "ВД"],
    ], header=_WEEK_HEADER)
    records = [
        _trip("ПЕРШИЙ Перший Перший", date(2026, 9, 14), date(2026, 10, 3)),
        _trip("ПЕРШИЙ Перший Перший", date(2026, 10, 4), date(2026, 10, 20)),
    ]

    assert amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind={"book61": "book61"}) == []


def test_return_at_the_planned_end_is_not_a_mismatch(tmp_path):
    """Дати прибуття в книзі немає, а "Дата кінця" здебільшого і є цим днем: РВЗ
    у Табелі того дня - не розбіжність, як і в будь-який день прибуття. Далі
    людина на службі, запис книги не потрібен."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВД", "ВД", 30, 100, 100, 100, 100],
    ], header=_WEEK_HEADER)
    records = [_trip("ПЕРШИЙ Перший Перший", date(2026, 9, 14), date(2026, 10, 3))]

    assert amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind={"book61": "book61"}) == []


def test_planned_end_does_not_bring_a_trip_back_after_a_later_event():
    """План не закриває запис: як і кожен відкритий період, він закінчується
    перед пізнішою подією й після неї не продовжується (лікування посеред
    відрядження без дати прибуття)."""
    trip = _trip("ПЕРШИЙ Перший Перший", date(2026, 10, 1), date(2026, 10, 7))
    treatment = _period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 10, 3), date(2026, 10, 4), "ШП")
    date_columns = [(day, date(2026, 10, day)) for day in range(1, 8)]

    timeline = amc._paint_timeline([trip, treatment], date_columns, date(2026, 10, 7))

    assert {day.day: record["status"] for day, record in timeline.items()} == {1: "ВД", 2: "ВД", 3: "ШП", 4: "ШП"}


@pytest.mark.parametrize("status, source_file, sheet", [
    ("ВД", "book61", "ВД"),
    ("ППД", "book61", "ППД"),
    ("БЗВП", "journal", "БЗВП"),
    ("Адап", "journal", "Адаптація"),
])
def test_number_without_a_book_record_is_reported_by_the_status_behind_it(tmp_path, status, source_file, sheet):
    """Реальні випадки: у ОБЛІК для виплат усі ці статуси - число 10, і людину
    у ВД чи ППД без жодного запису звірка не бачила. Шукаємо запис за статусом
    за числом і показуємо цей статус, бо 10 - лише вид оплати."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", 10, 10, "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    timesheet.statuses_behind_numbers = {(2, date(2026, 10, day)): status for day in (1, 2)}

    results = amc.find_period_mismatches(timesheet, [], files_by_kind={"book61": "book61", "journal": "journal"})

    no_record = f"немає запису на аркуші {sheet}"
    assert [(r["source_file"], r["sources"], r["dates"]) for r in results] == [(source_file, [f"{source_file}: {sheet}"], {
        date(2026, 10, 1): (status, no_record),
        date(2026, 10, 2): (status, no_record),
    })]


@pytest.mark.parametrize("cell, behind", [("НОВ", None), (10, "НОВ")])
def test_new_arrival_is_not_looked_up_in_the_books(tmp_path, cell, behind):
    """Рішення користувача: НОВ - людина ще без БЗВП чи Адаптації, тож запису в
    журналі в неї й не має бути. Реальний випадок: у Табелі НОВ до 21.10,
    у журналі Адаптація з 22.10, а звіт писав "немає запису на аркуші
    Адаптація" за дні до неї."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", cell, cell, "Адап", "Адап", "Адап", "Адап", "Адап"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    if behind:
        timesheet.statuses_behind_numbers = {(2, date(2026, 10, day)): behind for day in (1, 2)}
    records = [_period("ПЕРШИЙ Перший Перший", "Адап", date(2026, 10, 3), date(2026, 10, 9), "Адаптація", "journal")]

    assert amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61", "journal": "journal"}) == []


def test_new_arrival_is_still_compared_with_a_record_of_that_day(tmp_path):
    """Шукати запис не треба, але запис книги на цей день звіряється, як завжди."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "НОВ", "НОВ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [_period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 10, 1), date(2026, 10, 2), "ШП")]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind={"book61": "book61"})

    assert [r["dates"] for r in results] == [{date(2026, 10, 1): ("НОВ", "ШП"), date(2026, 10, 2): ("НОВ", "ШП")}]


def test_number_covered_by_a_book_record_needs_nothing_more(tmp_path):
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", 10, 10, 10, 10, 10, 10, 10],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    timesheet.statuses_behind_numbers = {(2, date(2026, 10, day)): "ППД" for day in range(1, 8)}
    records = [_period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 9, 1), None, "ППД")]

    assert amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"}) == []


def test_number_with_no_status_behind_it_needs_no_record(tmp_path):
    """Вид оплати людини на службі (100/30/70/170): статусу за ним немає, і
    жодна книга не мусить такого дня щось казати."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", 100, 30, 70, 170, 100, 30, 70],
    ], header=_WEEK_HEADER)

    assert amc.find_period_mismatches(Timesheet(oblik, "Табель"), [],
                                      files_by_kind={"book61": "book61", "journal": "journal"}) == []


_EVENING_DAY = date(2026, 10, 3)


def _roster_with_evening_arrival(tmp_path, day_value, behind=None):
    """ВЛК, ВЛК, день прибуття "після 17.00" (day_value), далі РВЗ."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", "ВЛК", day_value, "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    timesheet.after_hours_arrivals = {(2, _EVENING_DAY)}
    if behind:
        timesheet.statuses_behind_numbers = {(2, _EVENING_DAY): behind}
    return timesheet


@pytest.mark.parametrize("day_value, behind", [("ППД", None), (10, "ППД")])
def test_ppd_on_an_evening_arrival_day_needs_no_book_record(tmp_path, day_value, behind):
    """Реальний випадок: ВЛК у book61 закінчилась напередодні, людина прибула
    "після 17.00", і ППД цього дня дає сам рапорт (рішення користувача). На
    аркуші ППД такого запису не буває, тож і розбіжності немає."""
    timesheet = _roster_with_evening_arrival(tmp_path, day_value, behind)
    records = [_period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 9, 20), date(2026, 10, 2), "ВЛК")]

    assert amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"}) == []


@pytest.mark.parametrize("day_value, behind", [("ППД", None), (10, "ППД")])
def test_book_period_ending_on_an_evening_arrival_day_agrees_with_ppd(tmp_path, day_value, behind):
    """Реальний випадок: той самий запис рапорту, але ВЛК у book61 до дня
    прибуття включно. Книга тут теж права: день той самий, різниця лише в
    правилі "після 17.00"."""
    timesheet = _roster_with_evening_arrival(tmp_path, day_value, behind)
    records = [{
        **_period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 9, 20), _EVENING_DAY, "ВЛК"),
        "arrival_date": _EVENING_DAY,
    }]

    assert amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"}) == []


def test_ppd_without_an_evening_arrival_still_needs_a_book_record(tmp_path):
    timesheet = _roster_with_evening_arrival(tmp_path, "ППД")
    timesheet.after_hours_arrivals = set()

    results = amc.find_period_mismatches(timesheet, [], files_by_kind={"book61": "book61"})

    assert [r["dates"][_EVENING_DAY] for r in results] == [("ППД", "немає запису на аркуші ППД")]


def test_evening_arrival_excuses_only_ppd_of_that_day(tmp_path):
    """Вписане вручну звіряється, як завжди, і наступні дні теж: книга, де
    людина далі на ВЛК, суперечить рапорту."""
    timesheet = _roster_with_evening_arrival(tmp_path, "ШП")
    records = [_period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 9, 20), date(2026, 10, 4), "ВЛК")]

    results = amc.find_period_mismatches(timesheet, records, files_by_kind={"book61": "book61"})

    wrong_days = [day for day, (actual, expected) in results[0]["dates"].items() if not amc._statuses_match(expected, actual)]
    assert wrong_days == [_EVENING_DAY, date(2026, 10, 4)]


# --- журнал лікування ---------------------------------------------------------

_PERSON = "ПЕРШИЙ Перший Перший"
_JOURNAL = "журнал"
_WITH_JOURNAL = {ADDON_TREATMENT_JOURNAL_KIND: _JOURNAL}
_NO_TREATMENT_RECORD = "немає запису в журналі лікування"


def _treatment(start, end, pib_raw=_PERSON):
    return {
        **_period(pib_raw, ADDON_TREATMENT_JOURNAL_STATUS, start, end, "2026", _JOURNAL),
        "arrival_date": end, "end_before_start": False,
    }


def _week(tmp_path, *days):
    """Табель ПЕРШОГО на 01-07.10: days, далі РВЗ."""
    cells = [*days, *["РВЗ"] * (7 - len(days))]
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [["Посада1", "матрос", _PERSON, *cells]], header=_WEEK_HEADER)
    return Timesheet(oblik, "Табель")


def _october(day):
    return date(2026, 10, day)


@pytest.mark.parametrize("value", ["ШП", "БПШП", "ЛМР", "ЗВФН", "100_БПШП", "10_МЕД"])
def test_treatment_statuses_agree_with_the_journal(tmp_path, value):
    """Статусу в журналі немає, тож підходить будь-який з чотирьох статусів чи
    його форма в ОБЛІК для виплат."""
    results = amc.find_period_mismatches(
        _week(tmp_path, value, value, value), [_treatment(_october(1), _october(3))], files_by_kind=_WITH_JOURNAL,
    )

    assert results == []


@pytest.mark.parametrize("status", ["ШП", "БПШП", "ЛМР", "ЗВФН"])
def test_treatment_day_without_a_journal_record_is_reported(tmp_path, status):
    """Рішення користувача: людина зі статусом лікування має бути в журналі."""
    results = amc.find_period_mismatches(_week(tmp_path, status, status), [], files_by_kind=_WITH_JOURNAL)

    assert [(r["source_file"], r["sources"], r["dates"]) for r in results] == [(_JOURNAL, ["журнал лікування"], {
        _october(1): (status, _NO_TREATMENT_RECORD),
        _october(2): (status, _NO_TREATMENT_RECORD),
    })]


@pytest.mark.parametrize("behind, reported", [("ШП", True), ("ЛМР", True), ("ВЛК", False)])
def test_payment_category_is_looked_up_by_the_status_behind_it(tmp_path, behind, reported):
    """10_МЕД - це і ШП, і ЛМР, і ВЛК, а ВЛК журнал лікування не веде. У
    розбіжності показується статус."""
    timesheet = _week(tmp_path, "10_МЕД")
    timesheet.statuses_behind_categories = {(2, _october(1)): behind}

    results = amc.find_period_mismatches(timesheet, [], files_by_kind=_WITH_JOURNAL)

    assert [r["dates"] for r in results] == ([{_october(1): (behind, _NO_TREATMENT_RECORD)}] if reported else [])


def test_category_on_a_file_without_reports_is_shown_as_is(tmp_path):
    """Без рапортів статусу за категорією немає, а сама вона належить до форм ШП."""
    results = amc.find_period_mismatches(_week(tmp_path, "10_МЕД"), [], files_by_kind=_WITH_JOURNAL)

    assert [r["dates"] for r in results] == [{_october(1): ("10_МЕД", _NO_TREATMENT_RECORD)}]


def test_status_behind_a_category_is_not_used_for_a_plain_status(tmp_path):
    """Файл стройової: у клітинці сам статус ШП, хоч у виплатах за нього БПШП."""
    timesheet = _week(tmp_path, "ШП")
    timesheet.statuses_behind_categories = {(2, _october(1)): "БПШП"}

    results = amc.find_period_mismatches(timesheet, [], files_by_kind=_WITH_JOURNAL)

    assert [r["dates"] for r in results] == [{_october(1): ("ШП", _NO_TREATMENT_RECORD)}]


def test_journal_day_with_another_status_is_a_mismatch(tmp_path):
    """Журнал каже, що людина лікується, а в Табелі вона на службі."""
    results = amc.find_period_mismatches(
        _week(tmp_path, "ШП", "РВЗ", "ШП"), [_treatment(_october(1), _october(3))], files_by_kind=_WITH_JOURNAL,
    )

    assert [r["dates"][_october(2)] for r in results] == [("РВЗ", ADDON_TREATMENT_JOURNAL_STATUS)]


@pytest.mark.parametrize("next_status", ["РВЗ", 30, "ППД", "СЗЧ", "Смерть"])
def test_any_status_fits_the_discharge_day(tmp_path, next_status):
    """Реальні випадки: у день виписки людина вже в РВЗ, у ППД чи в СЗЧ."""
    results = amc.find_period_mismatches(
        _week(tmp_path, "ШП", next_status), [_treatment(_october(1), _october(2))], files_by_kind=_WITH_JOURNAL,
    )

    assert results == []


def test_open_treatment_without_later_events_expects_treatment_to_the_end(tmp_path):
    results = amc.find_period_mismatches(
        _week(tmp_path, "ШП", "ШП"), [_treatment(_october(1), None)], files_by_kind=_WITH_JOURNAL,
    )

    wrong_days = [day.day for day, (actual, expected) in results[0]["dates"].items()
                  if not amc._statuses_match(expected, actual)]
    assert wrong_days == [3, 4, 5, 6, 7]


def test_open_treatment_ends_before_a_later_event_of_any_book(tmp_path):
    """Виписку в журналі часто не заповнюють: пізніша відпустка в book61 каже,
    що лікування вже скінчилось."""
    records = [_treatment(_october(1), None), _period(_PERSON, "ВП", _october(3), _october(7))]

    assert amc.find_period_mismatches(
        _week(tmp_path, "ШП", "ШП", "ВП", "ВП", "ВП", "ВП", "ВП"), records, files_by_kind=_WITH_JOURNAL,
    ) == []


def test_open_treatment_ends_with_the_discharge_known_to_another_book(tmp_path):
    records = [_treatment(_october(1), None), _period(_PERSON, "ШП", _october(1), _october(2), "ШП")]

    assert amc.find_period_mismatches(_week(tmp_path, "ШП", "ШП"), records, files_by_kind=_WITH_JOURNAL) == []


def test_open_treatment_confirms_the_next_hospital_of_the_same_chain(tmp_path):
    """Реальний випадок: людину перевели до іншої лікарні, у book61 новий
    рядок, а в журналі той самий без виписки. Це лікування, а не "немає
    запису"."""
    records = [
        _treatment(_october(1), None),
        _period(_PERSON, "ШП", _october(1), _october(3), "ШП"),
        _period(_PERSON, "ШП", _october(4), None, "ШП"),
    ]

    assert amc.find_period_mismatches(_week(tmp_path, *["ШП"] * 7), records, files_by_kind=_WITH_JOURNAL) == []


def test_treatment_with_a_year_typo_covers_nothing(tmp_path):
    """Реальний випадок: виписка 20.10.2025 при госпіталізації 15.10.2026. Дні
    показуються як "немає запису", а саму дату друкує список підозрілих."""
    record = {**_treatment(_october(1), date(2025, 10, 5)), "end_before_start": True}

    results = amc.find_period_mismatches(_week(tmp_path, "ШП"), [record], files_by_kind=_WITH_JOURNAL)

    assert [r["dates"] for r in results] == [{_october(1): ("ШП", _NO_TREATMENT_RECORD)}]


def test_journal_does_not_hide_a_mismatch_of_a_more_precise_book(tmp_path):
    """book61 каже ВЛК, у Табелі ШП: "будь-який статус лікування" журналу не
    має перекрити цю розбіжність."""
    records = [_period(_PERSON, "ВЛК", _october(1), _october(3), "ВЛК"), _treatment(_october(1), _october(3))]

    results = amc.find_period_mismatches(_week(tmp_path, "ШП", "ШП", "ШП"), records, files_by_kind=_WITH_JOURNAL)

    assert [(r["source_file"], r["dates"]) for r in results] == [("book61", {
        _october(day): ("ШП", "ВЛК") for day in (1, 2, 3)
    })]


def test_journal_name_spelled_differently_is_shown(tmp_path):
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", "П'ЯТИЙ П'ятий П'ятий", "ШП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    records = [_treatment(_october(1), _october(1), pib_raw="ПЯТИЙ Пятий Пятий")]

    results = amc.find_period_mismatches(Timesheet(oblik, "Табель"), records, files_by_kind=_WITH_JOURNAL)

    assert [(r["source_file"], r["sources"], r["addon_pib_raw"]) for r in results] == [
        (_JOURNAL, ["журнал лікування"], "ПЯТИЙ Пятий Пятий"),
    ]


_FOURTH = "ЧЕТВЕРТИЙ Четвертий Четвертий"
_FOURTH_TWO_TYPOS = "ЧЕТВРТИЙ Четвертийй Четвертий"
_FOURTH_ONE_TYPO = "ЧЕТВЕРТИЙ Четвертийй Четвертий"


def _fourth_week(tmp_path, *days):
    cells = [*days, *["РВЗ"] * (7 - len(days))]
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [["Посада1", "матрос", _FOURTH, *cells]], header=_WEEK_HEADER)
    return Timesheet(oblik, "Табель")


def test_two_misspelled_words_are_still_the_same_person(tmp_path):
    """Реальний випадок: у журналі лікування описки і в прізвищі, і в імені,
    однакове лише по батькові. Дні лікування за цим рядком - не "немає запису",
    а різницю написання показує таблиця ПІБ."""
    records = [_treatment(_october(1), _october(3), pib_raw=_FOURTH_TWO_TYPOS)]

    results = amc.find_period_mismatches(_fourth_week(tmp_path, "ШП", "ШП", "ШП"), records, files_by_kind=_WITH_JOURNAL)

    assert [r["addon_pib_raw"] for r in results] == [_FOURTH_TWO_TYPOS]
    assert all(amc._statuses_match(expected, actual) for actual, expected in results[0]["dates"].values())


def test_every_misspelled_row_of_a_person_is_kept(tmp_path):
    """Реальний випадок: два рядки журналу, обидва з різними помилками."""
    records = [
        _treatment(_october(4), None, pib_raw=_FOURTH_ONE_TYPO),
        _treatment(_october(1), _october(3), pib_raw=_FOURTH_TWO_TYPOS),
    ]

    results = amc.find_period_mismatches(_fourth_week(tmp_path, *["ШП"] * 7), records, files_by_kind=_WITH_JOURNAL)

    assert [[r["pib_raw"] for r in result["misspelled"]] for result in results] == [[_FOURTH_TWO_TYPOS, _FOURTH_ONE_TYPO]]


@pytest.mark.parametrize("name, same_person", [
    (_FOURTH_TWO_TYPOS, True),
    ("ЧЕТВРТИЙ Четвертийй Четвертийй", False),  # жодне слово не однакове
    ("ЧЕТВРТИЙ Сьомий Четвертий", False),  # друге слово не схоже
    ("Сьомий Сьомий Четвертий", False),
])
def test_how_far_a_misspelled_name_may_go(name, same_person):
    words = [(_FOURTH.upper(), _FOURTH.upper().split())]

    assert (amc._misspelled_roster_name(name.upper(), words) is not None) == same_person


def test_a_name_close_to_two_people_is_not_guessed():
    upper = [_FOURTH.upper(), "ЧЕТВЕРТИЙ ЧЕТВЕРТИЙЙЙ ЧЕТВЕРТИЙ"]
    words = [(name, name.split()) for name in upper]

    assert amc._misspelled_roster_name(_FOURTH_TWO_TYPOS.upper(), words) is None


def test_misspelled_journal_row_beside_a_correct_one_is_the_same_person(tmp_path):
    """У журналі людина буває кілька разів, тож правильний рядок не доводить,
    що схожий ПІБ - інша людина (інакше дні цього рядка були б "немає запису")."""
    records = [
        _treatment(_october(1), _october(2), pib_raw=_FOURTH),
        _treatment(_october(3), _october(4), pib_raw=_FOURTH_ONE_TYPO),
    ]

    results = amc.find_period_mismatches(
        _fourth_week(tmp_path, "ШП", "ШП", "ШП", "ШП"), records, files_by_kind=_WITH_JOURNAL,
    )

    assert [[r["pib_raw"] for r in result["misspelled"]] for result in results] == [[_FOURTH_ONE_TYPO]]
    assert all(amc._statuses_match(expected, actual) for actual, expected in results[0]["dates"].values())


def test_a_letter_written_with_two_symbols_is_no_name_mismatch(tmp_path):
    """Реальний випадок: у журналі "й" записана двома символами. Це той самий
    ПІБ: рядок звіряється, а в таблиці ПІБ його немає."""
    two_symbols = unicodedata.normalize("NFD", _FOURTH)
    assert two_symbols != _FOURTH
    records = [_treatment(_october(1), _october(2), pib_raw=two_symbols)]

    assert amc.find_period_mismatches(_fourth_week(tmp_path, "ШП", "ШП"), records, files_by_kind=_WITH_JOURNAL) == []


def _report_sheet(tmp_path, results, sheet=_JOURNAL):
    report = str(tmp_path / "mis_match.xlsx")
    amc.write_addon_mismatch_report(results, [], report, source_files=[sheet])
    return load_workbook(report)[sheet]


def test_names_table_shows_every_wrong_spelling_with_its_period(tmp_path):
    """Рішення користувача: у звіті видно кожне неправильне ПІБ файлу, щоб
    виправити його там: де воно, за який період і як правильно."""
    records = [
        _treatment(_october(1), _october(3), pib_raw=_FOURTH_TWO_TYPOS),
        _treatment(_october(4), None, pib_raw=_FOURTH_ONE_TYPO),
    ]
    results = amc.find_period_mismatches(_fourth_week(tmp_path, *["ШП"] * 7), records, files_by_kind=_WITH_JOURNAL)

    ws = _report_sheet(tmp_path, results)

    assert _names_table(ws) == [
        ["журнал лікування: 2026", "—", _FOURTH, _FOURTH_TWO_TYPOS, "01.10.2026 - 03.10.2026"],
        ["журнал лікування: 2026", "—", _FOURTH, _FOURTH_ONE_TYPO, "04.10.2026 - не вказано"],
    ]
    assert ws.cell(row=3, column=3).fill.fgColor.rgb == "FFC6EFCE"
    assert ws.cell(row=3, column=4).fill.fgColor.rgb == "FFFFC7CE"
    assert ws.column_dimensions["D"].width >= pib_column_width([_FOURTH_TWO_TYPOS])


def test_person_with_only_a_name_difference_is_not_in_the_dates_table(tmp_path):
    results = amc.find_period_mismatches(
        _fourth_week(tmp_path, "ШП"), [_treatment(_october(1), _october(1), pib_raw=_FOURTH_ONE_TYPO)],
        files_by_kind=_WITH_JOURNAL,
    )

    ws = _report_sheet(tmp_path, results)

    assert ws.cell(row=1, column=1).value == "Невідповідності ПІБ"
    assert amc.only_names_differ(results[0])
    assert ws.column_dimensions["E"].width >= len("01.10.2026 - 01.10.2026")


def test_period_mismatch_stays_in_the_dates_table_above_the_names(tmp_path):
    """Людина з розбіжною датою лишається в таблиці за датами, а ПІБ її запису -
    у таблиці ПІБ під нею."""
    results = amc.find_period_mismatches(
        _fourth_week(tmp_path, "ШП", "ШП"), [_treatment(_october(1), _october(1), pib_raw=_FOURTH_ONE_TYPO)],
        files_by_kind=_WITH_JOURNAL,
    )

    ws = _report_sheet(tmp_path, results)

    assert not amc.only_names_differ(results[0])
    assert ws.cell(row=1, column=1).value == "Джерело"
    assert ws.cell(row=2, column=1).value == "Табель (ОБЛІК.xlsx)"
    assert _names_table(ws) == [
        ["журнал лікування: 2026", "—", _FOURTH, _FOURTH_ONE_TYPO, "01.10.2026 - 01.10.2026"],
    ]


def test_no_names_table_when_every_name_is_right(tmp_path):
    results = amc.find_period_mismatches(_fourth_week(tmp_path, "ШП"), [], files_by_kind=_WITH_JOURNAL)

    ws = _report_sheet(tmp_path, results)

    assert [ws.cell(row=row, column=1).value for row in range(1, ws.max_row + 1)].count("Невідповідності ПІБ") == 0


def test_period_rows_of_a_vacation_without_destination_are_not_name_only():
    base = {"roster_pib_raw": _FOURTH, "addon_pib_raw": _FOURTH_ONE_TYPO, "misspelled": [{"pib_raw": _FOURTH_ONE_TYPO}],
            "dates": {_october(1): ("ВП", "ВП")}}

    assert amc.only_names_differ(base)
    assert not amc.only_names_differ({**base, "no_destination": [_october(1)]})
    assert not amc.only_names_differ({**base, "dates": {_october(1): ("РВЗ", "ВП")}})
    assert not amc.only_names_differ({**base, "misspelled": []})


def test_without_the_journal_treatment_statuses_are_not_looked_for_in_it(tmp_path):
    assert amc.find_period_mismatches(_week(tmp_path, "ЗВФН"), [], files_by_kind={"book61": "book61"}) == []


def test_mismatch_report_gives_the_journal_its_own_sheet(tmp_path):
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [(_PERSON, _october(1), _october(2))]})
    oblik = _week(tmp_path, "ШП", "ШП", "ШП").file_path
    report = str(tmp_path / "mis_match.xlsx")

    check_addon_mismatches_file(oblik, [journal], report_file_name=report)

    ws = load_workbook(report)["журнал"]
    day_3 = next(c for c in range(4, ws.max_column + 1) if ws.cell(row=1, column=c).value == datetime(2026, 10, 3))
    assert ws.cell(row=3, column=1).value == "журнал лікування"
    assert (ws.cell(row=2, column=day_3).value, ws.cell(row=3, column=day_3).value) == ("ШП", _NO_TREATMENT_RECORD)


def _training_journal(path):
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Адаптація"
    ws.append(["ПІБ", "З", "До", "НЦ"])
    ws.append([_PERSON, datetime(2026, 10, 1), datetime(2026, 10, 2), "198"])
    wb.save(str(path))
    return str(path)


def _one_book_of_each_kind(tmp_path):
    return {
        "book61": write_full_book61(tmp_path / "book61.xlsm", {
            "ППД": [[1, "Підрозділ1", "матрос", _PERSON, "ППД", date(2026, 10, 1), None]],
        }),
        "cvs": write_cvs(tmp_path / "cvs.xlsx", {
            MEDICAL_SHEET: [[None, 1, "Підрозділ1", "матрос", _PERSON, "соматика", datetime(2026, 10, 1),
                             "Лікарня", "№1", datetime(2026, 10, 2)]],
        }),
        "journal": _training_journal(tmp_path / "журнал БП.xlsx"),
        ADDON_TREATMENT_JOURNAL_KIND: write_treatment_journal(tmp_path / "журнал.xlsx", {
            "2026": [(_PERSON, datetime(2026, 10, 1), datetime(2026, 10, 2))],
        }),
    }


def _run_both_checks(tmp_path, books):
    oblik = _week(tmp_path, "ШП", "ШП").file_path
    check = check_addon_mismatches_file(oblik, books, report_file_name=str(tmp_path / "mis_match.xlsx"))
    cross = cross_check_addons(books, roster_path=oblik, report_file_name=str(tmp_path / "addons_mis_match.xlsx"))
    return check, cross


def test_book_kinds_of_the_test_set_are_all_the_known_ones(tmp_path):
    assert set(_one_book_of_each_kind(tmp_path)) == set(ADDON_BOOK_KINDS)


@pytest.mark.parametrize("missing_kind", list(ADDON_BOOK_KINDS))
def test_checks_go_on_without_any_one_book(tmp_path, missing_kind):
    """Рішення користувача: якщо якогось файлу немає, звіряються решта."""
    books = _one_book_of_each_kind(tmp_path)
    present = [path for kind, path in books.items() if kind != missing_kind]

    check, cross = _run_both_checks(tmp_path, present)

    assert check.checked_files == present and cross.checked_files == present
    assert check.unrecognized_files == [] and cross.unrecognized_files == []
    assert check.missing_books == (ADDON_BOOK_KINDS[missing_kind],)
    assert cross.missing_books == (ADDON_BOOK_KINDS[missing_kind],)


@pytest.mark.parametrize("only_kind", list(ADDON_BOOK_KINDS))
def test_checks_go_on_with_a_single_book(tmp_path, only_kind):
    book = _one_book_of_each_kind(tmp_path)[only_kind]

    check, cross = _run_both_checks(tmp_path, [book])

    assert check.checked_files == [book] and cross.checked_files == [book]
    assert ADDON_BOOK_KINDS[only_kind] not in check.missing_books
    assert len(check.missing_books) == len(ADDON_BOOK_KINDS) - 1


def test_with_every_book_nothing_is_named_as_missing(tmp_path):
    check, cross = _run_both_checks(tmp_path, list(_one_book_of_each_kind(tmp_path).values()))

    assert check.missing_books == () and cross.missing_books == ()


# --- у звіті статуси, а не види оплати -----------------------------------------


def _day_cells(ws, day, first_row=2):
    """(рядок Табеля, рядок книги) колонки дати day у записі, що починається з first_row."""
    col = next(c for c in range(4, ws.max_column + 1) if ws.cell(row=1, column=c).value == datetime(2026, 10, day))
    return ws.cell(row=first_row, column=col), ws.cell(row=first_row + 1, column=col)


def _book61_sheet(tmp_path, results):
    report = str(tmp_path / "mis_match.xlsx")
    amc.write_addon_mismatch_report(results, [], report, source_files=["book61"])
    return load_workbook(report)["book61"]


@pytest.mark.parametrize("category", [30, 70, 100, 170])
def test_report_shows_rvz_instead_of_a_duty_payment(tmp_path, category):
    """Рішення користувача: 30/70/100/170 - вид оплати людини на службі, а
    звіт читають за статусами."""
    results = amc.find_period_mismatches(_week(tmp_path, category), [_period(_PERSON, "ШП", _october(1), _october(1), "ШП")])

    table_cell, book_cell = _day_cells(_book61_sheet(tmp_path, results), 1)

    assert (table_cell.value, book_cell.value) == ("РВЗ", "ШП")


@pytest.mark.parametrize("value, behind", [
    (10, "ППД"), (10, "ВД"), (10, "Адап"), (10, "НОВ"), ("10_МЕД", "ШП"), ("10_МЕД", "ЛМР"), ("10_МЕД", "ВЛК"),
    ("100_БПШП", "БПШП"), ("100_ВПБП", "ВПБП"),
])
def test_report_shows_the_status_instead_of_a_payment_value(tmp_path, value, behind):
    timesheet = _week(tmp_path, value)
    timesheet.statuses_behind_numbers = {(2, _october(1)): behind} if value == 10 else {}
    timesheet.statuses_behind_categories = {} if value == 10 else {(2, _october(1)): behind}
    records = [_period(_PERSON, "СЗЧ", _october(1), _october(1), "СЗЧ")]

    table_cell, book_cell = _day_cells(_book61_sheet(tmp_path, amc.find_period_mismatches(timesheet, records)), 1)

    assert (table_cell.value, book_cell.value) == (behind, "СЗЧ")


def test_green_day_shows_the_status_too(tmp_path):
    timesheet = _week(tmp_path, 10, "РВЗ")
    timesheet.statuses_behind_numbers = {(2, _october(1)): "ППД"}
    records = [_period(_PERSON, "ППД", _october(1), _october(1), "ППД"), _period(_PERSON, "СЗЧ", _october(2), _october(2), "СЗЧ")]

    table_cell, _book_cell = _day_cells(_book61_sheet(tmp_path, amc.find_period_mismatches(timesheet, records)), 1)

    assert table_cell.value == "ППД"
    assert table_cell.fill.fgColor.rgb == "FFC6EFCE"


def test_grey_day_shows_the_status_too(tmp_path):
    """День, який цей файл не звіряв, а колонка є через іншу дату."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", _PERSON, 100, 10, "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "РВЗ", 100, "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    timesheet = Timesheet(oblik, "Табель")
    timesheet.statuses_behind_numbers = {(2, _october(2)): "ВД"}
    records = [
        _period(_PERSON, "ШП", _october(1), _october(1), "ШП"),
        _period("ДРУГИЙ Другий Другий", "ШП", _october(2), _october(2), "ШП"),
    ]

    ws = _book61_sheet(tmp_path, amc.find_period_mismatches(timesheet, records))
    first_row = next(row for row in range(2, ws.max_row + 1) if ws.cell(row=row, column=3).value == _PERSON)

    table_cell, _book_cell = _day_cells(ws, 2, first_row)
    assert table_cell.value == "ВД"
    assert table_cell.fill.fgColor.rgb == "FFF2F2F2"


def test_comparison_still_uses_the_value_of_the_file(tmp_path):
    """Статус лише для показу: 10 і ВД у книзі - збіг, як і раніше."""
    timesheet = _week(tmp_path, 10)
    timesheet.statuses_behind_numbers = {(2, _october(1)): "ППД"}

    results = amc.find_period_mismatches(timesheet, [_period(_PERSON, "ВД", _october(1), _october(1), "ВД")])

    assert results == []


def test_report_keeps_a_value_with_an_unknown_status(tmp_path):
    """Без рапортів статус за 10 невідомий, а вгадувати його не можна."""
    results = amc.find_period_mismatches(_week(tmp_path, 10), [_period(_PERSON, "СЗЧ", _october(1), _october(1), "СЗЧ")])

    table_cell, _book_cell = _day_cells(_book61_sheet(tmp_path, results), 1)

    assert table_cell.value == 10


def test_report_names_the_statuses_when_a_book_expects_a_payment_category(tmp_path):
    """Журнал обліку БП: "закордон" - категорія 10, а не статус."""
    results = amc.find_period_mismatches(_week(tmp_path, "СЗЧ"), [_period(_PERSON, 10, _october(1), _october(1), "Фахова підготовка", "book61")])

    _table_cell, book_cell = _day_cells(_book61_sheet(tmp_path, results), 1)

    assert set(str(book_cell.value).split("/")) >= {"ППД", "ВД", "БЗВП", "НОВ", "Адап"}


# --- ВЛК під час відпустки ---------------------------------------------------------


_LEAVES = ["ВП", "ВПС", "ВПСЗ", "ВПБП"]


def _vlk(book="cvs", day=1):
    return {**_period(_PERSON, "ВЛК", _october(day), _october(day), "ВЛК", book), "arrival_date": _october(day)}


@pytest.mark.parametrize("book", ["cvs", "book61"])
@pytest.mark.parametrize("leave", _LEAVES)
def test_leave_in_the_timesheet_has_priority_over_vlk_of_any_book(tmp_path, leave, book):
    """Рішення користувача: періоди йдуть паралельно, відпустка будь-якого типу
    важливіша. Реальний випадок: ВПСЗ у Табелі, ВЛК у ЦВС, направлення №6068."""
    assert amc.find_period_mismatches(_week(tmp_path, leave), [_vlk(book)]) == []


def test_leave_paid_as_a_combat_injury_is_the_same_leave(tmp_path):
    """ВПСЗ, яку підрозділ подає як ВПБП: у файлі для виплат це 100_ВПБП, а
    звірка стройової бачить той самий день як ВПСЗ."""
    timesheet = _week(tmp_path, "100_ВПБП")
    timesheet.statuses_behind_categories = {(2, _october(1)): "ВПСЗ"}

    assert amc.find_period_mismatches(timesheet, [_vlk()]) == []


@pytest.mark.parametrize("status", [*_LEAVES, "100_ВПБП", " впсз "])
def test_leave_is_recognized_in_any_form(status):
    """100_ВПБП - так ВПБП записаний у файлі для виплат."""
    assert amc._is_leave(status) is True


@pytest.mark.parametrize("status", ["ШП", "ВЛК", "РВЗ", "ППД", 100, None])
def test_other_statuses_are_not_a_leave(status):
    assert amc._is_leave(status) is False


@pytest.mark.parametrize("status", ["ШП", "РВЗ", "ППД"])
def test_vlk_gives_way_only_to_a_leave(tmp_path, status):
    results = amc.find_period_mismatches(_week(tmp_path, status), [{**_vlk(), "arrival_date": None}])

    assert [r["dates"][_october(1)] for r in results] == [(status, "ВЛК")]


def test_number_against_vlk_stays_a_mismatch(tmp_path):
    """У клітинці 100, за якою немає статусу: це не відпустка."""
    results = amc.find_period_mismatches(_week(tmp_path, 100), [{**_vlk(), "arrival_date": None}])

    assert [r["dates"][_october(1)] for r in results] == [(100, "ВЛК")]


def test_only_vlk_gives_way_to_a_leave(tmp_path):
    record = {**_period(_PERSON, "ШП", _october(1), _october(1), "ШП", "book61"), "arrival_date": None}

    results = amc.find_period_mismatches(_week(tmp_path, "ВПСЗ"), [record])

    assert [r["dates"][_october(1)] for r in results] == [("ВПСЗ", "ШП")]


@pytest.mark.parametrize("leave", _LEAVES)
def test_leave_in_the_timesheet_has_priority_over_the_treatment_journal(tmp_path, leave):
    """Статусу в журналі немає: будь-який його запис."""
    timesheet = _week(tmp_path, leave, leave, leave)

    assert amc.find_period_mismatches(timesheet, [_treatment(_october(1), _october(3))], files_by_kind=_WITH_JOURNAL) == []


def test_leave_after_an_open_journal_record_has_priority_too(tmp_path):
    """Запис без виписки тягнеться до кінця місяця, а людина вже у відпустці."""
    timesheet = _week(tmp_path, "ШП", *["ВПСЗ"] * 6)

    assert amc.find_period_mismatches(timesheet, [_treatment(_october(1), None)], files_by_kind=_WITH_JOURNAL) == []


def test_journal_gives_way_only_to_a_leave(tmp_path):
    results = amc.find_period_mismatches(
        _week(tmp_path, "ШП", "ППД", "ШП"), [_treatment(_october(1), _october(3))], files_by_kind=_WITH_JOURNAL,
    )

    assert [r["dates"][_october(2)] for r in results] == [("ППД", ADDON_TREATMENT_JOURNAL_STATUS)]


def test_books_and_timesheet_end_to_end(tmp_path):
    """Від книг до результату: ЦВС зі справжнім направленням і журнал лікування."""
    oblik = _write_roster(tmp_path / "ОБЛІК.xlsx", [
        ["Посада1", "матрос", _PERSON, "ВПСЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
        ["Посада2", "матрос", "ДРУГИЙ Другий Другий", "ВП", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ", "РВЗ"],
    ], header=_WEEK_HEADER)
    cvs = write_cvs(tmp_path / "cvs.xlsx", {VLK_SHEET: [
        [1, _PERSON, "матрос", datetime(2026, 10, 1), datetime(2026, 10, 1), "№1 від 01.10.2026", None, "№6068"],
    ]})
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [
        ("ДРУГИЙ Другий Другий", datetime(2026, 10, 1), datetime(2026, 10, 1)),
    ]})

    result = check_addon_mismatches_file(oblik, [cvs, journal], report_file_name=str(tmp_path / "mis_match.xlsx"))

    assert result.period_mismatches == []


# --- ДТЛ - довготривале лікування ------------------------------------------------


@pytest.mark.parametrize("book_status", ["ВПБП", "ВПСЗ", "ШП", "БПШП", "ВЛК", "ЛМР"])
def test_long_treatment_in_the_timesheet_fits_every_treatment_status(tmp_path, book_status):
    """Рішення користувача: ДТЛ рівнозначний цим статусам."""
    records = [_period(_PERSON, book_status, _october(1), _october(3), "ШП")]

    assert amc.find_period_mismatches(_week(tmp_path, "ДТЛ", "ДТЛ", "ДТЛ"), records) == []


@pytest.mark.parametrize("book_status", [ADDON_NON_COMBAT_TREATMENT_STATUS, ADDON_TREATMENT_JOURNAL_STATUS])
def test_long_treatment_fits_books_that_allow_several_statuses(book_status):
    assert amc._statuses_match(book_status, "ДТЛ") is True


@pytest.mark.parametrize("book_status", ["ВП", "СЗЧ", "ВД", "ППД", "РВЗ"])
def test_long_treatment_does_not_fit_anything_else(book_status):
    assert amc._statuses_match(book_status, "ДТЛ") is False


def test_long_treatment_does_not_make_treatment_statuses_equal():
    """ДТЛ не об'єднує ці статуси між собою: ШП і ВЛК лишаються різними."""
    assert amc._statuses_match("ШП", "ВЛК") is False
    assert amc._statuses_match("ВЛК", "БПШП") is False


def test_long_treatment_needs_no_record_in_any_book(tmp_path):
    results = amc.find_period_mismatches(
        _week(tmp_path, "ДТЛ", "ДТЛ"), [],
        files_by_kind={"book61": "book61", "journal": "journal", ADDON_TREATMENT_JOURNAL_KIND: _JOURNAL},
    )

    assert results == []
