from datetime import date, datetime

import pytest
from book61_fixtures import write_full_book61
from cvs_fixtures import (
    CAPTIVITY_SHEET, KILLED_SHEET, MEDICAL_SHEET, MISSING_SHEET, VLK_SHEET, ZVFN_SHEET, write_cvs,
)
from openpyxl import Workbook

import content.cvs_reader as cr
from content.addon_reader import load_addon_workbook


@pytest.fixture(autouse=True)
def _own_unit(monkeypatch):
    monkeypatch.setattr(cr, "OWN_MILITARY_UNIT_CODES", frozenset({"0000"}))


def _records(tmp_path, rows_by_sheet):
    wb = load_addon_workbook(write_cvs(tmp_path / "cvs.xlsx", rows_by_sheet))
    return [
        (r["pib_raw"], r["status"], r["start_date"], r["end_date"], r["source_sheet"], r["arrival_date"])
        for r in cr.read_all_period_records(wb, "cvs")
    ]


def test_vlk_runs_from_the_passing_date_to_the_document_date(tmp_path):
    """Реальний випадок, підтверджений користувачем: "Дата проходження ВЛК"
    14.10, документ "від 20.10.2026" - ВЛК закінчилась 20.10. Лише одна з
    дат (на реальних даних колонку майже ніколи не заповнюють - дата в тексті
    документа після "від", буває й одруківка "віл") чи документ, старіший за
    початок, - один день."""
    records = _records(tmp_path, {VLK_SHEET: [
        [1, "ПЕРШИЙ Перший Перший", "матрос", datetime(2026, 9, 1), datetime(2026, 10, 3), "№1 від 05.10.2026", "придатний"],
        [2, "ДРУГИЙ Другий Другий", "матрос", datetime(2026, 9, 1), None, "ВЛК: 2026-1004-1 від 04.10.2026", "придатний"],
        [3, "ТРЕТІЙ Третій Третій", "матрос", datetime(2026, 9, 1), None, "№2026-1230-1 віл 06.10.2026", "придатний"],
        [4, "ЧЕТВЕРТИЙ Четвертий Четвертий", "матрос", datetime(2026, 9, 1), None, "Направлення: 155", None],
        [5, "П'ЯТИЙ П'ятий П'ятий", "матрос", datetime(2026, 9, 1), datetime(2026, 10, 7), "№2 від 01.09.2026", None],
        [6, "ШОСТИЙ Шостий Шостий", "матрос", datetime(2026, 9, 1), datetime(2026, 10, 8), None, None],
    ]})
    assert records == [
        ("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 10, 3), date(2026, 10, 5), "ВЛК", date(2026, 10, 5)),
        ("ДРУГИЙ Другий Другий", "ВЛК", date(2026, 10, 4), date(2026, 10, 4), "ВЛК", date(2026, 10, 4)),
        ("ТРЕТІЙ Третій Третій", "ВЛК", date(2026, 10, 6), date(2026, 10, 6), "ВЛК", date(2026, 10, 6)),
        ("П'ЯТИЙ П'ятий П'ятий", "ВЛК", date(2026, 10, 7), date(2026, 10, 7), "ВЛК", date(2026, 10, 7)),
        ("ШОСТИЙ Шостий Шостий", "ВЛК", date(2026, 10, 8), date(2026, 10, 8), "ВЛК", date(2026, 10, 8)),
    ]


def test_zvfn_is_a_period_between_start_and_end(tmp_path):
    records = _records(tmp_path, {ZVFN_SHEET: [
        [1, "ПЕРШИЙ Перший Перший", datetime(2026, 10, 2), datetime(2026, 10, 9), "№5", None],
        [2, "ДРУГИЙ Другий Другий", None, datetime(2026, 10, 9), "№6", None],
    ]})
    assert records == [("ПЕРШИЙ Перший Перший", "ЗВФН", date(2026, 10, 2), date(2026, 10, 9), ZVFN_SHEET.strip(), None)]


def test_medical_status_depends_on_the_facility_and_the_kind(tmp_path):
    """map: заклад власної частини - ЛМР; інакше поранення (чи бойове
    травмування) - БПШП, соматика/травмування - ШП; інших типів map не
    називає - рядок не звіряється. Дата виписки - день прибуття."""
    def row(name, kind, facility, start, end):
        return [None, 1, "Підрозділ1", "матрос", name, kind, start, facility, "№1", end]

    records = _records(tmp_path, {MEDICAL_SHEET: [
        row("ПЕРШИЙ Перший Перший", "соматика ", "військова частина А0000", datetime(2026, 10, 1), datetime(2026, 10, 4)),
        row("ДРУГИЙ Другий Другий", "поранення", "Лікарня", datetime(2026, 10, 2), datetime(2026, 10, 5)),
        row("ТРЕТІЙ Третій Третій", "травмування бойове", "Лікарня", datetime(2026, 10, 2), None),
        row("ЧЕТВЕРТИЙ Четвертий Четвертий", "травмування", "військова частина A0001", datetime(2026, 10, 3), "виписка 06.10.2026"),
        row("П'ЯТИЙ П'ятий П'ятий", "соматика", "Лікарня", "надійшов 03.10.2026", datetime(2026, 10, 7)),
        row("ШОСТИЙ Шостий Шостий", "після полону", "Лікарня", datetime(2026, 10, 1), datetime(2026, 10, 7)),
    ]})
    assert records == [
        ("ПЕРШИЙ Перший Перший", "ЛМР", date(2026, 10, 1), date(2026, 10, 4), MEDICAL_SHEET, date(2026, 10, 4)),
        ("ДРУГИЙ Другий Другий", "БПШП", date(2026, 10, 2), date(2026, 10, 5), MEDICAL_SHEET, date(2026, 10, 5)),
        ("ТРЕТІЙ Третій Третій", "БПШП", date(2026, 10, 2), None, MEDICAL_SHEET, None),
        ("ЧЕТВЕРТИЙ Четвертий Четвертий", "ШП", date(2026, 10, 3), date(2026, 10, 6), MEDICAL_SHEET, date(2026, 10, 6)),
        ("П'ЯТИЙ П'ятий П'ятий", "ШП", date(2026, 10, 3), date(2026, 10, 7), MEDICAL_SHEET, date(2026, 10, 7)),
    ]


def test_missing_starts_on_the_disappearance_date_and_skips_szch(tmp_path):
    records = _records(tmp_path, {MISSING_SHEET: [
        [1, "Стрілець", "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", datetime(2024, 4, 29), "безвісті", datetime(2026, 6, 14)],
        [2, "Стрілець", "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", datetime(2024, 4, 29), "СЗЧ", datetime(2026, 6, 14)],
    ]})
    assert records == [("ПЕРШИЙ Перший Перший", "БЗ", date(2026, 6, 14), None, MISSING_SHEET, None)]


def test_captivity_ends_the_day_before_release(tmp_path):
    records = _records(tmp_path, {CAPTIVITY_SHEET: [
        [1, "солдат", "ПЕРШИЙ Перший Перший", datetime(1990, 1, 1), datetime(2024, 8, 13), datetime(2026, 10, 5)],
        [2, "солдат", "ДРУГИЙ Другий Другий", datetime(1990, 1, 1), datetime(2024, 8, 4), None],
    ]})
    assert records == [
        ("ПЕРШИЙ Перший Перший", "полон", date(2024, 8, 13), date(2026, 10, 4), CAPTIVITY_SHEET, None),
        ("ДРУГИЙ Другий Другий", "полон", date(2024, 8, 4), None, CAPTIVITY_SHEET, None),
    ]


def test_killed_starts_from_the_death_certificate_date_and_skips_the_column_number_row(tmp_path):
    """За прямою вказівкою користувача "Загиб" - лише з лікарського
    свідоцтва про смерть (реальний випадок: дата втрат раніша, свідоцтво -
    від 08.10.2026, і саме з нього має бути "Загиб"); "Дата втрат (смерті)"
    - лише коли дати свідоцтва немає. Період відкритий."""
    records = _records(tmp_path, {KILLED_SHEET: [
        [None, 1, "матрос", "ПЕРШИЙ Перший Перший", datetime(2026, 10, 2), "№1 від 08.10.2026"],
        [None, 2, "матрос", "ДРУГИЙ Другий Другий", "03.10.2026 (уточнено)", None],
        [None, 3, "матрос", "ТРЕТІЙ Третій Третій", datetime(2026, 10, 1), datetime(2026, 10, 5)],
        [None, 4, "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", datetime(2026, 10, 1), "№2, 06.10.2026"],
        [None, 5, "матрос", "П'ЯТИЙ П'ятий П'ятий", datetime(2026, 10, 4), "№3"],
    ]})
    assert records == [
        ("ПЕРШИЙ Перший Перший", "Загиб", date(2026, 10, 8), None, KILLED_SHEET, None),
        ("ДРУГИЙ Другий Другий", "Загиб", date(2026, 10, 3), None, KILLED_SHEET, None),
        ("ТРЕТІЙ Третій Третій", "Загиб", date(2026, 10, 5), None, KILLED_SHEET, None),
        ("ЧЕТВЕРТИЙ Четвертий Четвертий", "Загиб", date(2026, 10, 6), None, KILLED_SHEET, None),
        ("П'ЯТИЙ П'ятий П'ятий", "Загиб", date(2026, 10, 4), None, KILLED_SHEET, None),
    ]


def test_every_record_carries_the_source_file(tmp_path):
    wb = load_addon_workbook(write_cvs(tmp_path / "cvs.xlsx", {KILLED_SHEET: [
        [None, 1, "матрос", "ПЕРШИЙ Перший Перший", datetime(2026, 10, 2)],
    ]}))
    assert [r["source_file"] for r in cr.read_all_period_records(wb, "Книга ЦВС")] == ["Книга ЦВС"]


def test_headers_are_found_regardless_of_case_spaces_and_apostrophe(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["№", "  прізвище  ІМ’Я по батькові ", "Соматика / травмування / поранення", "дата надходження в мед заклад",
               "лікувальний заклад", "ДАТА ВИПИСКИ"])
    ws.append([1, "ПЕРШИЙ Перший Перший", "соматика", datetime(2026, 10, 1), "Лікарня", datetime(2026, 10, 2)])
    assert cr.is_cvs_workbook(wb)
    assert [r["status"] for r in cr.read_all_period_records(wb, "cvs")] == ["ШП"]


def test_is_cvs_workbook_rejects_book61_and_unrelated_books(tmp_path):
    assert not cr.is_cvs_workbook(load_addon_workbook(write_full_book61(tmp_path / "book61.xlsm", {})))
    wb = Workbook()
    wb.active.append(["ПІБ", "Щось інше"])
    assert not cr.is_cvs_workbook(wb)
    assert cr.is_cvs_workbook(load_addon_workbook(write_cvs(tmp_path / "cvs.xlsx", {VLK_SHEET: []})))
