from datetime import date, datetime

from book61_fixtures import write_full_book61
from cvs_fixtures import VLK_SHEET, write_cvs
from openpyxl import Workbook
from treatment_fixtures import write_treatment_journal

import content.treatment_journal_reader as tjr
from constants import ADDON_TREATMENT_JOURNAL_KIND, ADDON_TREATMENT_JOURNAL_STATUS
from content.addon_reader import implausible_date_records, load_addon_workbook
from generators.check_addon_mismatches import _read_addon

_PERSON = "ПЕРШИЙ Перший Перший"


def _read(path):
    return tjr.read_all_period_records(load_addon_workbook(path), "журнал")


def test_journal_is_recognized_by_its_columns_not_by_the_file_name(tmp_path):
    journal = write_treatment_journal(tmp_path / "будь-яка назва.xlsx", {"2026": []})
    wrong_columns = Workbook()
    wrong_columns.worksheets[0].append(["Прізвище, ім’я та по батькові", "Дата госпіталізації"])
    wrong_columns_path = tmp_path / "wrong.xlsx"
    wrong_columns.save(str(wrong_columns_path))

    assert tjr.is_treatment_journal_workbook(load_addon_workbook(journal))
    assert not tjr.is_treatment_journal_workbook(load_addon_workbook(str(wrong_columns_path)))
    assert not tjr.is_treatment_journal_workbook(load_addon_workbook(write_cvs(tmp_path / "cvs.xlsx", {VLK_SHEET: []})))
    assert not tjr.is_treatment_journal_workbook(load_addon_workbook(write_full_book61(tmp_path / "book61.xlsm", {})))


def test_check_recognizes_the_journal_as_its_own_kind_of_book(tmp_path):
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [(_PERSON, datetime(2026, 10, 1), None)]})

    content = _read_addon(journal)

    assert content is not None
    records, unmapped, kind = content
    assert kind == ADDON_TREATMENT_JOURNAL_KIND
    assert unmapped == []
    assert [record["source_file"] for record in records] == ["журнал"]


def test_every_year_sheet_gives_a_period_with_any_of_the_treatment_statuses(tmp_path):
    """Статусу в журналі немає, лише період. Виписка - день прибуття, як і в
    інших книгах."""
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {
        "2025": [(_PERSON, datetime(2025, 12, 1), datetime(2025, 12, 5))],
        "2026": [("ДРУГИЙ Другий Другий", datetime(2026, 10, 1), datetime(2026, 10, 3))],
    })

    assert _read(journal) == [{
        "pib_raw": _PERSON, "status": ADDON_TREATMENT_JOURNAL_STATUS,
        "start_date": date(2025, 12, 1), "end_date": date(2025, 12, 5), "arrival_date": date(2025, 12, 5),
        "source_sheet": "2025", "source_file": "журнал", "end_before_start": False,
    }, {
        "pib_raw": "ДРУГИЙ Другий Другий", "status": ADDON_TREATMENT_JOURNAL_STATUS,
        "start_date": date(2026, 10, 1), "end_date": date(2026, 10, 3), "arrival_date": date(2026, 10, 3),
        "source_sheet": "2026", "source_file": "журнал", "end_before_start": False,
    }]


def test_treatment_without_a_discharge_date_stays_open(tmp_path):
    """Реальні випадки: клітинка порожня чи пробіл."""
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [
        (_PERSON, datetime(2026, 10, 1), None),
        ("ДРУГИЙ Другий Другий", datetime(2026, 10, 2), " "),
    ]})

    assert [(r["end_date"], r["arrival_date"]) for r in _read(journal)] == [(None, None), (None, None)]


def test_date_written_as_text_is_read(tmp_path):
    """Реальний випадок: виписка текстом із зайвим знаком у кінці."""
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2025": [(_PERSON, "01.02.2025", "18.02.2025\\")]})

    assert [(r["start_date"], r["end_date"]) for r in _read(journal)] == [(date(2025, 2, 1), date(2025, 2, 18))]


def test_rows_without_a_name_or_an_admission_date_are_skipped(tmp_path):
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [
        (None, datetime(2026, 10, 1), None),
        (_PERSON, None, datetime(2026, 10, 5)),
    ]})

    assert _read(journal) == []


def test_discharge_before_admission_is_not_corrected_but_shown_as_suspicious(tmp_path):
    """Реальний випадок: лікування 15.10.2026, виписка "20.10.2025" - рік 2025
    замість 2026. Мовчки не виправляється, виправити треба саму книгу."""
    journal = write_treatment_journal(tmp_path / "журнал.xlsx", {"2026": [
        (_PERSON, datetime(2026, 10, 15), datetime(2025, 10, 20)),
        ("ДРУГИЙ Другий Другий", datetime(2026, 10, 15), datetime(2026, 10, 15)),
    ]})

    records = _read(journal)

    assert [r["end_date"] for r in records] == [date(2025, 10, 20), date(2026, 10, 15)]
    assert [r["pib_raw"] for r in implausible_date_records(records, today=date(2026, 10, 31))] == [_PERSON]
