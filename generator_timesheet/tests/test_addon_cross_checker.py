"""Звірка книг resources/addons між собою (content/addon_cross_checker.py,
generators/check_addon_mismatches.py::cross_check_addons)."""

from datetime import date, datetime

from book61_fixtures import write_full_book61
from cvs_fixtures import MEDICAL_SHEET, write_cvs
from openpyxl import Workbook, load_workbook

import content.addon_cross_checker as acc
import content.addon_mismatch_checker as amc
from content.addon_reader import ANY_STATUS
from content.oblik_timesheet import Timesheet
from generators.check_addon_mismatches import cross_check_addons

_HEADER = ["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ"] + [datetime(2026, 10, day) for day in range(1, 8)]


def _roster(path, names):
    wb = Workbook()
    ws = wb.active
    ws.title = "Табель"
    ws.append(_HEADER)
    for index, name in enumerate(names, start=1):
        ws.append([f"Підрозділ{index}", f"Посада{index}", "сержант", name] + [None] * 7)
    wb.save(str(path))
    return str(path)


def _period(pib_raw, status, start, end, source_file, source_sheet, arrival_date=None):
    return {
        "pib_raw": pib_raw, "status": status, "start_date": start, "end_date": end,
        "source_sheet": source_sheet, "source_file": source_file, "arrival_date": arrival_date,
    }


def test_statuses_of_newcomers_are_identical():
    """За прямою вказівкою користувача "НОВ" і "Адап" - не розбіжність."""
    assert amc._statuses_match("Адап", "НОВ") is True
    assert amc._statuses_match("НОВ", "Адап.") is True


def test_find_addon_conflicts_compares_each_book_on_its_own_timeline(tmp_path):
    """Кожна книга - своя шкала; звіряються лише дати, де обидві книги щось
    очікують, тими самими правилами еквівалентності (ППД=НОВ, 10 = ВД,
    "будь-який"); день виписки однієї з книг не звіряється; людина не з
    Табеля - не звіряється."""
    timesheet = Timesheet(_roster(tmp_path / "ОБЛІК.xlsx", ["ПЕРШИЙ Перший Перший", "ДРУГИЙ Другий Другий"]), "Табель")
    records = [
        # ПЕРШИЙ: 01-02 збіг ППД/НОВ, 03-04 ВП проти ШП, 05 - день виписки ЦВС.
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 1), date(2026, 10, 2), "book61", "ППД"),
        _period("ПЕРШИЙ Перший Перший", "ВП", date(2026, 10, 3), date(2026, 10, 5), "book61", "ВП"),
        _period("ПЕРШИЙ Перший Перший", "НОВ", date(2026, 10, 1), date(2026, 10, 2), "cvs", "НОВ"),
        _period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 10, 3), date(2026, 10, 5), "cvs", MEDICAL_SHEET,
                arrival_date=date(2026, 10, 5)),
        # ДРУГИЙ: усе збігається.
        _period("ДРУГИЙ Другий Другий", "ВД", date(2026, 10, 1), date(2026, 10, 3), "book61", "ВД"),
        _period("ДРУГИЙ Другий Другий", 10, date(2026, 10, 1), date(2026, 10, 2), "journal", "Фахова підготовка"),
        _period("ДРУГИЙ Другий Другий", ANY_STATUS, date(2026, 10, 3), date(2026, 10, 3), "journal", "БЗВП"),
        # Не з Табеля.
        _period("ТРЕТІЙ Третій Третій", "ВП", date(2026, 10, 1), date(2026, 10, 2), "book61", "ВП"),
        _period("ТРЕТІЙ Третій Третій", "ШП", date(2026, 10, 1), date(2026, 10, 2), "cvs", MEDICAL_SHEET),
    ]

    conflicts = acc.find_addon_conflicts(timesheet, records, ["book61", "cvs", "journal"], last_date=date(2026, 10, 7))

    assert [(c["files"], c["roster_pib_raw"], c["roster_pidrozdil"], c["days"]) for c in conflicts] == [
        (("book61", "cvs"), "ПЕРШИЙ Перший Перший", "Підрозділ1", 2),
    ]
    assert conflicts[0]["periods"] == [{
        "start": date(2026, 10, 3), "end": date(2026, 10, 4), "status_a": "ВП", "status_b": "ШП",
        "sheet_a": "ВП", "sheet_b": MEDICAL_SHEET,
    }]


def test_find_addon_conflicts_checks_every_period_not_only_the_tabel_month(tmp_path):
    """За прямою вказівкою користувача - за ВСІ періоди: Табель за жовтень, а
    розбіжність книг у березні теж знаходиться; поспіль ідучі дні з тією
    самою парою статусів - один період."""
    timesheet = Timesheet(_roster(tmp_path / "ОБЛІК.xlsx", ["ПЕРШИЙ Перший Перший"]), "Табель")
    records = [
        _period("ПЕРШИЙ Перший Перший", "ППД", date(2026, 3, 1), date(2026, 3, 31), "book61", "ППД"),
        _period("ПЕРШИЙ Перший Перший", "ШП", date(2026, 3, 10), date(2026, 3, 12), "cvs", MEDICAL_SHEET),
        _period("ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 3, 20), date(2026, 3, 20), "cvs", "ВЛК"),
    ]

    (conflict,) = acc.find_addon_conflicts(timesheet, records, ["book61", "cvs"], last_date=date(2026, 10, 7))

    assert [(p["start"], p["end"], p["status_a"], p["status_b"]) for p in conflict["periods"]] == [
        (date(2026, 3, 10), date(2026, 3, 12), "ППД", "ШП"),
        (date(2026, 3, 20), date(2026, 3, 20), "ППД", "ВЛК"),
    ]
    assert conflict["days"] == 4


def test_write_addon_conflicts_report_lists_periods_in_one_row_per_person(tmp_path):
    conflicts = [{
        "files": ("book61", "cvs"), "roster_pib_raw": "ПЕРШИЙ Перший Перший", "roster_pidrozdil": None,
        "periods": [
            {"start": date(2026, 3, 10), "end": date(2026, 3, 12), "status_a": "ППД", "status_b": "ШП",
             "sheet_a": "ППД", "sheet_b": MEDICAL_SHEET},
            {"start": date(2026, 10, 1), "end": date(2026, 10, 1), "status_a": ANY_STATUS, "status_b": "ВЛК",
             "sheet_a": "Прибулі у ВД", "sheet_b": "ВЛК"},
        ],
        "days": 4,
    }]
    output_path = str(tmp_path / "addons_mis_match.xlsx")

    assert acc.write_addon_conflicts_report([], output_path) is None
    assert acc.write_addon_conflicts_report(conflicts, output_path) == output_path

    wb = load_workbook(output_path)
    assert wb.sheetnames == ["book61 - cvs"]
    ws = wb["book61 - cvs"]
    assert [ws.cell(row=1, column=c).value for c in (1, 2, 3, 4, 6)] == [
        "Підрозділ", "ПІБ", "book61: аркуші", "cvs: аркуші", "Днів",
    ]
    assert [ws.cell(row=2, column=c).value for c in (1, 2, 3, 4, 6)] == [
        "—", "ПЕРШИЙ Перший Перший", "ППД, Прибулі у ВД", f"{MEDICAL_SHEET}, ВЛК", 4,
    ]
    assert ws.cell(row=2, column=5).value == "10.03.2026-12.03.2026: ППД / ШП\n01.10.2026: будь-який / ВЛК"
    assert ws.max_row == 2


def test_pair_titles_keep_both_books_within_the_excel_limit():
    assert acc._pair_title("book61", "cvs") == "book61 - cvs"
    title = acc._pair_title("Журнал обліку БП 2026", "Книга ЦВС підрозділу")
    assert len(title) <= 31 and title.startswith("Журнал") and " - Книга" in title


def test_cross_check_addons_reads_real_books_and_reports_conflicts(tmp_path):
    roster = _roster(tmp_path / "ОБЛІК.xlsx", ["ПЕРШИЙ Перший Перший"])
    book61 = write_full_book61(tmp_path / "book61.xlsm", {
        "ППД": [[1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 9, 1), None]],
    })
    cvs = write_cvs(tmp_path / "cvs.xlsx", {
        MEDICAL_SHEET: [[None, 1, "Підрозділ1", "сержант", "ПЕРШИЙ Перший Перший", "соматика", datetime(2026, 10, 3),
                         "Лікарня", "№1", datetime(2026, 10, 6)]],
    })
    report = str(tmp_path / "addons_mis_match.xlsx")

    result = cross_check_addons([book61, cvs], roster_path=roster, report_file_name=report)

    assert result.checked_files == [book61, cvs] and result.unrecognized_files == []
    assert result.report_path == report
    (conflict,) = result.conflicts
    assert [(p["start"], p["end"], p["status_a"], p["status_b"]) for p in conflict["periods"]] == [
        (date(2026, 10, 3), date(2026, 10, 5), "ППД", "ШП"),  # 06.10 - день виписки, не звіряється
    ]
    assert load_workbook(report).sheetnames == ["book61 - cvs"]
