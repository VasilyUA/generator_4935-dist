import glob
import os
import zipfile
from typing import NamedTuple, Optional

from openpyxl.utils.exceptions import InvalidFileException

from constants import (
    ADDON_BOOK_KINDS,
    ADDON_MISMATCH_REPORT_FILE_NAME,
    ADDON_TREATMENT_JOURNAL_KIND,
    ADDONS_CROSS_CHECK_REPORT_FILE_NAME,
    ADDONS_DIR,
    INFORMATION_UNIT_DIR,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
)
from content import addon_reader, cvs_reader, training_journal_reader, treatment_journal_reader
from content.addon_cross_checker import find_addon_conflicts, write_addon_conflicts_report
from content.addon_mismatch_checker import find_period_mismatches, unmapped_people_in_roster, write_addon_mismatch_report
from content.information_unit_reader import read_payment_values
from content.oblik_timesheet import Timesheet, apply_payment_values
from generators.generate_timesheet import _build_timesheet
from utils.logging_utils import print_yellow


class AddonCheckResult(NamedTuple):
    period_mismatches: list
    unmapped_people: list
    report_path: Optional[str]
    checked_files: list
    unrecognized_files: list
    implausible_date_records: list  # дата далеко в майбутньому - ймовірна помилка в році книги
    missing_books: tuple = ()  # назви типів книг, яких немає, - звірено без них


def addon_files(addons_dir=ADDONS_DIR):
    """Усі книги Excel теки addons_dir (без lock-файлів Excel "~$...")."""
    return sorted(
        path for path in glob.glob(os.path.join(addons_dir, "*.xls[xm]"))
        if not os.path.basename(path).startswith("~$")
    )


def _source_file_name(path):
    """Ім'я книги без розширення - префікс "Джерело" і назва аркуша mis_match.xlsx."""
    return os.path.splitext(os.path.basename(path))[0]


def _read_addon(path):
    """(period-записи, нерозпізнані люди Алф2, тип книги з ADDON_BOOK_KINDS)
    книги path, або None - тип книги не розпізнано (жоден з ADDON_BOOK_KINDS чи
    це взагалі не книга Excel)."""
    try:
        wb = addon_reader.load_addon_workbook(path)
    except (InvalidFileException, zipfile.BadZipFile):
        return None
    source_file = _source_file_name(path)
    if addon_reader.is_book61_workbook(wb):
        return (addon_reader.read_all_period_records(wb, source_file), addon_reader.read_unmapped_alf2_people(wb),
                "book61")
    if cvs_reader.is_cvs_workbook(wb):
        return cvs_reader.read_all_period_records(wb, source_file), [], "cvs"
    if training_journal_reader.is_training_journal_workbook(wb):
        return training_journal_reader.read_all_period_records(wb, source_file), [], "journal"
    if treatment_journal_reader.is_treatment_journal_workbook(wb):
        return treatment_journal_reader.read_all_period_records(wb, source_file), [], ADDON_TREATMENT_JOURNAL_KIND
    return None


def _restore_facts_from_reports(timesheet, report_dir, personel_list_file_name, information_unit_dir):
    """Заповнює те, чого в готовому файлі вже немає:
    timesheet.statuses_behind_numbers і statuses_behind_categories (лишилось
    тільки значення, ВД чи ППД -> 10, ШП чи ВЛК -> 10_МЕД, а зворотній
    перевірці потрібен статус за ним) і timesheet.after_hours_arrivals (звідки
    цей ППД). Розрахунок той самий, що й
    у generate_payments_timesheet, і, як і Відомість РОП/ВОП, не залежить від
    файлу (реальний output/ОБЛІК.xlsx відтворює до клітинки). Рядки обраного
    вручну файлу можуть іти інакше, тож людина шукається за ПІБ. Без рапортів
    числа лишаються без статусу й запису в книгах не потребують."""
    try:
        rebuilt, _unresolved = _build_timesheet(report_dir, personel_list_file_name, PERSONEL_LIST_SHEET_NAME)
    except FileNotFoundError as e:
        print_yellow(f"Числа в ОБЛІК для виплат звірено без статусів за ними: {e}")
        return
    payment_values, _skipped_files = read_payment_values(information_unit_dir)
    apply_payment_values(rebuilt, payment_values)

    person_rows = timesheet.person_rows()
    target_rows = {row_idx: person_rows[name] for name, row_idx in rebuilt.person_rows().items() if name in person_rows}
    for behind, restored in ((rebuilt.statuses_behind_numbers, timesheet.statuses_behind_numbers),
                             (rebuilt.statuses_behind_categories, timesheet.statuses_behind_categories)):
        for (row_idx, date_value), status in behind.items():
            if row_idx in target_rows:
                restored[(target_rows[row_idx], date_value)] = status
    timesheet.after_hours_arrivals.update(
        (target_rows[row_idx], date_value) for row_idx, date_value in rebuilt.after_hours_arrivals if row_idx in target_rows
    )


def check_addon_mismatches_file(oblik_path, addon_paths=None, report_file_name=ADDON_MISMATCH_REPORT_FILE_NAME,
                                report_dir=None, personel_list_file_name=PERSONEL_LIST_FILE_NAME,
                                information_unit_dir=INFORMATION_UNIT_DIR):
    """Звіряє готовий oblik_path (обраний користувачем чи щойно згенерований
    "Табель") з усіма книгами resources/addons (addon_paths; None -
    addon_files(), без діалогу вибору). Тип кожної книги впізнається за
    вмістом (_read_addon), а записи всіх книг складаються в одну шкалу на
    людину (content/addon_mismatch_checker.py). report_file_name пишеться,
    лише коли є що показати: розбіжність періоду чи нерозпізнаний статус
    Алф2. report_dir (index.py передає REPORT_DIR) вмикає зворотну перевірку
    чисел ОБЛІК для виплат і дні прибуття "після 17.00"
    (_restore_facts_from_reports)."""
    timesheet = Timesheet(oblik_path, PERSONEL_LIST_SHEET_NAME)
    if report_dir is not None:
        _restore_facts_from_reports(timesheet, report_dir, personel_list_file_name, information_unit_dir)
    period_records, unmapped, checked, unrecognized, files_by_kind = _read_addons(addon_paths)

    unmapped_people = unmapped_people_in_roster(timesheet, unmapped)
    period_mismatches = find_period_mismatches(timesheet, period_records, files_by_kind=files_by_kind)
    written_path = write_addon_mismatch_report(
        period_mismatches, unmapped_people, report_file_name,
        source_files=[_source_file_name(path) for path in checked],
    )
    return AddonCheckResult(period_mismatches, unmapped_people, written_path, checked, unrecognized,
                            addon_reader.implausible_date_records(period_records), _missing_books(files_by_kind))


def _missing_books(files_by_kind):
    """Назви типів книг (ADDON_BOOK_KINDS), яких серед звірених немає."""
    return tuple(name for kind, name in ADDON_BOOK_KINDS.items() if kind not in files_by_kind)


def _read_addons(addon_paths):
    """(period-записи, нерозпізнані люди Алф2, розпізнані книги, нерозпізнані
    книги, {тип книги: назва першої такої}) - усі книги addon_paths (None -
    addon_files())."""
    period_records, unmapped, checked, unrecognized, files_by_kind = [], [], [], [], {}
    for path in addon_files() if addon_paths is None else addon_paths:
        content = _read_addon(path)
        if content is None:
            unrecognized.append(path)
            continue
        records, people, kind = content
        period_records.extend(records)
        unmapped.extend(people)
        checked.append(path)
        files_by_kind.setdefault(kind, _source_file_name(path))
    return period_records, unmapped, checked, unrecognized, files_by_kind


class AddonCrossCheckResult(NamedTuple):
    conflicts: list
    report_path: Optional[str]
    checked_files: list
    unrecognized_files: list
    implausible_date_records: list  # дата далеко в майбутньому - ймовірна помилка в році книги
    missing_books: tuple = ()  # назви типів книг, яких немає, - звірено без них


def cross_check_addons(addon_paths=None, roster_path=PERSONEL_LIST_FILE_NAME,
                       report_file_name=ADDONS_CROSS_CHECK_REPORT_FILE_NAME):
    """Звіряє книги resources/addons між собою (content/addon_cross_checker.py).
    Люди й дати - з Табеля roster_path (resources/ОБЛІК.xlsx). report_file_name
    пишеться, лише коли є розбіжності."""
    timesheet = Timesheet(roster_path, PERSONEL_LIST_SHEET_NAME)
    period_records, _unmapped, checked, unrecognized, files_by_kind = _read_addons(addon_paths)
    conflicts = find_addon_conflicts(timesheet, period_records, [_source_file_name(path) for path in checked])
    written_path = write_addon_conflicts_report(
        conflicts, report_file_name, source_files=[_source_file_name(path) for path in checked],
    )
    return AddonCrossCheckResult(conflicts, written_path, checked, unrecognized,
                                 addon_reader.implausible_date_records(period_records), _missing_books(files_by_kind))
