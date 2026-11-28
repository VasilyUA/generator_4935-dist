import glob
import os
import zipfile
from typing import NamedTuple, Optional

from openpyxl.utils.exceptions import InvalidFileException

from constants import (
    ADDON_BOOK_KINDS,
    ADDON_MISMATCH_IGNORE_LIST,
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
from content.addon_mismatch_checker import (
    drop_ignored_dates,
    find_period_mismatches,
    unmapped_people_in_roster,
    write_addon_mismatch_report,
)
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


class AddonContent(NamedTuple):
    records: list  # period-записи
    unmapped: list  # нерозпізнані люди Алф2 (book61)
    kind: str  # тип книги з ADDON_BOOK_KINDS
    without_bzvp: tuple = ()  # люди Алф1 без БЗВП (book61), з "source_file"


def _read_addon(path):
    """AddonContent книги path, або None - тип книги не розпізнано (жоден з
    ADDON_BOOK_KINDS чи це взагалі не книга Excel)."""
    try:
        wb = addon_reader.load_addon_workbook(path)
    except (InvalidFileException, zipfile.BadZipFile):
        return None
    source_file = _source_file_name(path)
    if addon_reader.is_book61_workbook(wb):
        without_bzvp = tuple({**person, "source_file": source_file}
                             for person in addon_reader.read_alf1_people_without_bzvp(wb))
        return AddonContent(addon_reader.read_all_period_records(wb, source_file),
                            addon_reader.read_unmapped_alf2_people(wb), "book61", without_bzvp)
    if cvs_reader.is_cvs_workbook(wb):
        return AddonContent(cvs_reader.read_all_period_records(wb, source_file), [], "cvs")
    if training_journal_reader.is_training_journal_workbook(wb):
        return AddonContent(training_journal_reader.read_all_period_records(wb, source_file), [], "journal")
    if treatment_journal_reader.is_treatment_journal_workbook(wb):
        return AddonContent(treatment_journal_reader.read_all_period_records(wb, source_file), [],
                            ADDON_TREATMENT_JOURNAL_KIND)
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
                                information_unit_dir=INFORMATION_UNIT_DIR, ignore_list=ADDON_MISMATCH_IGNORE_LIST):
    """Звіряє готовий oblik_path (обраний користувачем чи щойно згенерований
    "Табель") з усіма книгами resources/addons (addon_paths; None -
    addon_files(), без діалогу вибору). Тип кожної книги впізнається за
    вмістом (_read_addon), а записи всіх книг складаються в одну шкалу на
    людину (content/addon_mismatch_checker.py). report_file_name пишеться,
    лише коли є що показати: розбіжність періоду чи нерозпізнаний статус
    Алф2. report_dir (index.py передає REPORT_DIR) вмикає зворотну перевірку
    чисел ОБЛІК для виплат і дні прибуття "після 17.00"
    (_restore_facts_from_reports). ignore_list (IGNORE_LISTS з НАЛАШТУВАННЯ.xlsx) -
    дати людей, яких не показувати (drop_ignored_dates). Люди Алф1 без БЗВП
    на 30/70/100/170 - попередження (find_period_mismatches)."""
    timesheet = Timesheet(oblik_path, PERSONEL_LIST_SHEET_NAME)
    if report_dir is not None:
        _restore_facts_from_reports(timesheet, report_dir, personel_list_file_name, information_unit_dir)
    addons = _read_addons(addon_paths)

    unmapped_people = unmapped_people_in_roster(timesheet, addons.unmapped)
    period_mismatches = drop_ignored_dates(
        find_period_mismatches(timesheet, addons.period_records, files_by_kind=addons.files_by_kind,
                               without_bzvp=addons.without_bzvp),
        ignore_list, addons.kind_by_file,
    )
    written_path = write_addon_mismatch_report(
        period_mismatches, unmapped_people, report_file_name,
        source_files=[_source_file_name(path) for path in addons.checked],
    )
    return AddonCheckResult(period_mismatches, unmapped_people, written_path, addons.checked, addons.unrecognized,
                            addon_reader.implausible_date_records(addons.period_records),
                            _missing_books(addons.files_by_kind))


def _missing_books(files_by_kind):
    """Назви типів книг (ADDON_BOOK_KINDS), яких серед звірених немає."""
    return tuple(name for kind, name in ADDON_BOOK_KINDS.items() if kind not in files_by_kind)


class AddonsContent(NamedTuple):
    period_records: list
    unmapped: list  # нерозпізнані люди Алф2
    without_bzvp: list  # люди Алф1 без БЗВП
    checked: list  # розпізнані книги
    unrecognized: list  # нерозпізнані книги
    files_by_kind: dict  # {тип книги: назва першої такої}
    kind_by_file: dict  # {назва книги: її тип}


def _read_addons(addon_paths):
    """AddonsContent усіх книг addon_paths (None - addon_files())."""
    content = AddonsContent([], [], [], [], [], {}, {})
    for path in addon_files() if addon_paths is None else addon_paths:
        addon = _read_addon(path)
        if addon is None:
            content.unrecognized.append(path)
            continue
        content.period_records.extend(addon.records)
        content.unmapped.extend(addon.unmapped)
        content.without_bzvp.extend(addon.without_bzvp)
        content.checked.append(path)
        content.files_by_kind.setdefault(addon.kind, _source_file_name(path))
        content.kind_by_file[_source_file_name(path)] = addon.kind
    return content


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
    addons = _read_addons(addon_paths)
    source_files = [_source_file_name(path) for path in addons.checked]
    conflicts = find_addon_conflicts(timesheet, addons.period_records, source_files)
    written_path = write_addon_conflicts_report(conflicts, report_file_name, source_files=source_files)
    return AddonCrossCheckResult(conflicts, written_path, addons.checked, addons.unrecognized,
                                 addon_reader.implausible_date_records(addons.period_records),
                                 _missing_books(addons.files_by_kind))
