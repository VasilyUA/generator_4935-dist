import glob
import os
import zipfile
from typing import NamedTuple, Optional

from openpyxl.utils.exceptions import InvalidFileException

from constants import (
    ADDON_MISMATCH_REPORT_FILE_NAME,
    ADDONS_CROSS_CHECK_REPORT_FILE_NAME,
    ADDONS_DIR,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
)
from content import addon_reader, cvs_reader, training_journal_reader
from content.addon_cross_checker import find_addon_conflicts, write_addon_conflicts_report
from content.addon_mismatch_checker import find_period_mismatches, unmapped_people_in_roster, write_addon_mismatch_report
from content.oblik_timesheet import Timesheet


class AddonCheckResult(NamedTuple):
    period_mismatches: list
    unmapped_people: list
    report_path: Optional[str]
    checked_files: list
    unrecognized_files: list
    implausible_date_records: list  # дата далеко в майбутньому - ймовірна помилка в році книги


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
    """(period-записи, нерозпізнані люди Алф2, тип книги - "book61"/"cvs"/
    "journal") книги path, або None - тип книги не розпізнано (ні book61, ні
    ЦВС, ні журнал обліку БП, чи це взагалі не книга Excel)."""
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
    return None


def check_addon_mismatches_file(oblik_path, addon_paths=None, report_file_name=ADDON_MISMATCH_REPORT_FILE_NAME):
    """Звіряє готовий oblik_path (обраний користувачем чи щойно згенерований
    "Табель") з усіма книгами resources/addons (addon_paths; None -
    addon_files(), без діалогу вибору). Тип кожної книги впізнається за
    вмістом (_read_addon), а записи всіх книг складаються в одну шкалу на
    людину (content/addon_mismatch_checker.py). report_file_name пишеться,
    лише коли є що показати: розбіжність періоду чи нерозпізнаний статус
    Алф2."""
    timesheet = Timesheet(oblik_path, PERSONEL_LIST_SHEET_NAME)
    period_records, unmapped, checked, unrecognized, files_by_kind = _read_addons(addon_paths)

    unmapped_people = unmapped_people_in_roster(timesheet, unmapped)
    period_mismatches = find_period_mismatches(timesheet, period_records, files_by_kind=files_by_kind)
    written_path = write_addon_mismatch_report(
        period_mismatches, unmapped_people, report_file_name,
        source_files=[_source_file_name(path) for path in checked],
    )
    return AddonCheckResult(period_mismatches, unmapped_people, written_path, checked, unrecognized,
                            addon_reader.implausible_date_records(period_records))


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


def cross_check_addons(addon_paths=None, roster_path=PERSONEL_LIST_FILE_NAME,
                       report_file_name=ADDONS_CROSS_CHECK_REPORT_FILE_NAME):
    """Звіряє книги resources/addons між собою (content/addon_cross_checker.py).
    Люди й дати - з Табеля roster_path (resources/ОБЛІК.xlsx). report_file_name
    пишеться, лише коли є розбіжності."""
    timesheet = Timesheet(roster_path, PERSONEL_LIST_SHEET_NAME)
    period_records, _unmapped, checked, unrecognized, _files_by_kind = _read_addons(addon_paths)
    conflicts = find_addon_conflicts(timesheet, period_records, [_source_file_name(path) for path in checked])
    written_path = write_addon_conflicts_report(
        conflicts, report_file_name, source_files=[_source_file_name(path) for path in checked],
    )
    return AddonCrossCheckResult(conflicts, written_path, checked, unrecognized,
                                 addon_reader.implausible_date_records(period_records))
