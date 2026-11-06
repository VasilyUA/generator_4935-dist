import glob
import os

import content.information_unit_reader as iur
from constants import (
    LEGACY_TRANSFER_STATUS,
    OUTPUT_FILE_NAME,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    PIB_COLUMN_NAME,
    REPORT_DIR,
    REVIEW_EXCLUDE_PAIRS,
    REVIEW_FILE_NAME,
    TRANSFER_TO_BRIGADE_STATUS,
    TRANSFER_TO_OTHER_UNIT_STATUS,
)
from content.daily_report_reader import extract_status_events, report_date_from_filename
from content.oblik_timesheet import Timesheet, apply_events, normalize_name
from utils.logging_utils import print_green, print_red


def _review_line(item):
    """Рядок файлу "Потребує ручної перевірки". reason каже, що не так, але
    часто не каже, про кого (якщо ПІБ немає в тексті reason, напр. "Невідомий
    заголовок пункту рапорту"), який статус мав би застосуватись і з якого
    речення рапорту він узятий (для вільного тексту "Поза межами..."). Усе це,
    якщо є, додається в дужках."""
    report_date = item.get("report_date")
    prefix = f"[{report_date.strftime('%d.%m.%Y')}] " if report_date else ""

    details = []
    pib_raw = item.get("pib_raw")
    if pib_raw and pib_raw not in item["reason"]:
        details.append(f"ПІБ: {pib_raw}")
    if item.get("status"):
        details.append(f"мав би застосуватись статус \"{item['status']}\"")
    if item.get("raw_text"):
        details.append(f"текст рапорту: \"{item['raw_text']}\"")

    suffix = f" ({'; '.join(details)})" if details else ""
    return f"{prefix}{item['reason']}{suffix}"


def _write_review_file(unresolved, review_file_name):
    """Викликається лише тоді, коли є що перевіряти: порожнього файлу не пишемо."""
    os.makedirs(os.path.dirname(review_file_name) or ".", exist_ok=True)
    with open(review_file_name, "w", encoding="utf-8") as review_file:
        for item in unresolved:
            review_file.write(_review_line(item) + "\n")


def _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name):
    """Читає рапорти й застосовує події. Повертає (timesheet, unresolved) -
    ще не збережений табель і без записаного звіту. Спільна основа для
    generate_timesheet і generate_payments_timesheet."""
    # Lock-файли Word ("~$...") пропускаємо: коли сьогоднішній рапорт відкритий
    # у Word, з назви такого файлу ("~$.08.2026 - щоденний рапорт.docx") дату
    # не прочитати, і він хибно потрапляв у "не вдалося визначити дату", хоча
    # сам рапорт є.
    report_paths = [
        path for path in glob.glob(os.path.join(report_dir, "*.docx"))
        if not iur._is_office_lock_file(path)
    ]
    if not report_paths:
        raise FileNotFoundError(f"У теці {report_dir} не знайдено жодного .docx рапорту.")

    unresolved = []
    dated_reports = []
    for path in report_paths:
        report_date = report_date_from_filename(path)
        if report_date is None:
            unresolved.append({"reason": f"Не вдалося визначити дату з назви файлу {path}.", "report_date": None})
            continue
        dated_reports.append((report_date, path))
    dated_reports.sort()

    all_events = []
    for report_date, path in dated_reports:
        events, report_unresolved = extract_status_events(path)
        all_events.extend(events)
        unresolved.extend(report_unresolved)

    timesheet = Timesheet(personel_list_file_name, personel_list_sheet_name)
    unresolved.extend(_legacy_transfer_cells(timesheet))
    unresolved.extend(apply_events(timesheet, all_events, [d for d, _ in dated_reports]))
    return timesheet, unresolved


def _legacy_transfer_cells(timesheet):
    """Статусу "ПРВД" більше немає - замість нього ПРВДБ і ПРВДВЧ. Для
    вручну вписаного в resources/ОБЛІК.xlsx "ПРВД" (перша така дата людини) -
    нагадування замінити, бо з нього не видно, куди людину переведено."""
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    reminders = []
    for row_idx in timesheet.person_rows().values():
        for col_idx, date_value in timesheet.date_columns:
            value = timesheet.ws.cell(row=row_idx, column=col_idx).value
            if isinstance(value, str) and value.strip().upper() == LEGACY_TRANSFER_STATUS:
                reminders.append({
                    "reason": (
                        f"Статус \"{LEGACY_TRANSFER_STATUS}\" більше не використовується - замініть у "
                        f"{timesheet.file_path} на \"{TRANSFER_TO_BRIGADE_STATUS}\" (переведений у підрозділи "
                        f"бригади) чи \"{TRANSFER_TO_OTHER_UNIT_STATUS}\" (в іншу військову частину)."
                    ),
                    "pib_raw": timesheet.ws.cell(row=row_idx, column=pib_col).value,
                    "report_date": date_value,
                })
                break
    return reminders


def _without_excluded_pairs(unresolved, exclude_pairs):
    """Прибирає записи, чия пара (ПІБ, дата) є в exclude_pairs. ПІБ
    порівнюється через normalize_name, як і ПІБ табеля, дата - report_date (та,
    що в дужках [дд.мм.рррр] у рядку файлу). Той самий ПІБ на іншу дату
    лишається; запис без ПІБ чи без дати жодній парі не відповідає."""
    excluded = {(normalize_name(name), day) for day, name in exclude_pairs}
    return [
        item for item in unresolved
        if (normalize_name(item.get("pib_raw") or ""), item.get("report_date")) not in excluded
    ]


def _finish(timesheet, unresolved, output_file_name, review_file_name, exclude_pairs=()):
    """Зберігає табель і, якщо після exclude_pairs лишилось що перевіряти,
    файл review_file_name (порожнім він не створюється). exclude_pairs - пари
    (дата, ПІБ) з resources/data.json, які не треба показувати в цьому файлі;
    вони прибираються і з файлу, і з результату."""
    unresolved = _without_excluded_pairs(unresolved, exclude_pairs)
    timesheet.save(output_file_name)
    if unresolved:
        _write_review_file(unresolved, review_file_name)
        print_red(f"{len(unresolved)} запис(ів) потребують ручної перевірки - див. {review_file_name}")
    print_green(output_file_name)
    return output_file_name, unresolved


def generate_timesheet(
    report_dir=REPORT_DIR,
    personel_list_file_name=PERSONEL_LIST_FILE_NAME,
    personel_list_sheet_name=PERSONEL_LIST_SHEET_NAME,
    output_file_name=OUTPUT_FILE_NAME,
    review_file_name=REVIEW_FILE_NAME,
):
    """Будує output_file_name - копію personel_list_file_name з заповненими
    колонками-датами за щоденними рапортами (.docx) з report_dir. Для кожної
    людини й дати: новий статус, якщо рапорт цієї дати повідомляє про зміну,
    інакше - статус попередньої дати. Повертає (output_path, unresolved);
    unresolved записується і в review_file_name (порожній список - усе
    застосовано без зауважень)."""
    timesheet, unresolved = _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name)
    return _finish(timesheet, unresolved, output_file_name, review_file_name, exclude_pairs=REVIEW_EXCLUDE_PAIRS)


if __name__ == "__main__":  # pragma: no cover
    generate_timesheet()
