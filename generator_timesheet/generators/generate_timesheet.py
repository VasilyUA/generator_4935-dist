import glob
import os

import content.information_unit_reader as iur
from constants import (
    OUTPUT_FILE_NAME,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    REPORT_DIR,
    REVIEW_FILE_NAME,
)
from content.daily_report_reader import extract_status_events, report_date_from_filename
from content.oblik_timesheet import Timesheet, apply_events
from utils.logging_utils import print_green, print_red


def _review_line(item):
    """Один рядок звіту - reason ЗАВЖДИ каже, ЩО не так, але сам собою часто не
    каже, ПРО КОГО (якщо ПІБ не вписано в текст reason - напр. "Невідомий
    заголовок пункту рапорту") і, для подій, ЯКИЙ статус мав би застосуватись,
    якби не ця перешкода, і З ЯКОГО РЕЧЕННЯ рапорту він узятий (для вільного
    тексту розділу "Поза межами...") - усе це додається в дужках, якщо є."""
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
    os.makedirs(os.path.dirname(review_file_name) or ".", exist_ok=True)
    with open(review_file_name, "w", encoding="utf-8") as review_file:
        if not unresolved:
            review_file.write("Записів, що потребують ручної перевірки, не знайдено.\n")
            return
        for item in unresolved:
            review_file.write(_review_line(item) + "\n")


def _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name):
    """Читає рапорти й застосовує події - повертає (timesheet, unresolved),
    ще НЕ збережений і без записаного звіту - спільна основа для
    generate_timesheet і generate_payments_timesheet
    (generators/generate_payments_timesheet.py)."""
    # Тимчасовий lock-файл Word ("~$...", той самий принцип, що й
    # information_unit_reader._is_office_lock_file/schedule_reader.py) -
    # реальний випадок: користувач тримає рапорт СЬОГОДНІШНЬОГО дня
    # відкритим у Word під час генерації - інакше report_date_from_filename
    # не міг розпізнати дату з мангленого імені ("~$.08.2026 - щоденний
    # рапорт.docx") і файл потрапляв у unresolved як "не вдалося визначити
    # дату", хоча САМ рапорт із такою назвою насправді ІСНУЄ (це не помилка
    # користувача - лише сам lock-файл не є справжнім рапортом).
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
    unresolved.extend(apply_events(timesheet, all_events, [d for d, _ in dated_reports]))
    return timesheet, unresolved


def _finish(timesheet, unresolved, output_file_name, review_file_name, skip_review_if_empty=False):
    """skip_review_if_empty - за прямою вказівкою користувача, лише для
    generate_payments_timesheet: коли unresolved порожній, review_file_name
    узагалі НЕ створюється (немає сенсу писати файл лише з "не знайдено").
    generate_timesheet НЕ передає цей прапорець (лишається False) -
    Потребує_ручної_перевірки.txt і далі пишеться ЗАВЖДИ, навіть порожній."""
    timesheet.save(output_file_name)
    if not (skip_review_if_empty and not unresolved):
        _write_review_file(unresolved, review_file_name)

    if unresolved:
        print_red(f"⚠ {len(unresolved)} запис(ів) потребують ручної перевірки - див. {review_file_name}")
    print_green(output_file_name)
    return output_file_name, unresolved


def generate_timesheet(
    report_dir=REPORT_DIR,
    personel_list_file_name=PERSONEL_LIST_FILE_NAME,
    personel_list_sheet_name=PERSONEL_LIST_SHEET_NAME,
    output_file_name=OUTPUT_FILE_NAME,
    review_file_name=REVIEW_FILE_NAME,
):
    """Формує output_file_name (копію personel_list_file_name з оновленими
    колонками-датами) на основі щоденних рапортів (.docx) у report_dir: для
    кожної людини й кожної дати - новий статус, якщо рапорт цієї дати
    повідомляє про зміну, інакше - статус з попередньої дати ("продовжує
    статус"). Повертає (output_path, unresolved); unresolved так само
    записується у review_file_name (порожній список - усе застосовано без
    зауважень)."""
    timesheet, unresolved = _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name)
    return _finish(timesheet, unresolved, output_file_name, review_file_name)


if __name__ == "__main__":  # pragma: no cover
    generate_timesheet()
