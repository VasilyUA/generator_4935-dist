import os
from datetime import date

from constants import (
    INFORMATION_UNIT_DIR,
    MONTH_NAMES_NOMINATIVE_UPPER,
    OUTPUT_DIR,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    REPORT_DIR,
    SCHEDULE_DIR,
    COMMANDERS_MESSAGE_FILE_NAME,
    SHORT_UNIT_BATTALION,
)
from content.daily_report_reader import latest_report_date
from content.information_unit_reader import read_payment_values
from content.oblik_timesheet import apply_payment_values
from content.rop_vop_statement import (
    build_statement_rows,
    find_schedule_mismatches,
    write_schedule_mismatch_message,
    write_schedule_mismatch_report,
    write_statement,
)
from content.schedule_reader import latest_schedule_file, read_schedule_marks
from generators.generate_timesheet import _build_timesheet
from utils.logging_utils import print_green, print_red


def generate_rop_vop_statement(
    report_dir=REPORT_DIR,
    personel_list_file_name=PERSONEL_LIST_FILE_NAME,
    personel_list_sheet_name=PERSONEL_LIST_SHEET_NAME,
    information_unit_dir=INFORMATION_UNIT_DIR,
    schedule_dir=SCHEDULE_DIR,
    output_dir=OUTPUT_DIR,
    output_file_name=None,
    mismatch_report_file_name=None,
    year=None,
    month=None,
):
    """"Відомість обліку днів участі військовослужбовців у виконанні бойових
    (спеціальних завдань)" за зразком реального файлу
    resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx: для кожного
    дня цільового місяця людина з категорією виплати 70 отримує "роп", зі 170 -
    "воп" (ротний і взводний опорний пункт, STATUS_MEANINGS).

    Цільовий місяць (year/month=None) - місяць найпізнішого рапорту в
    report_dir, а date.today() - лише якщо рапортів немає зовсім. На початку
    місяця дані за нього ще не подані (звичайна затримка), і "поточний місяць"
    дав би порожню відомість, хоча дані за останній місяць з рапортами готові.
    Розрахунок той самий, що й у generate_payments_timesheet (_build_timesheet +
    read_payment_values + apply_payment_values), але незалежний: output/ОБЛІК.xlsx
    не читається, і неважливо, які пункти меню запускались раніше.

    output_file_name (None) будується з output_dir, назви підрозділу й місяця;
    "Згенерувати все" (index.py) передає фіксовану назву "Відомість 170_70.xlsx".

    Перед збереженням результат звіряється з найновішим файлом schedule_dir за
    цей місяць (latest_schedule_file + read_schedule_marks) - вручну
    підтвердженою відомістю, яку подають щодня. Якщо такого файлу немає,
    звірка просто пропускається. Якщо є хоч одна розбіжність
    (find_schedule_mismatches), Відомість не зберігається - замість неї
    пишеться mismatch_report_file_name (за замовчуванням
    output_dir/"error_mis_statuses_Відомість.xlsx"), і результат - (None, None,
    mismatch_report_path).

    Повертає (output_path, people_count, mismatch_report_path). Якщо днів
    РОП/ВОП за місяць немає ні в кого, people_count = 0, а файл однаково
    зберігається (лише заголовок і підписний блок): це формальний документ,
    який чекають щомісяця."""
    if year is None or month is None:
        fallback = latest_report_date(report_dir) or date.today()
        year = year or fallback.year
        month = month or fallback.month
    if mismatch_report_file_name is None:
        mismatch_report_file_name = os.path.join(output_dir, "error_mis_statuses_Відомість.xlsx")

    timesheet, _unresolved = _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name)
    payment_values, _skipped_files = read_payment_values(information_unit_dir)
    apply_payment_values(timesheet, payment_values)

    schedule_path = latest_schedule_file(schedule_dir, year, month)
    if schedule_path is not None:
        schedule_marks, schedule_people = read_schedule_marks(schedule_path)
        mismatch_records = find_schedule_mismatches(timesheet, year, month, schedule_marks, schedule_people)
        if mismatch_records:
            mismatch_report_path = write_schedule_mismatch_report(mismatch_records, mismatch_report_file_name)
            print_red(f"{len(mismatch_records)} розбіжність(ей) у Відомості РОП/ВОП - див. {mismatch_report_path}")
            # Поруч зі звітом - повідомлення командирам (без днів, де файл БЧС
            # підрозділу вже збігається з відомістю).
            message_path = write_schedule_mismatch_message(mismatch_records, os.path.join(
                os.path.dirname(mismatch_report_file_name), os.path.basename(COMMANDERS_MESSAGE_FILE_NAME),
            ), submitted_values=payment_values)
            if message_path:
                print_green(f"Повідомлення командирам підрозділів - {message_path}")
            return None, None, mismatch_report_path

    records, days_in_month = build_statement_rows(timesheet, year, month)

    month_name = MONTH_NAMES_NOMINATIVE_UPPER[month]
    if output_file_name is None:
        output_file_name = os.path.join(output_dir, f"{SHORT_UNIT_BATTALION.replace(' ', '')}ВОП-РОП_{month_name}_.xlsx")
    output_path = write_statement(records, days_in_month, output_file_name, year, month)

    if records:
        print_green(f"{output_path} ({len(records)} осіб)")
    else:
        print_red(f"{output_path}: жодна людина не має днів РОП/ВОП за {month_name} {year}.")

    return output_path, len(records), None


if __name__ == "__main__":  # pragma: no cover
    generate_rop_vop_statement()
