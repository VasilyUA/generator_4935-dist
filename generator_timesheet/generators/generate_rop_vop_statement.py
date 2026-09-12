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
    SHORT_UNIT_BATTALION,
)
from content.daily_report_reader import latest_report_date
from content.information_unit_reader import read_payment_values
from content.oblik_timesheet import apply_payment_values
from content.rop_vop_statement import build_statement_rows, find_schedule_mismatches, write_schedule_mismatch_report, write_statement
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
    (спеціальних завдань)" - за прямою вказівкою користувача, за зразком
    реального файлу resources/{SHORT_UNIT_BATTALION без пробілу}ВОП-РОП_ЛИПЕНЬ_.xlsx: для КОЖНОГО дня
    ЦІЛЬОВОГО (за замовчуванням - year/month=None -> місяць НАЙПІЗНІШОГО
    рапорту в report_dir, content/daily_report_reader.latest_report_date;
    date.today(), лише якщо report_dir узагалі не містить жодного рапорту)
    місяця людини з категорією виплати 70 - "роп", 170 - "воп"
    (STATUS_MEANINGS, constants.py - РОП/ВОП, ротний/взводний опорний
    пункт). За прямою вказівкою користувача - НЕ date.today() за
    замовчуванням: реальний випадок - на початку нового місяця дані
    (рапорти/information_unit/schedule) за ЦЕЙ місяць ще НЕ подані (типова
    затримка), тож "поточний місяць" дав би ПОРОЖНІЙ результат, хоча дані
    за ОСТАННІЙ місяць, за який рапорти РЕАЛЬНО є, цілком готові. Той самий
    пайплайн, що й generate_payments_timesheet.py (_build_timesheet +
    read_payment_values + apply_payment_values) - ПОВНІСТЮ незалежний виклик,
    не читає output/ОБЛІК.xlsx з диска й не залежить від того, який ще пункт
    меню запускався до цього.

    output_file_name - за замовчуванням (None) будується динамічно з
    output_dir + назви підрозділу + місяця (як і раніше); передається
    напряму лише для пункту меню "Згенерувати все" (index.py), де ОЧІКУЄТЬСЯ
    фіксована назва "Відомість 170_70.xlsx" - за прямою вказівкою
    користувача.

    За прямою вказівкою користувача - ПЕРЕД збереженням фінального файлу,
    звіряє ЩОЙНО ОБЧИСЛЕНИЙ результат із НАЙНОВІШИМ файлом schedule_dir на
    ЦЕЙ САМЕ year/month (content/schedule_reader.latest_schedule_file +
    read_schedule_marks - "офіційна", вручну підтверджена відомість, яку
    подають щодня). Немає такого файлу взагалі (schedule_dir порожня чи без
    файлів на цей місяць) - звірка ПРОПУСКАЄТЬСЯ, генерація продовжується
    ЯК РАНІШЕ (немає джерела - немає з чим звіряти, це НЕ помилка). Є файл,
    Є хоча б ОДНА розбіжність (content/rop_vop_statement.
    find_schedule_mismatches) - фінальний файл ВЗАГАЛІ НЕ зберігається:
    замість нього зберігається mismatch_report_file_name (за замовчуванням -
    output_dir/"error_mis_statuses_Відомість.xlsx") з переліком розбіжностей
    (write_schedule_mismatch_report) - результат generate_rop_vop_statement
    тоді (None, None, mismatch_report_path). Розбіжностей немає (чи файлу
    schedule немає) - звичайна генерація, як раніше, третій елемент
    результату - None.

    Повертає (output_path, people_count, mismatch_report_path).
    people_count = 0, якщо ЖОДНА людина не мала жодного дня РОП/ВОП цього
    місяця (файл усе одно зберігається - лише заголовок і підписний блок,
    немає сенсу приховувати сам факт "порожньо" від користувача, як для
    error_mis_statuses.xlsx - ЦЕЙ файл - формальний документ на кожен
    місяць, завжди очікуваний)."""
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
            print_red(f"⚠ {len(mismatch_records)} розбіжність(ей) у Відомості РОП/ВОП - див. {mismatch_report_path}")
            return None, None, mismatch_report_path

    records, days_in_month = build_statement_rows(timesheet, year, month)

    month_name = MONTH_NAMES_NOMINATIVE_UPPER[month]
    if output_file_name is None:
        output_file_name = os.path.join(output_dir, f"{SHORT_UNIT_BATTALION.replace(' ', '')}ВОП-РОП_{month_name}_.xlsx")
    output_path = write_statement(records, days_in_month, output_file_name, year, month)

    if records:
        print_green(f"{output_path} ({len(records)} осіб)")
    else:
        print_red(f"⚠ {output_path}: жодна людина не має днів РОП/ВОП за {month_name} {year}.")

    return output_path, len(records), None


if __name__ == "__main__":  # pragma: no cover
    generate_rop_vop_statement()
