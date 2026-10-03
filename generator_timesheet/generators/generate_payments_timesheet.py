from constants import (
    INFORMATION_UNIT_DIR,
    PAYMENT_MISMATCH_REPORT_FILE_NAME,
    PAYMENTS_OUTPUT_FILE_NAME,
    PAYMENTS_REVIEW_FILE_NAME,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    REPORT_DIR,
)
from content.information_unit_reader import read_payment_values
from content.oblik_timesheet import apply_payment_values, normalize_name
from content.payment_mismatch_checker import find_mismatches, read_information_unit_snapshot, write_mismatch_report
from generators.generate_timesheet import _build_timesheet, _finish
from utils.logging_utils import print_green, print_red


def _drop_unresolved_already_in_mismatch_report(unresolved, records):
    """За прямою вказівкою користувача - unresolved-запис apply_payment_values
    ("підрозділ подає [...] - не єдине числове значення категорії виплати") для
    (людина, дата), яка ВЖЕ показана як розбіжність у error_mis_statuses.xlsx
    (find_mismatches - DEFAULT_STATUS ЗАВЖДИ "порівнюваний", тож ЦЯ САМА
    розбіжність там ВЖЕ є - records["dates"]), дублює ТОЙ САМИЙ факт одразу в
    ДВОХ файлах - тут прибирається з unresolved, лишаючи ЛИШЕ
    error_mis_statuses.xlsx (детальніший звіт: показує КОНКРЕТНИЙ файл-
    джерело розбіжності, а не лише зведений список значень). Записи БЕЗ
    відповідника в records (напр. "Немає категорії виплати" - підрозділ
    узагалі нічого не подав) НЕ прибираються - для них error_mis_statuses.xlsx
    нічого не показує (немає даних для порівняння)."""
    flagged = {
        (normalize_name(record["roster_pib_raw"]), date_value)
        for record in records
        for date_value in record["dates"]
    }
    return [
        item for item in unresolved
        if (normalize_name(item.get("pib_raw") or ""), item.get("report_date")) not in flagged
    ]


def generate_payments_timesheet(
    report_dir=REPORT_DIR,
    personel_list_file_name=PERSONEL_LIST_FILE_NAME,
    personel_list_sheet_name=PERSONEL_LIST_SHEET_NAME,
    information_unit_dir=INFORMATION_UNIT_DIR,
    output_file_name=PAYMENTS_OUTPUT_FILE_NAME,
    review_file_name=PAYMENTS_REVIEW_FILE_NAME,
    mismatch_report_file_name=PAYMENT_MISMATCH_REPORT_FILE_NAME,
):
    """Як generate_timesheet (той самий день-за-днем табель зі щоденних
    рапортів), але ДРУГИМ проходом підставляє категорію виплати (число,
    напр. 100/30 - те, що подає власний підрозділ людини) у кожну клітинку
    зі статусом DEFAULT_STATUS ("РВЗ") - джерело content/
    information_unit_reader.py (resources/information_unit/*.xlsx, аркуші
    ТАБЕЛЬ/Табель і БЧС). За прямою вказівкою користувача - output_file_name
    ТОЙ САМИЙ файл "ОБЛІК.xlsx", що й у generate_timesheet (обидва прогони
    завжди окремі, output/ очищається перед кожним - колізії немає);
    review_file_name - ОКРЕМИЙ від generate_timesheet файл, і, на відміну
    від нього, НЕ створюється взагалі, коли unresolved порожній
    (skip_review_if_empty у _finish).

    ПІСЛЯ застосування (ще ДО збереження й ДО запису review_file_name) -
    звіряє ЦЕЙ САМЕ результат (timesheet) із КОЖНИМ файлом information_unit_dir
    ОКРЕМО (content/payment_mismatch_checker - read_information_unit_snapshot +
    find_mismatches) - ЩЕ ДО _finish, щоб unresolved-записи apply_payment_values,
    які ВЖЕ дублюють розбіжність, показану в error_mis_statuses.xlsx, можна
    було прибрати з review_file_name ПЕРЕД його записом
    (_drop_unresolved_already_in_mismatch_report вище). Записує розбіжності в
    mismatch_report_file_name (error_mis_statuses.xlsx) - АВТОМАТИЧНО, як
    частина ЦІЄЇ ж генерації, а не окремий пункт меню/окремий виклик. Так
    само, як review_file_name - файл узагалі НЕ створюється, коли розбіжностей
    немає (write_mismatch_report повертає None у цьому випадку).

    Повертає (output_path, unresolved, mismatch_report_path) - третій елемент
    ЦЕ те, що повернув write_mismatch_report: шлях, ЯКЩО є хоч одна розбіжність
    (інакше None) - дозволяє index.py вирішити, чи відкривати
    error_mis_statuses.xlsx автоматично, не читаючи сам файл повторно."""
    timesheet, unresolved = _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name)

    payment_values, skipped_files = read_payment_values(information_unit_dir)
    payment_unresolved = apply_payment_values(timesheet, payment_values)

    snapshot, _skipped_snapshot = read_information_unit_snapshot(information_unit_dir)
    records = find_mismatches(timesheet, snapshot)
    payment_unresolved = _drop_unresolved_already_in_mismatch_report(payment_unresolved, records)

    unresolved = unresolved + skipped_files + payment_unresolved

    output_path, unresolved = _finish(timesheet, unresolved, output_file_name, review_file_name, skip_review_if_empty=True)

    mismatch_report_path = write_mismatch_report(records, mismatch_report_file_name)
    if records:
        print_red(f"⚠ {len(records)} розбіжність(ей) ОБЛІК для виплат / information_unit - див. {mismatch_report_file_name}")
    else:
        print_green("Розбіжностей ОБЛІК для виплат / information_unit не знайдено.")

    return output_path, unresolved, mismatch_report_path


if __name__ == "__main__":  # pragma: no cover
    generate_payments_timesheet()
