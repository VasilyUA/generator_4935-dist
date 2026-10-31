import os

from constants import (
    CATEGORY_100_CHECK_FILE_NAME,
    INFORMATION_UNIT_DIR,
    COMMANDERS_MESSAGE_FILE_NAME,
    PAYMENT_MISMATCH_REPORT_FILE_NAME,
    PAYMENTS_OUTPUT_FILE_NAME,
    PAYMENTS_REVIEW_EXCLUDE_PAIRS,
    PAYMENTS_REVIEW_FILE_NAME,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    REPORT_DIR,
)
from content.information_unit_reader import read_information_unit_snapshot, snapshot_values
from content.oblik_timesheet import Timesheet, apply_payment_values, check_payment_category_100_days, normalize_name
from content.payment_mismatch_checker import find_mismatches, write_mismatch_message, write_mismatch_report
from generators.generate_timesheet import _build_timesheet, _finish, _write_review_file
from utils.logging_utils import print_green, print_red


def _drop_unresolved_already_in_mismatch_report(unresolved, records):
    """Прибирає з unresolved записи apply_payment_values ("підрозділ подає
    [...] - не єдине числове значення категорії виплати") для тих людини й
    дати, що вже показані як розбіжність в error_mis_statuses.xlsx
    (find_mismatches завжди порівнює DEFAULT_STATUS, тож ця розбіжність там
    уже є). Інакше той самий факт був би у двох файлах; лишається докладніший
    звіт - він показує конкретний файл-джерело. Записи без відповідника в
    records (напр. "Немає категорії виплати" - підрозділ нічого не подав)
    лишаються: error_mis_statuses.xlsx про них нічого не каже."""
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
    mismatch_message_file_name=None,
):
    """Той самий табель, що й generate_timesheet, але другим проходом у кожну
    клітинку зі статусом DEFAULT_STATUS ("РВЗ") підставляється категорія
    виплати (число, напр. 100/30), яку подає підрозділ людини у
    resources/information_unit/*.xlsx (аркуші ТАБЕЛЬ і БЧС, content/
    information_unit_reader.py). output_file_name - той самий "ОБЛІК.xlsx",
    що й у generate_timesheet: прогони завжди окремі, а output/ очищається
    перед кожним. review_file_name - окремий файл, і, як і там, не
    створюється, якщо зауважень немає.

    Ще до збереження результат звіряється з кожним файлом information_unit
    окремо (find_mismatches) - щоб прибрати з review_file_name зауваження, які
    вже є в error_mis_statuses.xlsx. Розбіжності пишуться в
    mismatch_report_file_name (error_mis_statuses.xlsx) у межах цієї ж
    генерації; якщо їх немає, файл не створюється. Разом зі звітом пишеться
    повідомлення командирам mismatch_message_file_name (None - поруч зі
    звітом, з назвою COMMANDERS_MESSAGE_FILE_NAME).

    Повертає (output_path, unresolved, mismatch_report_path). Третій елемент -
    шлях до error_mis_statuses.xlsx, якщо є хоч одна розбіжність, інакше
    None: так index.py знає, чи пропонувати відкрити звіт."""
    timesheet, unresolved = _build_timesheet(report_dir, personel_list_file_name, personel_list_sheet_name)

    # Файли information_unit читаються один раз: той самий знімок дає і
    # категорії виплати, і звірку з кожним файлом окремо.
    snapshot, skipped_files = read_information_unit_snapshot(information_unit_dir)
    payment_unresolved = apply_payment_values(timesheet, snapshot_values(snapshot))
    records = find_mismatches(timesheet, snapshot)
    payment_unresolved = _drop_unresolved_already_in_mismatch_report(payment_unresolved, records)
    # Правило категорії 100 сюди не входить: index.py одразу після генерації
    # пише його в окремий CATEGORY_100_CHECK_FILE_NAME, і ті самі рядки у двох
    # файлах дублювались (рішення користувача).

    unresolved = unresolved + skipped_files + payment_unresolved

    output_path, unresolved = _finish(
        timesheet, unresolved, output_file_name, review_file_name,
        exclude_pairs=PAYMENTS_REVIEW_EXCLUDE_PAIRS,
    )

    mismatch_report_path = write_mismatch_report(records, mismatch_report_file_name)
    if records:
        print_red(f"{len(records)} розбіжність(ей) ОБЛІК для виплат / information_unit - див. {mismatch_report_file_name}")
        # None - повідомлення лягає поруч зі звітом.
        message_path = write_mismatch_message(records, mismatch_message_file_name or os.path.join(
            os.path.dirname(mismatch_report_file_name), os.path.basename(COMMANDERS_MESSAGE_FILE_NAME),
        ))
        if message_path:
            print_green(f"Повідомлення командирам підрозділів - {message_path}")
    else:
        print_green("Розбіжностей ОБЛІК для виплат / information_unit не знайдено.")

    return output_path, unresolved, mismatch_report_path


def check_payment_category_100_file(file_path, check_file_name=CATEGORY_100_CHECK_FILE_NAME):
    """Правило "категорія виплати 100 - не більше ніж
    PAYMENT_CATEGORY_100_MAX_DAYS днів поспіль у місяці"
    (check_payment_category_100_days, винятки - EXCLUDE_NAMES) для
    готового файлу ОБЛІК - обраного користувачем чи щойно згенерованого, без
    перебудови табеля й без запису ОБЛІК. Порушення пишуться в
    check_file_name, лише якщо вони є. Повертає (violations, check_file_name
    або None)."""
    timesheet = Timesheet(file_path, PERSONEL_LIST_SHEET_NAME)
    violations = check_payment_category_100_days(timesheet)
    if not violations:
        return violations, None
    _write_review_file(violations, check_file_name)
    return violations, check_file_name


if __name__ == "__main__":  # pragma: no cover
    check_payment_category_100_file(generate_payments_timesheet()[0])
