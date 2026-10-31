import os
import sys
import warnings
from collections import Counter
from datetime import date

from InquirerPy.prompts.list import ListPrompt

from constants import (
    ADDON_IMPLAUSIBLE_FUTURE_DAYS,
    ADDONS_DIR,
    INFORMATION_UNIT_DIR,
    INFORMATION_UNIT_KEEP_FILE_NAME,
    MONEY_PROJECT_RESOURCES_DIR,
    MONEY_PROJECT_SHEET_NAME,
    OUTPUT_DIR,
    PAYMENTS_REVIEW_FILE_NAME,
    REPORT_DIR,
    RESOURCES_DIR,
    REVIEW_FILE_NAME,
)
from content.addon_mismatch_checker import is_warning_only
from content.bchs_subdivisions import missing_subdivisions
from content.oblik_timesheet import export_for_money_project
from content.signal_bchs import SignalFetchError, fetch_bchs_files
from content.signal_reports import sync_reports_and_schedules
from generators.check_addon_mismatches import addon_files, check_addon_mismatches_file, cross_check_addons
from generators.generate_payments_timesheet import check_payment_category_100_file, generate_payments_timesheet
from generators.generate_rop_vop_statement import generate_rop_vop_statement
from generators.generate_timesheet import generate_timesheet
from utils.backup import backup_resources_and_output
from utils.folder_picker import pick_file, pick_folder
from utils.logging_utils import print_green, print_red, print_yellow
from utils.output_folder import clear_information_unit_directory, clear_output_directory, open_file, open_folder
from utils.prompts import ask_date, ask_yes_no

# Реальні resources/ОБЛІК.xlsx і файли information_unit мають розширення
# Excel, яких openpyxl не розбирає (умовне форматування й перевірка даних
# "extLst", малюнки), і він попереджає про це при кожному читанні. На
# результат це не впливає - проєкт читає лише значення клітинок і пише власні
# кольори, - тож ховаємо саме ці повідомлення, а не всі попередження.
for _message in (
    r"Conditional Formatting extension is not supported.*",
    r"Data Validation extension is not supported.*",
    r"DrawingML support is incomplete.*",
):
    warnings.filterwarnings("ignore", message=_message, category=UserWarning)

GENERATE_PERSONEL_LIST = "personel_list"
GENERATE_PAYMENTS_LIST = "payments_list"
GENERATE_ROP_VOP_STATEMENT = "rop_vop_statement"
GENERATE_ALL = "generate_all"
CHECK_CATEGORY_100 = "check_category_100"
CHECK_ADDON_MISMATCH = "check_addon_mismatch"
STOP_SCRIPT = "stop_script"

_GENERATE_CHOICES = [
    {"name": "Згенерувати все", "value": GENERATE_ALL},
    {"name": "Припинити скрипт", "value": STOP_SCRIPT},
    {"name": "Перевірка: категорія виплати 100 поспіль (без запису ОБЛІК)", "value": CHECK_CATEGORY_100},
    {"name": "Перевірка: звірка Облік.xlsx з addons (resources/addons, без запису ОБЛІК)", "value": CHECK_ADDON_MISMATCH},
    {"name": "Файл ОБЛІК для виплат з БЧС", "value": GENERATE_PAYMENTS_LIST},
    {"name": "Файл обліку особового складу (ОБЛІК.xlsx)", "value": GENERATE_PERSONEL_LIST},
    {"name": "Відомість РОП/ВОП", "value": GENERATE_ROP_VOP_STATEMENT},
]

_HANDLERS = {
    GENERATE_PERSONEL_LIST: generate_timesheet,
    GENERATE_PAYMENTS_LIST: generate_payments_timesheet,
}

# Файл "Потребує ручної перевірки", який пише відповідний генератор, коли
# його викликано без аргументів.
_REVIEW_FILE_NAME_BY_CHOICE = {
    GENERATE_PERSONEL_LIST: REVIEW_FILE_NAME,
    GENERATE_PAYMENTS_LIST: PAYMENTS_REVIEW_FILE_NAME,
}

# "Згенерувати все" кладе всі результати в одну теку. Обидва ОБЛІК за
# замовчуванням називаються "ОБЛІК.xlsx", тож стройовий отримує власну
# назву, а Відомість - фіксовану замість "підрозділ + місяць" (так попросив
# користувач).
_ALL_PERSONEL_LIST_OUTPUT_FILE_NAME = os.path.join(OUTPUT_DIR, "ОБЛІК по стройовій.xlsx")
_ALL_ROP_VOP_STATEMENT_OUTPUT_FILE_NAME = os.path.join(OUTPUT_DIR, "Відомість 170_70.xlsx")

_PICK_OBLIK_FILE = "pick_oblik_file"
_PAYMENTS_SUFFIX = " (для виплат)"


def _open_if_confirmed(file_path):
    """Питає, чи відкрити file_path (Enter - так). Після "Так" скрипт
    завершується: користувач іде працювати з цим файлом, решта кроків - у
    наступному запуску. "Ні" - скрипт іде далі."""
    if ask_yes_no(f"Відкрити {file_path}?", default=True):
        open_file(file_path)
        sys.exit(0)


def _check_category_100_in(file_path):
    """Правило "категорія виплати 100 поспіль" для файлу ОБЛІК: друкує
    результат і пропонує відкрити файл із порушеннями."""
    try:
        violations, check_path = check_payment_category_100_file(file_path)
    except (PermissionError, ValueError) as e:
        print_red(str(e))
        return
    if check_path is None:
        print_green("Порушень правила категорії 100 поспіль не знайдено.")
    else:
        print_red(f"{len(violations)} порушень - див. {check_path}")
        _open_if_confirmed(check_path)


def _pick_and_check_category_100():
    picked_file = pick_file("Оберіть файл ОБЛІК для перевірки", initial_dir=RESOURCES_DIR)
    if not picked_file:
        print_red("Файл не обрано - перевірку не виконано.")
        return
    _check_category_100_in(picked_file)


def _print_unrecognized_addons(paths):
    for path in paths:
        print_red(f"{os.path.basename(path)} - не розпізнано (ні book61, ні ЦВС, ні журнал обліку БП) - пропущено.")


def _print_implausible_dates(records):
    """Записи addons з датою, надто далекою в майбутньому
    (ADDON_IMPLAUSIBLE_FUTURE_DAYS) - найпевніше, помилка в році. Розбіжності
    через них хибні, а виправляти треба саму книгу, тож друкуємо файл, аркуш і
    ПІБ."""
    if not records:
        return
    print_red(
        f"Підозрілі дати в addons (пізніше за сьогодні більш ніж на {ADDON_IMPLAUSIBLE_FUTURE_DAYS} дн. - "
        "ймовірно, помилка в році; розбіжності через них хибні - виправте у файлі):"
    )
    for record in records:
        end = f"{record['end_date']:%d.%m.%Y}" if record.get("end_date") else "не вказано"
        print_red(f"  {record['source_file']}, аркуш «{record['source_sheet']}»: {record['pib_raw']} - "
                  f"{record['start_date']:%d.%m.%Y}-{end}")


def _check_addon_mismatches_in(oblik_path):
    """Звіряє аркуш "Табель" файлу oblik_path з усіма книгами resources/addons
    (тип кожної розпізнається за вмістом), друкує результат і пропонує
    відкрити mis_match.xlsx."""
    if not addon_files():
        print_red(f"У {ADDONS_DIR} немає жодної книги Excel - перевірку не виконано.")
        return
    try:
        result = check_addon_mismatches_file(oblik_path)
    except (PermissionError, ValueError) as e:
        print_red(str(e))
        return
    _print_unrecognized_addons(result.unrecognized_files)
    if not result.checked_files:
        return
    print_green("Звірено з: " + ", ".join(os.path.basename(path) for path in result.checked_files))
    _print_implausible_dates(result.implausible_date_records)
    if result.report_path is None:
        print_green("Розбіжностей з addons не знайдено.")
        return

    # Справжні розбіжності - червоним, а записи лише з попередженням
    # (відпустка без "Куди вибув", статуси збігаються) - окремо жовтим.
    warning_records = [record for record in result.period_mismatches if is_warning_only(record)]
    error_records = [record for record in result.period_mismatches if not is_warning_only(record)]
    if error_records or result.unmapped_people:
        by_file = Counter(record["source_file"] for record in error_records)
        per_file = ", ".join(f"{name} - {count}" for name, count in by_file.items())
        print_red(
            f"{len(error_records)} розбіжність(ей) періодів{f' ({per_file})' if per_file else ''}, "
            f"{len(result.unmapped_people)} нерозпізнаних статусів Алф2 - див. {result.report_path} "
            "(окремий аркуш на кожен файл addons)"
        )
    if warning_records:
        print_yellow(
            f"{len(warning_records)} попередження (не помилка): відпустка в book61 без \"Куди вибув\" - "
            f"жовтим у {result.report_path}"
        )
    _open_if_confirmed(result.report_path)


def _pick_and_check_addon_mismatches():
    """Користувач сам обирає файл ОБЛІК, книги addons беруться всі, без діалогу."""
    picked_oblik = pick_file("Оберіть файл ОБЛІК (аркуш Табель) для звірки", initial_dir=RESOURCES_DIR)
    if not picked_oblik:
        print_red("Файл ОБЛІК не обрано - перевірку не виконано.")
        return
    _check_addon_mismatches_in(picked_oblik)


def _choose_and_check_addon_mismatches(generated_obliks):
    """Після "Здійснити звірку з addons? Так": перевірити ОБЛІК, згенерований
    у цьому прогоні (generated_obliks - [(назва, шлях)], для виплат першим),
    чи обрати файл діалогом. Якщо нічого не згенеровано - одразу діалог."""
    if not generated_obliks:
        _pick_and_check_addon_mismatches()
        return
    picked = ListPrompt(
        message="Який файл ОБЛІК звірити з addons?",
        choices=[{"name": f"Перевірити згенерований {name}", "value": path} for name, path in generated_obliks]
        + [{"name": "Обрати файл ОБЛІК (.xlsx)", "value": _PICK_OBLIK_FILE}],
        default=generated_obliks[0][1],
    ).execute()
    if picked == _PICK_OBLIK_FILE:
        _pick_and_check_addon_mismatches()
    else:
        _check_addon_mismatches_in(picked)


def _cross_check_addons():
    """Звіряє книги resources/addons між собою (без Табеля), друкує результат і
    пропонує відкрити addons_mis_match.xlsx."""
    try:
        result = cross_check_addons()
    except (PermissionError, ValueError) as e:
        print_red(str(e))
        return
    _print_unrecognized_addons(result.unrecognized_files)
    if len(result.checked_files) < 2:
        print_red(f"У {ADDONS_DIR} менше двох розпізнаних книг - звіряти між собою нічого.")
        return
    print_green("Звірено між собою: " + ", ".join(os.path.basename(path) for path in result.checked_files))
    _print_implausible_dates(result.implausible_date_records)
    if result.report_path is None:
        print_green("Розбіжностей між книгами addons не знайдено.")
        return
    by_pair = Counter(" і ".join(record["files"]) for record in result.conflicts)
    print_red(
        "Розбіжності між книгами (людей): "
        + ", ".join(f"{pair} - {count}" for pair, count in by_pair.items())
        + f" - див. {result.report_path} (аркуш на кожен файл: його порівняння з усіма іншими)"
    )
    _open_if_confirmed(result.report_path)


def _dates_text(days):
    return ", ".join(day.strftime("%d.%m") for day in days)


def _sync_reports_from_signal():
    """Звіряє рапорти й відомості ВОП-РОП з групою Signal: нові
    довантажуються, застарілі замінюються найновішими. Якщо Signal
    недоступний - червоне повідомлення, і скрипт іде далі."""
    try:
        results = sync_reports_and_schedules()
    except SignalFetchError as e:
        print_red(str(e))
        return
    for result in results:
        parts = [f"нових {len(result.new)}" + (f" ({_dates_text(result.new)})" if result.new else ""),
                 f"оновлено {len(result.updated)}" + (f" ({_dates_text(result.updated)})" if result.updated else ""),
                 f"актуальних {len(result.actual)}"]
        print_green(f"Signal, {result.label}: " + ", ".join(parts) + ".")
        for file_name in result.failed:
            print_red(f"Signal, {result.label}: не вдалося завантажити {file_name}.")
        for day, file_name in result.missing:
            print_red(f"Signal, {result.label}: немає за {day:%d.%m.%Y} - {file_name} (ні в Signal, ні в теці).")


def _fetch_bchs_files():
    """Очищає INFORMATION_UNIT_DIR (крім INFORMATION_UNIT_KEEP_FILE_NAME) і
    завантажує туди файли БЧС за обрану дату з групи Signal. Очікувані збої
    (Signal не встановлено, групу не знайдено) друкуються червоним, і скрипт
    іде далі зі старою чи порожньою текою."""
    bchs_date = ask_date("Оберіть дату, за яку вивантажити файли", default=date.today())
    date_text = bchs_date.strftime("%d.%m.%Y")
    clear_information_unit_directory(INFORMATION_UNIT_DIR, INFORMATION_UNIT_KEEP_FILE_NAME)
    try:
        fetched = fetch_bchs_files(bchs_date, INFORMATION_UNIT_DIR)
    except SignalFetchError as e:
        print_red(str(e))
        return
    if not fetched:
        print_red(f"Жодного файлу БЧС за {date_text} у групі Signal не знайдено.")
        return
    print_green(f"Завантажено {len(fetched)} файл(ів) БЧС за {date_text}.")
    for subdivision in missing_subdivisions(fetched):
        print_red(f"Немає файлу БЧС підрозділу {subdivision.upper()} за {date_text}.")


def _generate_all():
    """Усі три генератори поспіль в одну теку. Самі ОБЛІК і Відомість не
    пропонується відкрити (кожен генератор друкує свій результат), а звіт
    розбіжностей - так, одразу після повідомлення свого генератора, як в
    окремому пункті меню. "Так" завершує скрипт, решта кроків - у наступному
    запуску. Файл, заблокований іншою програмою, зупиняє лише свій результат,
    а не решту. Правило категорії 100 перевіряється по ОБЛІК для виплат - лише
    в ньому є категорії-числа. Повертає [(назва, шлях)] згенерованих ОБЛІК,
    для виплат першим."""
    generated_obliks = []
    try:
        personel_output_path, _ = generate_timesheet(output_file_name=_ALL_PERSONEL_LIST_OUTPUT_FILE_NAME)
    except PermissionError as e:
        print_red(str(e))
    else:
        generated_obliks.append((os.path.basename(personel_output_path), personel_output_path))

    try:
        payments_output_path, _, mismatch_report_path = generate_payments_timesheet()
    except PermissionError as e:
        print_red(str(e))
    else:
        generated_obliks.insert(0, (os.path.basename(payments_output_path) + _PAYMENTS_SUFFIX, payments_output_path))
        if mismatch_report_path:
            _open_if_confirmed(mismatch_report_path)
        _check_category_100_in(payments_output_path)

    try:
        _statement_path, _people_count, schedule_mismatch_path = generate_rop_vop_statement(
            output_file_name=_ALL_ROP_VOP_STATEMENT_OUTPUT_FILE_NAME,
        )
    except PermissionError as e:
        print_red(str(e))
    else:
        # Користувач не бачив питання після "N розбіжність(ей) у Відомості
        # РОП/ВОП" - звіт однаково треба відкрити, Відомість же не збережено.
        if schedule_mismatch_path:
            _open_if_confirmed(schedule_mismatch_path)
    return generated_obliks


def _generate_rop_vop_statement():
    """Відомість РОП/ВОП. Якщо звірка з resources/schedule знайшла
    розбіжності, Відомість не зберігається - тоді пропонується відкрити звіт
    розбіжностей."""
    try:
        output_path, _people_count, schedule_mismatch_path = generate_rop_vop_statement()
    except PermissionError as e:
        print_red(str(e))
        return
    _open_if_confirmed(output_path if output_path is not None else schedule_mismatch_path)


def _generate_oblik(choice):
    """Один файл ОБЛІК - стройовий чи для виплат. Пропонує відкрити спершу
    "Потребує ручної перевірки" (якщо є зауваження) або сам ОБЛІК, далі звіт
    розбіжностей з БЧС (для виплат). Повертає [(назва, шлях)] згенерованого
    ОБЛІК (порожній список, якщо не вдалося)."""
    try:
        output_path, unresolved, *extra = _HANDLERS[choice]()
    except PermissionError as e:
        # Напр. ОБЛІК.xlsx відкритий в Excel. Самі нічого не закриваємо, щоб
        # не втратити незбережені зміни.
        print_red(str(e))
        return []

    is_payments = choice == GENERATE_PAYMENTS_LIST
    generated_obliks = [(os.path.basename(output_path) + (_PAYMENTS_SUFFIX if is_payments else ""), output_path)]
    _check_category_100_in(output_path)
    _open_if_confirmed(_REVIEW_FILE_NAME_BY_CHOICE[choice] if unresolved else output_path)

    mismatch_report_path = extra[0] if extra else None  # лише generate_payments_timesheet
    if mismatch_report_path:
        _open_if_confirmed(mismatch_report_path)
    if is_payments and not unresolved and not mismatch_report_path:
        _offer_export_for_money_project(output_path)
    return generated_obliks


def _offer_export_for_money_project(output_path):
    """Повністю чистий прогін "ОБЛІК для виплат" (ні зауважень, ні
    розбіжностей) можна записати ще кудись - теку користувач обирає
    діалогом, MONEY_PROJECT_RESOURCES_DIR - лише стартова тека діалогу."""
    if not ask_yes_no(f"Записати {output_path}?", default=False):
        return
    chosen_folder = pick_folder("Оберіть папку, куди записати файл", initial_dir=MONEY_PROJECT_RESOURCES_DIR)
    if not chosen_folder:
        print_red("Папку не обрано - файл нікуди не записано.")
        return
    target_path = os.path.join(chosen_folder, os.path.basename(output_path))
    try:
        export_for_money_project(output_path, target_path, MONEY_PROJECT_SHEET_NAME)
    except PermissionError as e:
        print_red(str(e))
    else:
        print_green(target_path)


def _backup_resources_and_output():
    """Ручний знімок поточних resources/ і output/ в обрану теку (напр. перед
    початком нового місяця)."""
    chosen_folder = pick_folder("Оберіть папку, куди зберегти resources та output")
    if not chosen_folder:
        print_red("Папку не обрано - resources та output нікуди не збережено.")
        return
    try:
        target_dir = backup_resources_and_output(chosen_folder, REPORT_DIR, RESOURCES_DIR, OUTPUT_DIR)
    except OSError as e:
        print_red(str(e))
        return
    if target_dir is None:
        print_red(f"У {REPORT_DIR} не знайдено жодного рапорту з розпізнаваною датою - не вдалось назвати нову підтеку.")
    else:
        print_green(target_dir)


def main():
    # Найперше питання: "Так" - лише звірка книг addons між собою, і кінець.
    if ask_yes_no("Звірити додатки (resources/addons) між собою?", default=False):
        _cross_check_addons()
        sys.exit(0)
    if ask_yes_no("Отримати файли resources/report та resources/schedule?", default=False):
        _sync_reports_from_signal()
    if ask_yes_no("Отримати файли БЧС?", default=False):
        _fetch_bchs_files()

    choice = ListPrompt(message="Що зробити?", choices=_GENERATE_CHOICES, default=GENERATE_ALL).execute()
    if choice == STOP_SCRIPT:  # ще до очищення output/
        sys.exit(0)
    clear_output_directory(OUTPUT_DIR)

    generated_obliks = []
    if choice == CHECK_CATEGORY_100:
        _pick_and_check_category_100()
    elif choice == CHECK_ADDON_MISMATCH:
        _pick_and_check_addon_mismatches()
    elif choice == GENERATE_ALL:
        generated_obliks = _generate_all()
    elif choice == GENERATE_ROP_VOP_STATEMENT:
        _generate_rop_vop_statement()
    else:
        generated_obliks = _generate_oblik(choice)

    # Ці питання - після будь-якого пункту меню, навіть якщо генерація не вдалась.
    if ask_yes_no("Здійснити звірку з addons?", default=False):
        _choose_and_check_addon_mismatches(generated_obliks)
    if ask_yes_no(f"Відкрити {OUTPUT_DIR}?", default=False):
        open_folder(OUTPUT_DIR)
    if ask_yes_no("Зберегти resources та output в іншій папці?", default=False):
        _backup_resources_and_output()


if __name__ == "__main__":
    main()
