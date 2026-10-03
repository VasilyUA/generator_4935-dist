import os
import warnings

# Має виконатись ДО будь-якого імпорту, що може читати .xlsx (openpyxl) - constants.py
# сам відкриває ОБЛІК.xlsx (автовизначення колонок, можливо й етап/бекап) вже під
# час імпорту нижче, тож фільтр, встановлений ПІСЛЯ цих import-рядків, запізнюється
# і попередження встигає прорватись.
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

# "Що зробити?" питається ЩЕ ДО важких імпортів (user_input.py/constants.py
# нижче) - constants.py САМ ставить два питання (місяць генерації/"Відкрити
# папку проєкту?") прямо під час імпорту, а вони НЕ по темі деяких режимів
# (напр. RUN_MODE_CHECK_REPORT_LOG_WAR - підтверджено користувачем) - тож
# спершу дізнаємось режим ЦИМ окремим, легким модулем (без залежності від
# constants.py), і лише тоді вирішуємо, чи взагалі варто ставити ці два
# питання (env-прапорець нижче, зчитується в constants.py).
from run_mode_prompt import ask_run_mode, MODES_WITHOUT_CONSTANTS_PROMPTS

_run_mode = ask_run_mode()
if _run_mode in MODES_WITHOUT_CONSTANTS_PROMPTS:
    os.environ["SKIP_CONSTANTS_INTERACTIVE_PROMPTS"] = "1"

from user_input import (
    gather_inputs, RUN_MODE_ALL, RUN_MODE_BR_ONLY, RUN_MODE_REPORT_ONLY,
    RUN_MODE_COMMANDER_REPORT_ONLY, RUN_MODE_CHANGES_ONLY, RUN_MODE_CHECK_ACCOUNTING,
    RUN_MODE_TIMETABLE_ONLY, RUN_MODE_VERIFY_REPORT_ONLY, RUN_MODE_CHECK_REPORT_LOG_WAR,
    RUN_MODE_CHECK_REPORT_PERIODS,
)
from utils.output_folder import open_output_folder, open_file
from constants import OUTPUT_DIR
from process_generate_br import process_generate_br
from process_generate_report_for_get_additional_money import (
    process_generate_report_for_get_money, process_generate_report_for_commander_money, process_generate_changes_report,
)
from checker_accounting.checker import check_accounting
from checker_accounting.report_checker import check_report_against_oblik, check_report_periods
from checker_accounting.report_log_war_checker import check_report_against_log_war
from generators.generate_timetable import process_generate_timetable

inputs = gather_inputs(run_mode=_run_mode)

if inputs.run_mode == RUN_MODE_CHECK_ACCOUNTING:
    # Звірка кількох ОБЛІК*.xlsx з resources/check - окремий інструмент, повністю
    # не пов'язаний з рештою генерації (БР/рапорти) - більше нічого тут не робимо.
    open_file(check_accounting())
elif inputs.run_mode == RUN_MODE_TIMETABLE_ONLY:
    # Табель - .xlsx сітка ПОСАДА/ЗВАННЯ/ПІБ x дні, побудована з обраного файлу
    # рапорту (не з ОБЛІК.xlsx) - так само окремий інструмент, без БР/рапортів/КСП.
    # Файл не обрано (діалог скасовано) - inputs.timetable_report_path порожній,
    # gather_inputs уже повідомив про це користувачу.
    if inputs.timetable_report_path:
        open_file(process_generate_timetable(inputs.timetable_report_path))
elif inputs.run_mode == RUN_MODE_VERIFY_REPORT_ONLY:
    # Звірка обраного рапорту з ОБЛІК.xlsx (періоди/дні/ПІБ) - окремий інструмент,
    # без БР/рапортів/КСП. Файл не обрано - inputs.verify_report_path порожній,
    # gather_inputs уже повідомив про це користувачу. verify_basis_folder -
    # опційна папка OUTPUT для звірки посилань "БР"/"ЖБД" (порожня - ця частина
    # звірки просто пропускається, як і раніше).
    if inputs.verify_report_path:
        open_file(check_report_against_oblik(inputs.verify_report_path, basis_folder=inputs.verify_basis_folder or None))
elif inputs.run_mode == RUN_MODE_CHECK_REPORT_PERIODS:
    # Перевірка рапорту (лише періоди/дні) - окремий, найлегший
    # інструмент, без БР/рапортів/КСП. Файл не обрано - inputs.periods_report_path
    # порожній, gather_inputs уже повідомив про це користувачу.
    if inputs.periods_report_path:
        open_file(check_report_periods(inputs.periods_report_path))
elif inputs.run_mode == RUN_MODE_CHECK_REPORT_LOG_WAR:
    # Перевірка рапорту та ЖБД (покриття по днях) - окремий інструмент, без
    # БР/рапортів/КСП. Файл рапорту чи папку ЖБД не обрано - відповідне поле
    # порожнє, gather_inputs уже повідомив про це користувачу.
    if inputs.log_war_report_path and inputs.log_war_folder_path:
        open_file(check_report_against_log_war(
            inputs.log_war_report_path, inputs.log_war_folder_path, logic_mode=inputs.log_war_logic_mode,
        ))
else:
    should_run_br = inputs.run_mode in (RUN_MODE_ALL, RUN_MODE_BR_ONLY)
    should_run_main_report = inputs.run_mode in (RUN_MODE_ALL, RUN_MODE_REPORT_ONLY)
    # Рапорт на командира/ТВО генерується в усіх режимах, де взагалі йдеться про рапорт
    # (разом з головним - RUN_MODE_ALL/RUN_MODE_REPORT_ONLY) і в окремому режимі "лише він".
    should_run_commander_report = inputs.run_mode in (RUN_MODE_ALL, RUN_MODE_REPORT_ONLY, RUN_MODE_COMMANDER_REPORT_ONLY)

    is_generate_br_finish = True
    if should_run_br:
        is_generate_br_finish = process_generate_br(
            inputs.column_names_personel, inputs.rows_with_data,
            inputs.rows_with_tvo_data, inputs.rows_with_dowries_data, inputs.city, inputs.coordinates,
            inputs.rows_with_pridani_sheet_data,
        )
    else:
        # process_generate_br зазвичай і створює output/ - якщо його пропущено (лише рапорт),
        # створюємо папку самі, бо рапорт зберігається саме туди.
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    report_file_path = None
    commander_report_file_path = None
    changes_report_file_path = None
    if is_generate_br_finish:
        if should_run_main_report:
            report_file_path = process_generate_report_for_get_money(
                inputs.rows_with_data, inputs.rows_with_tvo_data, inputs.column_names_personel,
                inputs.changes_file_pairs,
            )
        if should_run_commander_report:
            # Окремий рапорт на додаткову винагороду для штатного командира/ТВО (від першої
            # особи - сам командир не може фігурувати в ГОЛОВНОМУ рапорті, який він же й пише).
            # changes_file_pairs - їхні ВЛАСНІ зміни за попередні місяці (якщо є) мають
            # потрапити в ЦЕЙ рапорт, а не в головний (generate_report_for_commander_money.py,
            # build_changes_entries(..., commander_and_tvo_only=True)).
            commander_report_file_path = process_generate_report_for_commander_money(
                inputs.rows_with_data, inputs.rows_with_tvo_data, inputs.column_names_personel,
                inputs.changes_file_pairs,
            )
        if inputs.run_mode == RUN_MODE_CHANGES_ONLY:
            # Окремий документ ЛИШЕ з "Прошу внести зміни..." за попередні місяці - без
            # БР, звичайного рапорту чи рапорту на командира/ТВО (жоден з них не в
            # should_run_* вище для цього режиму).
            changes_report_file_path = process_generate_changes_report(
                inputs.rows_with_data, inputs.rows_with_tvo_data, inputs.changes_file_pairs,
            )

    if inputs.run_mode == RUN_MODE_REPORT_ONLY:
        # Тут немає питання "Відкрити папку output?" - одразу відкриваємо сам файл рапорту.
        open_file(report_file_path)
    elif inputs.run_mode == RUN_MODE_COMMANDER_REPORT_ONLY:
        open_file(commander_report_file_path)
    elif inputs.run_mode == RUN_MODE_CHANGES_ONLY:
        open_file(changes_report_file_path)
    elif inputs.should_open_output:
        open_output_folder(OUTPUT_DIR)

_constants_backup_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "constants.py.bak")
if os.path.isfile(_constants_backup_path):
    os.remove(_constants_backup_path)
