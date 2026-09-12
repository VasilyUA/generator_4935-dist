# Усе, що запитується у користувача та зчитується з файлів перед генерацією,
# зібрано тут в одному місці — в тому самому порядку, в якому воно відбувається.
import os
from dataclasses import dataclass

from InquirerPy.prompts.list import ListPrompt

from utils.logging_utils import print_red, print_green
from utils.excel_reader import (
    read_datafile, read_optional_datafile, PERSONEL_LIST_FALLBACK_SHEET_NAMES, TVO_LIST_FALLBACK_SHEET_NAMES,
    read_optional_pridani_sheet,
)
from utils.validation import get_validated_ksp_data, validate_battalion_commander_exists
from utils.folder_picker import pick_folder, pick_file
from utils.prompts import ask_yes_no as _ask_yes_no
from sync.constants_sync import sync_document_numbers_from_folder, sync_brigade_document_numbers_from_folder
from sync.changes_folder import prepare_changes_files, prompt_grounds_folders_for_changes
from run_mode_prompt import (
    ask_run_mode as _ask_run_mode,
    RUN_MODE_ALL,
    RUN_MODE_BR_ONLY,
    RUN_MODE_REPORT_ONLY,
    RUN_MODE_COMMANDER_REPORT_ONLY,
    RUN_MODE_CHANGES_ONLY,
    RUN_MODE_CHECK_ACCOUNTING,
    RUN_MODE_TIMETABLE_ONLY,
    RUN_MODE_VERIFY_REPORT_ONLY,
    RUN_MODE_CHECK_REPORT_LOG_WAR,
    RUN_MODE_CHECK_REPORT_PERIODS,
)
from checker_accounting.report_log_war_checker import (
    LOG_WAR_LOGIC_OLD, LOG_WAR_LOGIC_NEW, LOG_WAR_LOGIC_MIXED,
)
from constants import (
    PERSONEL_LIST_COLUMNS_LETTERS,
    TVO_LIST_SHEET_NAME,
    TVO_LIST_COLUMNS_LETTERS,
    PERSONEL_LIST_FILE_NAME,
    PERSONEL_LIST_SHEET_NAME,
    PRIDANI_SHEET_NAME,
    DOWRIES_LIST_FILE_NAME,
    DOWRIES_LIST_SHEET_NAME,
    DOWRIES_LIST_COLUMNS_LETTERS,
    NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
    NUMBER_OF_DOCUMENTS_BRS_SAVE,
    NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY,
    USING_STAGE_FOLDER,
    STAGE_DIR,
    MONTH,
    YEAR,
)

_CONSTANTS_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "constants.py")

# Режими, де НЕ генеруються бойові розпорядження (тож не питаємо КСП, номери БАТ/БЗ тощо).
_REPORT_ONLY_MODES = (RUN_MODE_REPORT_ONLY, RUN_MODE_COMMANDER_REPORT_ONLY, RUN_MODE_CHANGES_ONLY)

# Режими, чий вихідний документ може містити розділ "Прошу внести зміни..." за
# попередні місяці (RUN_MODE_ALL/RUN_MODE_REPORT_ONLY - додається в кінець звичайного
# рапорту; RUN_MODE_CHANGES_ONLY - єдиний вміст окремого документа) - лише для них
# має сенс шукати/зчитувати resources/changes.
_CHANGES_CAPABLE_MODES = (RUN_MODE_ALL, RUN_MODE_REPORT_ONLY, RUN_MODE_CHANGES_ONLY)

# Режими, що генерують БР (process_generate_br.py) - лише для них має сенс
# питати "чи додати приданий особовий склад (аркуш ПРИДАНІ) у БР" (money-звіти
# НЕ отримують ці рядки взагалі, незалежно від відповіді - підтверджено
# користувачем).
_BR_CAPABLE_MODES = (RUN_MODE_ALL, RUN_MODE_BR_ONLY)


@dataclass
class GeneratorInputs:
    column_names_personel: list
    rows_with_data: list
    rows_with_tvo_data: list
    rows_with_dowries_data: list
    rows_with_pridani_sheet_data: list
    city: str
    coordinates: str
    should_open_output: bool
    run_mode: str
    changes_file_pairs: list
    timetable_report_path: str = ""
    verify_report_path: str = ""
    verify_basis_folder: str = ""
    log_war_report_path: str = ""
    log_war_folder_path: str = ""
    periods_report_path: str = ""
    log_war_logic_mode: str = LOG_WAR_LOGIC_NEW


def _sync_document_numbers_if_requested(run_mode):
    """1. Чи синхронізувати номери БАТ/БЗ з папки з документами (і оновити constants.py).

    Пропускається:
    - якщо обрано лише один з рапортів (_REPORT_ONLY_MODES) - номери БАТ/БЗ там не запитуються;
    - якщо працюємо з папкою етапу — там це питання (за потреби) вже було задане в
      constants.py під час вибору етапу, а номери БАТ/БЗ вже визначені (згенеровано
      або зчитано з constants.py.bak цього етапу)."""
    if run_mode in _REPORT_ONLY_MODES:
        return

    if USING_STAGE_FOLDER:
        print_green("Працюємо з папкою етапу — номери БАТ/БЗ вже визначено (див. вище).")
        return

    if not _ask_yes_no("Зчитати обрану папку з документами (для номерів БАТ / БЗ)?", False):
        return

    chosen_folder = pick_folder("Оберіть папку з документами")
    if not chosen_folder:
        print_red("Папку не обрано — синхронізація номерів пропущена.")
        return

    # Знімок порядку ключів ДО синхронізації — рядки в constants.py йдуть у цьому ж
    # порядку, тож саме за ним зіставляємо значення назад у файл.
    original_day_keys = list(NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.keys())
    original_week_keys = list(NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.keys())
    original_save_keys = list(NUMBER_OF_DOCUMENTS_BRS_SAVE.keys())
    sync_document_numbers_from_folder(
        chosen_folder, _CONSTANTS_FILE_PATH,
        NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, original_day_keys,
        NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, original_week_keys,
        NUMBER_OF_DOCUMENTS_BRS_SAVE, original_save_keys,
    )


def _sync_brigade_document_numbers_if_requested(run_mode):
    """2. Чи синхронізувати номери БРГ (окреме питання й окрема папка від номерів
    БАТ/БЗ вище - підтверджено користувачем: реальні документи БРГ надходять
    ЗАПАКОВАНИМИ у .zip, а не звичайними .docx/.pdf з номером у назві файлу, тож
    і сканування, і сама папка - інші).

    Ті самі умови пропуску, що й для БАТ/БЗ вище (_sync_document_numbers_if_requested).
    Як і БАТ/БЗ вище - записує у constants.py одразу, без окремого підтвердження
    (підтверджено користувачем: зайве питання лише плутає - людина, для якої
    зроблено весь цей інструмент, не програміст)."""
    if run_mode in _REPORT_ONLY_MODES:
        return

    if USING_STAGE_FOLDER:
        print_green("Працюємо з папкою етапу — номери БРГ тут не синхронізуються.")
        return

    if not _ask_yes_no("Зчитати обрану папку з документами (для номерів БРГ)?", False):
        return

    chosen_folder = pick_folder("Оберіть папку з документами (.zip) для номерів БРГ")
    if not chosen_folder:
        print_red("Папку не обрано — синхронізація номерів БРГ пропущена.")
        return

    day_keys = list(NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.keys())
    save_keys = list(NUMBER_OF_DOCUMENTS_BRS_SAVE.keys())
    sync_brigade_document_numbers_from_folder(
        chosen_folder, _CONSTANTS_FILE_PATH,
        NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, day_keys,
        NUMBER_OF_DOCUMENTS_BRS_SAVE, save_keys,
    )


def _prepare_changes_if_requested(run_mode):
    """2. Пошук/звірка resources/changes (для розділу "Прошу внести зміни..." за
    попередні місяці) + підстави для кожного такого місяця - лише для режимів, чий
    вихідний документ дійсно може містити цей розділ (_CHANGES_CAPABLE_MODES).

    RUN_MODE_CHANGES_ONLY - сам вибір ЦЬОГО режиму в "Що згенерувати?" вже є "згодою"
    користувача, окремого питання тут не задається. RUN_MODE_ALL/RUN_MODE_REPORT_ONLY -
    розділ лише ДОДАЄТЬСЯ до звичайного рапорту, тож перед пошуком resources/changes
    задається окреме питання: "Ні" - рапорт генерується без розділу змін узагалі (як
    і раніше, до появи цієї функції); "Так" - далі пошук пар + питання підстав за
    кожен елігібельний місяць, як завжди."""
    if run_mode not in _CHANGES_CAPABLE_MODES:
        return []

    if run_mode != RUN_MODE_CHANGES_ONLY and not _ask_yes_no(
        "Чи потрібно вказувати в рапорті зміни до наказу за попередні місяці?", False,
    ):
        return []

    resources_dir = os.path.join(STAGE_DIR, "resources") if STAGE_DIR else "resources"
    changes_file_pairs = prepare_changes_files(resources_dir)
    if changes_file_pairs:
        prompt_grounds_folders_for_changes(changes_file_pairs)
    return changes_file_pairs


def gather_inputs(run_mode=None):
    """Збирає всі вхідні дані для генерації в порядку, в якому вони запитуються:

    run_mode - якщо вже відомий (index.py питає його ЩЕ до імпорту цього
    модуля/constants.py через run_mode_prompt.ask_run_mode() - див. коментар
    там про MODES_WITHOUT_CONSTANTS_PROMPTS), передається сюди готовим, і
    "Що зробити?" ВДРУГЕ не питається. None (типово, напр. виклик із тестів) -
    питається тут, як і раніше.
    1. що саме генерувати (звірка обліку / розпорядження / рапорт на підлеглих /
       рапорт на поправки в наказі / рапорт на командира-ТВО / табель обліку
       робочого часу / все разом);
    2. синхронізація номерів БАТ/БЗ з папки з документами (опційно, пропускається для _REPORT_ONLY_MODES);
    3. resources/changes (опційно, лише для _CHANGES_CAPABLE_MODES) - звірка prev/actual
       пар + підстави за кожен такий місяць;
    4. дані з ПРИДАНІ.xlsx (опційно — якщо файл відсутній, повертається порожній список);
    5. дані особового складу та ТВО, а також аркуш "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx
       (constants.PRIDANI_SHEET_NAME - НЕ плутати з файлом на кроці 4) - якщо
       такого аркуша немає, просто немає приданого особового складу, без
       жодного питання; якщо є І обраний режим генерує БР (_BR_CAPABLE_MODES) -
       окреме питання "чи додати" (за замовчуванням - ні);
    6. населений пункт і координати КСП (пропускається для _REPORT_ONLY_MODES - там не використовуються);
    7. чи відкривати папку output після генерації (пропускається для _REPORT_ONLY_MODES -
       там замість цього одразу відкривається сам згенерований файл рапорту).

    RUN_MODE_CHECK_ACCOUNTING - зовсім окремий інструмент (звіряє resources/check/*.xlsx,
    жодним чином не пов'язаний з ОБЛІК.xlsx/БР), тож нічого з іншого тут
    навіть не запитується і не зчитується - одразу повертаємось із порожніми/типовими
    значеннями для решти полів.

    RUN_MODE_TIMETABLE_ONLY - так само зовсім окремий інструмент: табель будується
    ЦІЛКОМ з обраного користувачем готового рапорту (.docx з resources, діалог
    вибору файлу), а НЕ з ОБЛІК.xlsx - тож ОБЛІК.xlsx/ПРИДАНІ.xlsx/ТВО тут узагалі
    не зчитуються (підтверджено користувачем - табель не повинен показувати нікого,
    кого немає в самому рапорті, навіть якщо ОБЛІК.xlsx каже інакше).

    RUN_MODE_VERIFY_REPORT_ONLY - так само окремий інструмент (обраний рапорт
    .docx через діалог вибору файлу), але, на відміну від табеля, ЦЕЙ інструмент
    САМ читає ОБЛІК.xlsx (checker_accounting.report_checker) - для звірки, а не
    для генерації - тож тут ОБЛІК.xlsx НЕ пропускається. Додатково (опційно,
    окреме питання "так/ні") можна вказати папку OUTPUT - тоді звірка ще й
    перевіряє посилання "БР"/"ЖБД" у колонці "Підстава для виплати" рапорту:
    чи існує відповідний файл у цій папці і чи згадане в ньому ПІБ людини.

    RUN_MODE_CHECK_REPORT_LOG_WAR - так само окремий інструмент: обраний рапорт
    .docx (діалог вибору файлу) + обрана папка з журналами бойових дій (ЖБД,
    по одному .docx на день - діалог вибору папки, вручну щоразу, без
    фіксованого шляху) - checker_accounting.report_log_war_checker перевіряє,
    чи ДІЙСНО підтверджено участь кожної людини за КОЖЕН день її періоду в
    наративі ЖБД (пункт "5. Рішення командира..." сьогодні -> пункт "6.
    Відомості про виконання..." -> "Хід виконання спланованих завдань" завтра),
    а не лише що якийсь файл існує.

    RUN_MODE_CHECK_REPORT_PERIODS - так само окремий інструмент, але
    НАЙЛЕГШИЙ із трьох: обраний рапорт .docx перевіряється САМ НА СЕБЕ (чи
    "ПЕРІОД" кожного рядка дає стільки днів, скільки написано в "ДНІ") - без
    ОБЛІК.xlsx чи будь-якого іншого зовнішнього джерела взагалі (підтверджено
    користувачем: для швидкої перевірки самого рапорту, коли звірка з
    ОБЛІК.xlsx чи ЖБД - зайва)."""
    if run_mode is None:
        run_mode = _ask_run_mode()

    if run_mode == RUN_MODE_CHECK_ACCOUNTING:
        return GeneratorInputs(
            column_names_personel=[], rows_with_data=[], rows_with_tvo_data=[],
            rows_with_dowries_data=[], rows_with_pridani_sheet_data=[], city="", coordinates="", should_open_output=False, run_mode=run_mode,
            changes_file_pairs=[],
        )

    if run_mode == RUN_MODE_TIMETABLE_ONLY:
        report_path = pick_file("Оберіть файл рапорту ДВ (.docx) для табеля", initial_dir="resources")
        if not report_path:
            print_red("Файл рапорту не обрано - табель не сформовано.")
        return GeneratorInputs(
            column_names_personel=[], rows_with_data=[], rows_with_tvo_data=[],
            rows_with_dowries_data=[], rows_with_pridani_sheet_data=[], city="", coordinates="", should_open_output=False, run_mode=run_mode,
            changes_file_pairs=[], timetable_report_path=report_path,
        )

    if run_mode == RUN_MODE_CHECK_REPORT_PERIODS:
        # На відміну від RUN_MODE_VERIFY_REPORT_ONLY - НЕ читає ОБЛІК.xlsx
        # взагалі (підтверджено користувачем: окремий легкий інструмент,
        # лише сам рапорт, без звірки з жодним зовнішнім джерелом).
        report_path = pick_file("Оберіть файл рапорту (.docx) для перевірки періодів/днів", initial_dir="resources")
        if not report_path:
            print_red("Файл рапорту не обрано - перевірку не виконано.")
        return GeneratorInputs(
            column_names_personel=[], rows_with_data=[], rows_with_tvo_data=[],
            rows_with_dowries_data=[], rows_with_pridani_sheet_data=[], city="", coordinates="", should_open_output=False, run_mode=run_mode,
            changes_file_pairs=[], periods_report_path=report_path,
        )

    if run_mode == RUN_MODE_VERIFY_REPORT_ONLY:
        report_path = pick_file("Оберіть файл рапорту (.docx) для звірки з ОБЛІК.xlsx", initial_dir="resources")
        if not report_path:
            print_red("Файл рапорту не обрано - звірку не виконано.")

        basis_folder = ""
        if report_path and _ask_yes_no("Перевірити підстави БР/ЖБД витягів у папці OUTPUT (існування файлу + чи згадане там ПІБ)?", False):
            basis_folder = pick_folder("Оберіть папку OUTPUT для перевірки підстав БР/ЖБД", initial_dir="output")
            if not basis_folder:
                print_red("Папку не обрано - перевірку підстав БР/ЖБД пропущено.")

        return GeneratorInputs(
            column_names_personel=[], rows_with_data=[], rows_with_tvo_data=[],
            rows_with_dowries_data=[], rows_with_pridani_sheet_data=[], city="", coordinates="", should_open_output=False, run_mode=run_mode,
            changes_file_pairs=[], verify_report_path=report_path, verify_basis_folder=basis_folder,
        )

    if run_mode == RUN_MODE_CHECK_REPORT_LOG_WAR:
        # Пояснення ПЕРЕД самими діалогами вибору (підтверджено користувачем:
        # людина, що не знає цього інструменту, не розуміла, який саме файл і
        # яку папку від неї хочуть і навіщо).
        print_green(
            "Перевірка рапорту та ЖБД: звіряє вже готовий рапорт на додаткову винагороду "
            "з журналами бойових дій - для КОЖНОЇ людини перевіряє, чи підтверджено її "
            "участь у ЖБД за КОЖЕН день, зазначений у рапорті. Спочатку - файл самого "
            "рапорту (.docx), потім - папка з файлами ЖБД ЗА ТОЙ САМИЙ місяць."
        )
        report_path = pick_file("Крок 1/3: оберіть файл рапорту (.docx) для перевірки з ЖБД", initial_dir="resources")
        if not report_path:
            print_red("Файл рапорту не обрано - перевірку не виконано.")

        log_war_folder = ""
        if report_path:
            log_war_folder = pick_folder(
                "Крок 2/3: оберіть папку з документами ЖБД (по одному .docx на день, за той самий місяць)",
                initial_dir="resources",
            )
            if not log_war_folder:
                print_red("Папку з ЖБД не обрано - перевірку не виконано.")

        logic_mode = LOG_WAR_LOGIC_NEW
        if report_path and log_war_folder:
            logic_mode = ListPrompt(
                message="Крок 3/3: за якою методикою перевіряти покриття днів у ЖБД?",
                choices=[
                    {
                        "name": "Нова - кожен день окремо: пункт 5 цього дня + пункт 6 наступного (рекомендую)",
                        "value": LOG_WAR_LOGIC_NEW,
                    },
                    {
                        "name": "Стара - лише вхід/вихід періоду (пункт 5 за день до + пункт 6 в день входу/виходу)",
                        "value": LOG_WAR_LOGIC_OLD,
                    },
                    {
                        "name": "Змішана - стара методика до автоматично визначеної межі стилю ЖБД, далі нова",
                        "value": LOG_WAR_LOGIC_MIXED,
                    },
                ],
                default=LOG_WAR_LOGIC_NEW,
            ).execute()

        return GeneratorInputs(
            column_names_personel=[], rows_with_data=[], rows_with_tvo_data=[],
            rows_with_dowries_data=[], rows_with_pridani_sheet_data=[], city="", coordinates="", should_open_output=False, run_mode=run_mode,
            changes_file_pairs=[], log_war_report_path=report_path, log_war_folder_path=log_war_folder,
            log_war_logic_mode=logic_mode,
        )

    _sync_document_numbers_if_requested(run_mode)
    _sync_brigade_document_numbers_if_requested(run_mode)
    changes_file_pairs = _prepare_changes_if_requested(run_mode)

    rows_with_dowries_data = read_optional_datafile(
        DOWRIES_LIST_FILE_NAME, DOWRIES_LIST_SHEET_NAME, DOWRIES_LIST_COLUMNS_LETTERS, date_can_be_empty=True,
    )

    rows_with_data = read_datafile(
        PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, PERSONEL_LIST_COLUMNS_LETTERS,
        extra_fallback_sheet_names=PERSONEL_LIST_FALLBACK_SHEET_NAMES,
    )
    validate_battalion_commander_exists(rows_with_data.get("rows", []), PERSONEL_LIST_FILE_NAME)

    # В/сл, хто пішов у СЗЧ ПРОТЯГОМ поточного місяця (на відміну від СЗЧ, що
    # триває з попереднього місяця - можливе повернення) - БІЛЬШЕ не
    # прибирається тут: process_generate_br.py й generate_report_for_get_money.py/
    # generate_report_for_commander_money.py тепер застосовують ЦЕ ПРАВИЛО
    # ОКРЕМО, по-різному (БР - повне виключення рядка; рапорт на додаткову
    # винагороду - лише порожнить дні ПОЗА "СЗЧ", щоб людина й далі з'являлась
    # у "NOT_PAID" - підтверджено користувачем: див. docstring-коментарі
    # _exclude_by_szch_started_this_month/_blank_non_szch_days_for_szch_started_this_month,
    # content/money_report_helpers.py). rows_with_data тут лишається ПОВНИМ.

    # sheet_can_be_missing=True — файл, збережений сусіднім проєктом generator_timesheet
    # (яке про ТВО не знає взагалі), не матиме аркуша "ТВО" - це так само означає
    # "немає ТВО" (порожній список), як і аркуш "ТВО" без жодного рядка (див.
    # utils.excel_reader.read_optional_datafile). extra_fallback_sheet_names -
    # окремого аркуша "ТВО" вже немає (підтверджено користувачем): таблицю
    # Start/End/ПОСАДА/ПІБ/ТВО тепер ведуть прямо в аркуші поточного місяця
    # (напр. "СЕРПЕНЬ") - TVO_LIST_FALLBACK_SHEET_NAMES (лише назви місяців,
    # без "Табель" - див. коментар там) дозволяє знайти її й там, коли самого
    # аркуша "ТВО" немає.
    rows_with_tvo_data = read_optional_datafile(
        PERSONEL_LIST_FILE_NAME, TVO_LIST_SHEET_NAME, TVO_LIST_COLUMNS_LETTERS, sheet_can_be_missing=True,
        extra_fallback_sheet_names=TVO_LIST_FALLBACK_SHEET_NAMES,
    )

    # Аркуш "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx (constants.PRIDANI_SHEET_NAME) - СВОЯ
    # структура (НЕ Табель-подібна - реальний файл підтвердив: "ПОСАДА" немає
    # взагалі, "ПІБ" не на 4-й позиції тощо), колонки шукаються за назвою - див.
    # docstring read_optional_pridani_sheet. Опційний: якщо такого аркуша немає
    # у файлі (чи в ньому немає потрібних колонок/дат обраного місяця) -
    # повертається порожній список, без жодної помилки і БЕЗ жодного питання
    # (питати нема про що).
    #
    # Питання "чи додати" задається ЛИШЕ якщо в аркуші дійсно Є люди з даними
    # за обраний місяць І обраний режим генерує БР (_BR_CAPABLE_MODES) -
    # підтверджено користувачем: для решти режимів (лише рапорт тощо) приданий
    # особовий склад однаково НІКУДИ не потрапляє (money-звіти не отримують ці
    # рядки взагалі), тож питання було б безглуздим.
    pridani_rows = read_optional_pridani_sheet(PERSONEL_LIST_FILE_NAME, PRIDANI_SHEET_NAME, MONTH, YEAR)
    rows_with_pridani_sheet_data = []
    if pridani_rows and run_mode in _BR_CAPABLE_MODES:
        if _ask_yes_no('Додати приданий особовий склад (аркуш "ПРИДАНІ") у бойові розпорядження?', False):
            rows_with_pridani_sheet_data = pridani_rows

    city, coordinates = ("", "") if run_mode in _REPORT_ONLY_MODES else get_validated_ksp_data()

    should_open_output = (
        False if run_mode in _REPORT_ONLY_MODES
        else _ask_yes_no("Відкрити папку output після завершення генерації?", False)
    )

    return GeneratorInputs(
        column_names_personel=rows_with_data.get("columns", []),
        rows_with_data=rows_with_data.get("rows", []),
        rows_with_tvo_data=rows_with_tvo_data,
        rows_with_dowries_data=rows_with_dowries_data,
        rows_with_pridani_sheet_data=rows_with_pridani_sheet_data,
        city=city,
        coordinates=coordinates,
        should_open_output=should_open_output,
        run_mode=run_mode,
        changes_file_pairs=changes_file_pairs,
    )
