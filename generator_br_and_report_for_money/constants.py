import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt

from sync.stage_folder import prompt_should_open_stage_folder, prompt_stage_selection, sync_stage_constants_backup, bootstrap_empty_stage_folder
from utils import project_settings
from utils.column_names import column_names_of, collect_column_basis_names
from utils.settings_book import SETTINGS_FILE_NAME as _SETTINGS_BASENAME
# detect_personel_list_columns_letters імпортується ЛОКАЛЬНО, нижче, у місці
# виклику - utils.excel_reader сам читає OPTIONAL_PERSONEL_COLUMN_NAMES звідси
# (constants.py), тож імпорт цієї функції тут, нагорі файлу, створив би
# циклічний імпорт (цей модуль ще не встиг би визначити свої константи, коли
# сам довантажується ЧЕРЕЗ utils.excel_reader).

FILE_LOG = 'output/errors.xlsx'
PERSONEL_LIST_FILE_NAME = os.path.join("resources", "ОБЛІК.xlsx")
PERSONEL_LIST_SHEET_NAME = 'ТРАВЕНЬ'
# ТВО - окремий аркуш у ОБЛІК.xlsx, тому читається з PERSONEL_LIST_FILE_NAME
# (user_input.py), а не з окремого файлу.
TVO_LIST_SHEET_NAME = 'ТВО'
DOWRIES_LIST_FILE_NAME = os.path.join("resources", "ПРИДАНІ.xlsx")
DOWRIES_LIST_SHEET_NAME = 'ПРИДАНІ'
# Аркуш "ПРИДАНІ" У САМОМУ ОБЛІК.xlsx (Табель-подібна структура: ПІДРОЗДІЛ/
# ПОСАДА/ЗВАННЯ/ПІБ + дні місяця з тими ж значеннями 30/70/100/170 тощо) - НЕ
# плутати з DOWRIES_LIST_SHEET_NAME вище (аркуш ОКРЕМОГО файлу ПРИДАНІ.xlsx,
# зовсім інша структура: військове звання/ПІБ/дата прибуття/дата відбуття, без
# прив'язки до конкретних днів). Ці люди мають потрапляти в БР документи й
# витяги з ЖБД (як звичайний особовий склад - process_generate_br.py), АЛЕ НЕ
# в рапорт на додаткову винагороду - підтверджено користувачем. Опційний -
# якщо аркуша немає у файлі, просто немає жодного приданого особового складу
# (utils.excel_reader.detect_optional_sheet_columns_letters).
PRIDANI_SHEET_NAME = 'ПРИДАНІ'
# Колонки з файлу "ОБЛІК.xlsx", а саме: ПІДРОЗДІЛ, ПОСАДА, звання фактичне, ПІБ, дати з 01 та по кінець місяця.
# Початкове значення до автоматичного визначення (нижче) - саме воно завжди й лишається
# чинним: якщо для обраного місяця не вдалось знайти жодної колонки з датою, скрипт
# одразу зупиняється (а не мовчки читає щось цим списком під невірний місяць), тож це
# значення ніколи фактично не використовується як "робочий" запасний варіант - лише
# заповнене з запасом (4 фіксовані + 31 день) про всяк випадок.
PERSONEL_LIST_COLUMNS_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
# Колонки в аркуші "ТВО" файлу ОБЛІК.xlsx, а саме: Підрозділ, ПОСАДА, Start, End, ЗВАННЯ, ПІБ, ТВO
TVO_LIST_COLUMNS_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

DOWRIES_LIST_COLUMNS_LETTERS = ['B', 'C', 'D', 'E', 'F', 'L']

# Усе, що налаштовується під батальйон і місяць (назви, номери документів,
# підстави, пункти й категорії рапорту, тексти БР і рапортів), - у книзі Excel
# resources/НАЛАШТУВАННЯ.xlsx, а не тут: код оновлюється з дзеркала через
# "git reset --hard" (scripts/common.ps1), і правка в цьому файлі зникла б при
# наступному запуску START.bat, а resources/ git не чіпає. utils/project_settings.py
# читає книгу й будує з неї ті самі структури, що колись були записані прямо тут
# (NUMBER_OF_DOCUMENTS_BRS_*, MONEY_REPORT_CATEGORIES ...), тож решта коду їх не
# розрізняє. Колишні resources/data.json і значення цього файлу при першому
# запуску переносяться в книгу самі (utils/settings_migration.py). Шлях до книги
# абсолютний (від теки цього файлу) - не залежить від поточної теки.
_PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE_NAME = os.path.join(_PROJECT_DIR, "resources", _SETTINGS_BASENAME)
LEGACY_DATA_FILE_NAME = os.path.join(_PROJECT_DIR, "resources", "data.json")
_SETTINGS = project_settings.load_settings(SETTINGS_FILE_NAME, LEGACY_DATA_FILE_NAME, _PROJECT_DIR)
# Підстави без кінця діють "по сьогодні" - як колись datetime.now() прямо в цьому файлі.
_FROM_SETTINGS = project_settings.build_constants_or_exit(_SETTINGS, datetime.now().strftime("%d.%m.%Y"),
                                                          SETTINGS_FILE_NAME)

SHORT_UNIT_BATTALION = _FROM_SETTINGS.SHORT_UNIT_BATTALION
SHORT_UNIT_BRIGADE = _FROM_SETTINGS.SHORT_UNIT_BRIGADE
FULL_UNIT_BUT = _FROM_SETTINGS.FULL_UNIT_BUT
FULL_MILITARY_UNIT = _FROM_SETTINGS.FULL_MILITARY_UNIT


SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.xls']
# "ДАТА ЗНИКНЕННЯ"/"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ" - дата, з якої людину внесено
# в статус "100_СПЕЦКОНТИНГЕНТ" (полон/зниклий безвісти/заручник) - потрібна
# ЛИШЕ для цього пункту (checker_accounting.report_checker,
# generators.generate_report_for_get_money), решта пунктів її ігнорують.
#
# "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"/"ПІДСТАВИ ВПБП"/"ПІДСТАВИ БПШП"/"ПІДСТАВИ 10К"/
# "ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ" - кожна з них вказана в колонці "Колонка підстав"
# своєї категорії (чи пункту) книги налаштувань - поле "column_basis_from_tracking"
# MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES (підтверджено користувачем: раніше кілька
# статусів помилково ділили ОДНУ й ту саму колонку "ПІДСТАВИ"). КОЖНА названа
# категорія ОБОВ'ЯЗКОВО має це поле - None, якщо своєї колонки немає (рішення
# користувача: жодної прихованої, спільної "загальної" колонки для категорій
# без власної більше немає - порожнє поле видно в самому об'єкті, а не
# замовчується пропуском ключа). Те саме написання МОЖЕ повторюватись у кількох
# різних категорій/пунктів (рішення користувача) - напр. "ПІДСТАВИ ДЛЯ НЕ
# ВИПЛАТИ" одразу в кількох категоріях пункту невиплати. Повний
# перелік назв колонок, які читаються з файлу, - OPTIONAL_PERSONEL_COLUMN_NAMES,
# визначений нижче, після обох словників категорій (щоб не переписувати його
# вручну).

# Деякі режими (напр. "Перевірка рапорту та ЖБД" - run_mode_prompt.
# MODES_WITHOUT_CONSTANTS_PROMPTS) самі визначають потрібні дати з обраного
# файлу/папки і НІДЕ не використовують MONTH/"папку етапу" - index.py ставить
# цей прапорець (env, ДО імпорту цього файлу) саме для них, підтверджено
# користувачем: питання місяця й "Відкрити папку проєкту?" були геть не по
# темі обраного інструменту.
_SKIP_INTERACTIVE_PROMPTS = os.environ.get("SKIP_CONSTANTS_INTERACTIVE_PROMPTS") == "1"

# місяць
_current_month = datetime.now().strftime("%m")
_previous_month = (datetime.now() - relativedelta(months=1)).strftime("%m")

if _SKIP_INTERACTIVE_PROMPTS:
    MONTH = _current_month
else:
    _month_choice = ListPrompt(
            message="Введіть номер місяця за яким здійснюється генерація на приклад 01:",
            choices=[
                {"name": f"Поточний місяць ({_current_month})", "value": _current_month},
                {"name": f"Попередній місяць ({_previous_month})", "value": _previous_month},
                {"name": "Ввести вручну", "value": None},
            ],
            default=_current_month,
        ).execute()

    MONTH = _month_choice if _month_choice is not None else InputPrompt(
            message="Введіть номер місяця (напр. 01):",
            validate=lambda x: x.isdigit() and len(x) == 2,
            invalid_message="❌ Має бути двозначне число!"
        ).execute()
YEAR = datetime.now().strftime("%Y")
DATA = f'{MONTH}.{YEAR}'

# Номери документів за днями місяця (аркуш "Номери документів" книги
# налаштувань) - ключі "дд.мм.рррр" обраного вище місяця; рядок "попередній" -
# останній день попереднього місяця. Питання
# "Зчитати обрану папку з документами?" заповнює їх і записує назад у книгу
# (sync/constants_sync.py).
# - NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY: {'бат', 'брг', 'general_br_bat'} - для загального бойового розпорядження щоденні;
# - NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK: {'бат', 'посилання_брг', 'посилання_бат'} - бойові розпорядження (щотижневі "ЗАВДАННЯ");
# - NUMBER_OF_DOCUMENTS_BRS_SAVE: {'бз_бат', 'бз_брг'} - розпорядження з безпеки застосування військ (бат/брг).
NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, NUMBER_OF_DOCUMENTS_BRS_SAVE = (
    project_settings.build_number_dicts(_SETTINGS["numbers"], int(MONTH), int(YEAR))
)

# Тексти рапорту на додаткову винагороду (аркуші "Тексти рапорту" і "Тексти пунктів").
MONEY_REPORT_STATIC = _FROM_SETTINGS.MONEY_REPORT_STATIC

# Підстави - таблиці "ЖБД", "БН", "Підстави вищого штабу" й "Інші підстави"
# аркуша "Підстави" книги налаштувань, записи {"start", "end", "lines"[, "id"]}.
# ЖБД батальйону (таблиця "ЖБД", група "ЖБД")
MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR = _FROM_SETTINGS.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR

# БН батальйону (таблиця "БН", група "БН")
MONEY_REPORT_GENERAL_REFERENCES_BN = _FROM_SETTINGS.MONEY_REPORT_GENERAL_REFERENCES_BN

# Загальні підстави які надає бригада: {група: [записи]} - рядки аркуша
# "Підстави вищого штабу"; звірка рапорту з файлами їх не перевіряє
# (checker_accounting/report_checker.py).
BR_HIGHT_UNIT = _FROM_SETTINGS.BR_HIGHT_UNIT

# Бригадні накази за якими здійснилась оплата для зміни в наказі (таблиця "Накази виплат")
MONEY_REPORT_CHANGES_ORDER_REFERENCES = _FROM_SETTINGS.MONEY_REPORT_CHANGES_ORDER_REFERENCES

# Пункти рапорту (аркуш "Пункти рапорту" книги налаштувань: рядок "весь пункт" і під ним категорії):
# {пункт: {"general": [записи підстав для всіх категорій пункту],
#          ["column_basis_from_tracking": колонка підстав, спільна для пункту,]
#          категорія: {"grounds", "use_brs", "use_brs_from_selected_folder",
#                      "exclude_general", "default", "generate_br",
#                      "include_to_report", "required_basis",
#                      "column_basis_from_tracking"}}}
# Пункт - число (категорія виплати в колонках дат ОБЛІК.xlsx) або назва
# ("NOT_PAID", "100_ВПБП" ...). Порядок пунктів і категорій - як у книзі:
# перший збіг статусу перемагає (resolve_status_category).
MONEY_REPORT_CATEGORIES = _FROM_SETTINGS.MONEY_REPORT_CATEGORIES

# Статус невиплати не враховується, якщо в місяці в людини є й "підтверджуючий"
# статус: {статус: {підтверджуючі статуси}} - таблиця "Винятки невиплати" книги
# налаштувань (початкове значення й пояснення - utils/project_settings.py).
NOT_PAID_IGNORED_WHEN_CONFIRMED_BY = _FROM_SETTINGS.NOT_PAID_IGNORED_WHEN_CONFIRMED_BY

# Та сама структура, що й MONEY_REPORT_CATEGORIES вище (ті самі точки/назви категорій -
# _build_categories працює з обома однаково), але ОКРЕМА й незалежно редагована (аркуш
# "Пункти рапорту командира") - підстави для рапорту на
# командира/ТВО (generate_report_for_commander_money.py) можуть відрізнятись від тих, що
# застосовуються до підлеглих у головному рапорті, тож зміна одного не повинна впливати на
# інший.
COMMANDER_MONEY_REPORT_CATEGORIES = _FROM_SETTINGS.COMMANDER_MONEY_REPORT_CATEGORIES

# Назви колонок, які detect_personel_list_columns_letters читає з ОБЛІК.xlsx ЗА ЗАГОЛОВКОМ.
# Будується з усіх column_basis_from_tracking, дописаних у категоріях
# MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES вище (collect_column_basis_names,
# utils/column_names.py) - тож досить дописати назву колонки в "Колонка підстав" потрібної
# категорії книги налаштувань, без ще одного реєстру тут.
OPTIONAL_PERSONEL_COLUMN_NAMES = tuple(dict.fromkeys(
    collect_column_basis_names(MONEY_REPORT_CATEGORIES)
    + collect_column_basis_names(COMMANDER_MONEY_REPORT_CATEGORIES)
    + column_names_of(("ДАТА ЗНИКНЕННЯ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"))
))

# Директорія для збереження вихідних документів
OUTPUT_DIR = 'output'   
OUTPUT_DIR_BR = 'output/br'
OUTPUT_DIR_BR_SAVE = 'output/br_save'
OUTPUT_DIR_EXTRACT_BR = 'output/extractbr'
OUTPUT_DIR_COMBAT_LOG_EXTRACT_WAR = 'output/logwar'

# Робота з папкою етапу (опційно): якщо обрано — resources/output і книга номерів
# документів етапу (НОМЕРИ ДОКУМЕНТІВ.xlsx, sync/stage_folder.py) беруться з обраної
# папки-етапу замість локальних папок проєкту.
STAGE_DIR = None
USING_STAGE_FOLDER = False
STAGE_CONSTANTS_BACKUP_PATH = None

if not _SKIP_INTERACTIVE_PROMPTS and prompt_should_open_stage_folder():
    _project_dir = os.path.dirname(os.path.abspath(__file__))
    _month_folder, _stage_dir, _is_last_stage = prompt_stage_selection(_project_dir)

    if _stage_dir:
        STAGE_DIR = _stage_dir
        USING_STAGE_FOLDER = True

        bootstrap_empty_stage_folder(STAGE_DIR, _project_dir)

        _stage_resources_dir = os.path.join(STAGE_DIR, "resources")
        PERSONEL_LIST_FILE_NAME = os.path.join(_stage_resources_dir, "ОБЛІК.xlsx")
        DOWRIES_LIST_FILE_NAME = os.path.join(_stage_resources_dir, "ПРИДАНІ.xlsx")

        _stage_output_dir = os.path.join(STAGE_DIR, "output")
        FILE_LOG = os.path.join(_stage_output_dir, "errors.xlsx")
        OUTPUT_DIR = _stage_output_dir
        OUTPUT_DIR_BR = os.path.join(_stage_output_dir, "br")
        OUTPUT_DIR_BR_SAVE = os.path.join(_stage_output_dir, "br_save")
        OUTPUT_DIR_EXTRACT_BR = os.path.join(_stage_output_dir, "extractbr")
        OUTPUT_DIR_COMBAT_LOG_EXTRACT_WAR = os.path.join(_stage_output_dir, "logwar")

        STAGE_CONSTANTS_BACKUP_PATH = sync_stage_constants_backup(
            STAGE_DIR, MONTH, _month_folder, _is_last_stage,
            SETTINGS_FILE_NAME,
            list(NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY.keys()), list(NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK.keys()), list(NUMBER_OF_DOCUMENTS_BRS_SAVE.keys()),
            NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY, NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK, NUMBER_OF_DOCUMENTS_BRS_SAVE,
        )


# Автоматично перевизначаємо PERSONEL_LIST_COLUMNS_LETTERS на основі реальних колонок
# у файлі ОБЛІК.xlsx (ПІДРОЗДІЛ, ПОСАДА, звання фактичне, ПІБ + дати обраного місяця),
# щоб не тримати список літер вручну. Якщо для обраного місяця немає жодної колонки з
# датою (напр. за поточний місяць ще не почали вести облік) - НЕ продовжуємо мовчки з
# резервним списком вище (це могло б непомітно зчитати чужі дані під невірним місяцем,
# а СТАТУС/БР/рапорт далі так само орієнтуються на MONTH/YEAR) - натомість одразу
# зупиняємось із чіткою підказкою перезапустити скрипт і обрати місяць, за який дані
# в файлі дійсно є (напр. "Попередній місяць").
#
# Для режимів із _SKIP_INTERACTIVE_PROMPTS (напр. "Перевірка рапорту та ЖБД")
# ЦЕ ОБЧИСЛЕННЯ теж пропускається (не лише самі питання вище) - PERSONEL_LIST_
# COLUMNS_LETTERS для них ніде не використовується (ОБЛІК.xlsx вони взагалі не
# читають), тож MONTH тут не обов'язково відповідає реальним даним файлу -
# інакше довелось би зупинятись (SystemExit) через МІСЯЦЬ, який цьому
# інструменту геть не потрібен (виявлено користувачем на реальному запуску).
if not _SKIP_INTERACTIVE_PROMPTS:
    try:
        from utils.excel_reader import detect_personel_list_columns_letters
        PERSONEL_LIST_COLUMNS_LETTERS = detect_personel_list_columns_letters(
            PERSONEL_LIST_FILE_NAME, PERSONEL_LIST_SHEET_NAME, MONTH, YEAR
        )
    except Exception as _detect_error:
        raise SystemExit(
            f"❌ Не вдалося автоматично визначити колонки {PERSONEL_LIST_FILE_NAME} за {MONTH}.{YEAR} ({_detect_error}). "
            f"Перезапустіть скрипт і оберіть місяць, за який в файлі дійсно є дані (напр. 'Попередній місяць')."
        ) from _detect_error

# Посади (аркуш "Загальні" книги налаштувань) - точно як у колонці ПОСАДА ОБЛІК.xlsx.
HIGHER_COMMANDER_TITLE = _FROM_SETTINGS.HIGHER_COMMANDER_TITLE
COMMANDER_TITLE = _FROM_SETTINGS.COMMANDER_TITLE

# Відмінки для посад
COMMANDER_POSITIONS = {
    "Командир взводу": {
        "родовий": "Командира взводу",
        "давальний": "Командиру взводу"
    },
    "Командир роти": {
        "родовий": "Командира роти",
        "давальний": "Командиру роти"
    },
    "Командир батареї": {
        "родовий": "Командира батареї",
        "давальний": "Командиру батареї"
    },
    "Начальник медичного пункту": {
        "родовий": "Начальника медичного пункту",
        "давальний": "Начальнику медичного пункту"
    },
    "Командир батальйону": {
        "родовий": "Командира батальйону",
        "давальний": "Командиру батальйону"
    },
    "Начальник штабу-заступник командира батальйону": {
        "родовий": "Начальника штабу-заступника командира батальйону",
        "давальний": "Начальнику штабу-заступнику командира батальйону"
    },
    "Офіцер": {
        "родовий": "Офіцера",
        "давальний": "Офіцеру"
    }
}



COMMANDER_SHORT_POSITION = {
    'Заступник командира батальйону': 'ЗКБ',
    'Заступник командира батальйону з артилерії': 'ЗКБзА',
    'Заступник командира батальйону з психологічної підтримки персоналу': 'ЗКБзППП',
    'Заступник командира батальйону з озброєння': 'ЗКБзО',
    'Заступник командира батальйону з тилу': 'ЗКБзТ',
    'Начальник штабу-заступник командира батальйону': 'НШ-ЗКБ',
    'Головний сержант': 'ГС',
    'Офіцер': "Офіцеру",
    'Начальник групи гбс': "НГБС",
    'Начальник групи говп': "НГВП",
}

# Адресати БР після "КОМАНДИРУ ..." (аркуш "Загальні", SEQUENCE_HIGHER_COMMANDER).
SEQUENCE_HIGHER_COMMANDER = _FROM_SETTINGS.SEQUENCE_HIGHER_COMMANDER

# Назви рот/взводів/батарей тощо разом з текстами їхніх завдань у БР і витягу з
# ЖБД - таблиці "Підрозділи" й "Тексти підрозділів" книги налаштувань, щоб
# перейменування чи додавання підрозділу відбувалось РІВНО в одному місці.
# SEQUENCE - підрозділи з "У переліку підрозділів" = Так; PARAGRAPH_MAP/
# PARAGRAPH_MAP_FOR_EXTRACT - {ключ підрозділу (без пробілів, нижній регістр):
# текст} (utils/project_settings.py::section_key).
SEQUENCE = _FROM_SETTINGS.SEQUENCE
PARAGRAPH_MAP = _FROM_SETTINGS.PARAGRAPH_MAP
PARAGRAPH_MAP_FOR_EXTRACT = _FROM_SETTINGS.PARAGRAPH_MAP_FOR_EXTRACT

# Тексти щотижневого бойового розпорядження "ЗАВДАННЯ" (таблиці "Тексти ЗАВДАННЯ"
# і "Варіанти ЗАВДАННЯ" аркуша "БР і ЗАВДАННЯ") - content/br_task_order_text.py.
BR_TASK_ORDER_TEXTS = _FROM_SETTINGS.BR_TASK_ORDER_TEXTS
TASK_ORDER_VARIANTS = _FROM_SETTINGS.TASK_ORDER_VARIANTS

# Ключі - ті самі, що й у PARAGRAPH_MAP - секція "упр, штаб" єдина складена,
# представлена кортежем ('упр', 'штаб'), а не рядком (generate_content_br_general/
# generate_content_br перевіряють належність ПІДРОЗДІЛ через "s in k", де k -
# кортеж саме для цієї секції).
SECTION_LISTS = {
    (tuple(key.split(", ")) if ", " in key else key): []
    for key in PARAGRAPH_MAP
}
