import json
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt

from sync.stage_folder import prompt_should_open_stage_folder, prompt_stage_selection, sync_stage_constants_backup
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

# Усі статичні JSON-ресурси проєкту (назви частини, тексти БР "ЗАВДАННЯ", шаблони
# рапорту на додаткову винагороду) об'єднано в один файл - resources/data.json,
# кожен - під власним ключем (за колишньою назвою файлу), щоб не тримати три окремі
# відкриття файлу в трьох різних місцях коду.
DATA_FILE_NAME = os.path.join("resources", "data.json")

with open(DATA_FILE_NAME, encoding="utf-8") as _data_file:
    _DATA = json.load(_data_file)

_unit = _DATA["unit"]
SHORT_UNIT_BATTALION = _unit["SHORT_UNIT_BATTALION"]
SHORT_UNIT_BRIGADE = _unit["SHORT_UNIT_BRIGADE"]
FULL_UNIT_BUT = _unit["FULL_UNIT_BUT"]
FULL_MILITARY_UNIT = _unit["FULL_MILITARY_UNIT"]
RTGR_JBD_UNIT_REFERENCE = _unit["RTGR_JBD_UNIT_REFERENCE"]

# 'ПРВД' - переведено (людину передали в інший підрозділ/частину).
CHANGE_SET = [30, 'ВП', 'ВД', 'ВЛК', 'ПРВД']

SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.xls']
# "ДАТА ЗНИКНЕННЯ"/"ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ" - дата, з якої людину внесено
# в статус "100_СПЕЦКОНТИНГЕНТ" (полон/зниклий безвісти/заручник) - потрібна
# ЛИШЕ для цього пункту (checker_accounting.report_checker,
# generators.generate_report_for_get_money), решта пунктів її ігнорують.
#
# "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ" - ЗАГАЛЬНА підстава (для БУДЬ-ЯКОГО пункту/категорії,
# як і стара єдина "ПІДСТАВИ"/"ПІДСТАВА" - обидва написання лишені як запасний
# варіант, якщо в файлі й далі є стара колонка) - див. GENERAL_PIDSTAVY_COLUMN_NAMES.
# "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"/"ПІДСТАВИ ВПБП"/"ПІДСТАВИ БПШП" - ОКРЕМІ, ВЛАСНІ
# колонки підстави для КОЖНОГО з BASIS_REQUIRED_POINTS - див.
# BASIS_REQUIRED_POINT_COLUMN_NAMES нижче - підтверджено користувачем: раніше
# всі три статуси помилково ділили ОДНУ й ту саму колонку "ПІДСТАВИ".
OPTIONAL_PERSONEL_COLUMN_NAMES = (
    "ПІДСТАВИ", "ПІДСТАВА", "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ",
    "ПІДСТАВИ СПЕЦКОНТИНГЕНТУ", "ПІДСТАВИ ВПБП", "ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП",
    "ДАТА ЗНИКНЕННЯ", "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ",
)

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

# для загального бойового розпорядження щоденні
NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {
    f'{(datetime.strptime(f"{MONTH}.{YEAR}", "%m.%Y") - relativedelta(days=1)).day}.{(datetime.strptime(DATA, "%m.%Y") - relativedelta(months=1)).strftime("%m.%Y")}': {'бат': '', 'брг': '2665'},
    f'01.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '255', 'брг': '', 'general_br_bat': '254 від 01.09.2026'},
    f'02.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '257', 'брг': '',  'general_br_bat': '254 від 01.09.2026'},
    f'03.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '259', 'брг': '', 'general_br_bat': '254 від 01.09.2026'},
    f'04.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '261', 'брг': '',  'general_br_bat': '254 від 01.09.2026'},
    f'05.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '265', 'брг': '',  'general_br_bat': '254 від 01.09.2026'},
    f'06.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '268', 'брг': '',  'general_br_bat': '254 від 01.09.2026'},
    f'07.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '272', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'08.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '275', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'09.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '277', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'10.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '279', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'11.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '282', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'12.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '286', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'13.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'14.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'15.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'16.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'17.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'18.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'19.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'20.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'21.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'22.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'23.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'24.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'25.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'26.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'27.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'28.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'29.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
    f'30.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'брг': '', 'general_br_bat': '271 від 07.09.2026'},
}

# бойових розпоряджень (щотижневі "ЗАВДАННЯ")
NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK = {
    f'{(datetime.strptime(f"{MONTH}.{YEAR}", "%m.%Y") - relativedelta(days=1)).day}.{(datetime.strptime(DATA, "%m.%Y") - relativedelta(months=1)).strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'01.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '254', 'посилання_брг': '3071 від 19.08.2026', 'посилання_бат': '239 від 26.08.2026'},
    f'02.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'03.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'04.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'05.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'06.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'07.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '271', 'посилання_брг': '3071 від 19.08.2026', 'посилання_бат': '239 від 26.08.2026'},
    f'08.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'09.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'10.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'11.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'12.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'13.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'14.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': '239 від 26.08.2026'},
    f'15.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'16.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'17.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'18.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'19.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'20.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'21.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': '239 від 26.08.2026'},
    f'22.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'23.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'24.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'25.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'26.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'27.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'28.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'29.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
    f'30.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бат': '', 'посилання_брг': '', 'посилання_бат': ''},
}

# розпоряджень з безпеки застосування військ (бат/брг)
NUMBER_OF_DOCUMENTS_BRS_SAVE = {
    f'01.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '256', 'бз_брг': ''},
    f'02.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '258', 'бз_брг': ''},
    f'03.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'04.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '262', 'бз_брг': ''},
    f'05.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '266', 'бз_брг': ''},
    f'06.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '269', 'бз_брг': ''},
    f'07.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '273', 'бз_брг': ''},
    f'08.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '276', 'бз_брг': ''},
    f'09.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '278', 'бз_брг': ''},
    f'10.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '280', 'бз_брг': ''},
    f'11.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '283', 'бз_брг': ''},
    f'12.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '287', 'бз_брг': ''},
    f'13.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'14.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'15.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'16.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'17.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'18.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'19.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'20.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'21.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'22.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'23.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'24.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'25.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'26.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'27.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'28.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'29.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
    f'30.{datetime.strptime(DATA, "%m.%Y").strftime("%m.%Y")}': {'бз_бат': '', 'бз_брг': ''},
}

MONEY_REPORT_STATIC = _DATA["data_statik_report_for_get_additional_money"]

# ЖБД батальйону
MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR = [
    {"id": "JBD_1", "start": "28.02.2026", "end": "17.03.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/1 від 28.02.2026"]},
    {"id": "JBD_2", "start": "18.03.2026", "end": "08.07.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/2 від 18.03.2026"]},
    {"id": "JBD_3", "start": "08.04.2026", "end": "03.05.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/3 від 08.04.2026"]},
    {"id": "JBD_4", "start": "04.05.2026", "end": "30.06.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/4 від 04.05.2026"]},
    {"id": "JBD_5", "start": "01.06.2026", "end": "22.06.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/5 від 01.07.2026"]},
    {"id": "JBD_6", "start": "23.06.2026", "end": "11.06.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/6 від 23.06.2026"]},
    {"id": "JBD_7", "start": "12.07.2026", "end": "28.07.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/7 від 12.07.2026"]},
    {"id": "JBD_8", "start": "29.07.2026", "end": "08.08.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/8 від 26.07.2026"]},
    {"id": "JBD_9", "start": "09.08.2026", "end": "22.08.2026", "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/9 від 09.08.2026"]},
    {"id": "JBD_10", "start": "23.08.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"ЖБД {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №359дск/10 від 23.08.2026"]},
]

# БН батальйону
MONEY_REPORT_GENERAL_REFERENCES_BN = [
    {"id": "BN_1", "start": "24.05.2026", "end": "07.07.2026", "lines": [f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №1 від 24.05.2026"]},
    {"id": "BN_2", "start": "08.07.2026", "end": "17.07.2026", "lines": [f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №1 від 08.07.2026"]},
    {"id": "BN_3", "start": "18.07.2026", "end": "25.08.2026", "lines": [f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №1 від 18.07.2026"]},
    {"id": "BN_4", "start": "26.08.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №3 від 26.08.2026"]},
]

# Загальні підстави які надає бригада
BR_HIGHT_UNIT = {
    "general": [{"start": "01.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": ["ЖБД (Справа·№2т)"]}],
    "ЖИТТЄДІЯЛЬНІСТЬ": [
        {"start": "11.05.2026", "end": "31.08.2026", "lines": ['БН 9 АК УВ (с) "Схід" від 11.05.2026 №119/1/2474т/9']},
        {"start": "12.05.2026", "end": "31.08.2026", "lines": ['БР 9 АК №119/1/2500т/9 від 12.05.2026']},
        {"start": "12.07.2026", "end": "31.08.2026", "lines": [f"БН командира {SHORT_UNIT_BRIGADE} № 4т від 12.07.2026"]},
        {"start": "13.07.2026", "end": "31.08.2026", "lines": [f"БР {SHORT_UNIT_BRIGADE} № 2400 від 13.07.2026"]},
        {"start": "03.08.2026", "end": "31.08.2026", "lines": ['БР 9 АК від 03.08.2026 №119/1/4596т/9']},
        {"start": "08.08.2026", "end": "31.08.2026", "lines": ['БР 9 АК від 08.08.2026 №119/1/4745т/9']},
        {"start": "18.08.2026", "end": "31.08.2026", "lines": [f"БН командира {SHORT_UNIT_BRIGADE} № 5т від 18.08.2026"]},
        {"start": "19.08.2026", "end": "31.08.2026", "lines": [f"БР {SHORT_UNIT_BRIGADE} № 3071 від 19.08.2026"]},
    ],
    "РТГр": [
        {"start": "03.05.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": ["БН ТГр «Грім» № 73/ВРСО/92т від 03.05.2026"]},
        {"start": "05.05.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БР {SHORT_UNIT_BRIGADE} № 1547 від 05.05.2026"]}
    ]
}

# Бригадні накази за якими здійснилась оплата для зміни в наказі
MONEY_REPORT_CHANGES_ORDER_REFERENCES = [
    {"id": "C_1", "start": "01.05.2026", "end": "31.05.2026", "lines": ["№1319 від 07.06.2026"]},
    {"id": "C_2", "start": "01.06.2026", "end": "30.06.2026", "lines": ["№1657 від 08.07.2026"]},
    {"id": "C_3", "start": "01.07.2026", "end": "31.07.2026", "lines": ["№1921 від 06.08.2026"]},
]


MONEY_REPORT_CATEGORIES = {
    30: {
        "general": BR_HIGHT_UNIT["general"],
        "ЖИТТЄДІЯЛЬНІСТЬ": {
            "grounds": BR_HIGHT_UNIT["ЖИТТЄДІЯЛЬНІСТЬ"],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "30_РТГр": {
            "grounds": BR_HIGHT_UNIT["РТГр"],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    50: {
        "general": [],
        "ПУ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": False,
            "required_basis": False,
        },
    },
    100: {
        "general": [*MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, *MONEY_REPORT_GENERAL_REFERENCES_BN],
        "БД(СЗ)": {
            "grounds": [],
            "use_brs": True,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
        "МЕДИК": {
            "grounds": [],
            "use_brs": True,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "100_РТГр": {
        "РТГр": {
            "grounds": [{"start": "01.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БР ком РТГр №1 від 05.05.2026; ЖБД РТГР {RTGR_JBD_UNIT_REFERENCE} №341дск від 05.05.2026;"]}],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": "all",
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "100_ЗРАДН": {
        "ЗРАДН": {
            "grounds": [
                {"start": "01.08.2026", "end": "11.08.2026", "lines": [f"№365дск, Том 3 від 15.06.2026", "БР ЗРАДн №223 від 01.08.2026"]},
                {"start": "12.08.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"№365дск, Том 4 від 12.08.2026", "БР ЗРАДн №236 від 15.08.2026"]}
            ],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        }
    },
    170: {
        "general": [*MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, *MONEY_REPORT_GENERAL_REFERENCES_BN],
        "170_РТГр": {
            "grounds": [{"start": "01.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БР ком РТГр №1 від 05.05.2026; ЖБД РТГР {RTGR_JBD_UNIT_REFERENCE} №341дск від 05.05.2026;"]}],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": "all",
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "БД(СЗ)": {
            "grounds": [],
            "use_brs": True,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    70: {
        "general": [*MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, *MONEY_REPORT_GENERAL_REFERENCES_BN],
        "70_РТГр": {
            "grounds": [{"start": "01.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БР·{SHORT_UNIT_BRIGADE}·№2221 від·02.07.2026", f"БР-{SHORT_UNIT_BRIGADE} №1547 від 05.05.2026", "ЖБД (Справа·№2т)"]}],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": "all",
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "БД(СЗ)": {
            "grounds": [],
            "use_brs": True,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    10: {
        "general": [],
        "ЖИТТЄДІЯЛЬНІСТЬ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": False,
        },
    },
    "NOT_PAID": {
        "СЗЧ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "Задув.|задув|задут": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "100_ВПБП": {
        "general": [],
        "ВПБП": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "100_БПШП": {
        "general": [],
        "БПШП|100_БПШП|100_ШПБП": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "100_БПШПДЛ": {
        "БПШПДЛ": {
            "grounds": [{"start": "18.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"Довідка про обставини травми (поранення, контузії, каліцтва)  №2371/2026-д5/37 від 26.02.2026;  Мед. карта стаціонарного хворого   №3094 від 18.08.2026;  Довідка військово-лікарської комісії №2026-0813-1038-3049-7  від 13.08.2026 проведено медичний огляд «Позаштатна (госпітальна) постійно діюча ВЛК КНП «Центральна міська лікарня» Гайворонської міської ради» (травма тяжка) потребує тривалого лікування протягом 60 календарних днів.  Довідка про перебування на лікуванні №4346 від 25.08.2026"]}],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        }
    },
    "100_СПЕЦКОНТИНГЕНТ": {
        "general": [],
        "полон": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],  
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "безвісті|безв|зниклі|безв.знк.|Безвісти|БЗ|100_БЗ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],  
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "інд": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        # Дні рахуються НЕ зі сканування комірок дат ОБЛІК.xlsx (їх для таких
        # людей просто не заповнюють), а напряму від дати зникнення до кінця
        # місяця генерації - content.money_report_helpers._DISAPPEARANCE_ACCRUAL_CATEGORY/
        # build_category_rows. Той самий формат запису, що й "полон"/"інд" вище -
        # логіка обчислення днів для цієї категорії живе окремо (не в grounds).
        "Очк.БЗ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    }
}

# Пункти MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES, де порожня колонка
# "ПІДСТАВИ" ОБЛІК.xlsx (build_pidstavy_extra_grounds_by_person, content/money_report_helpers.py)
# виключає людину з рапорту ЦІЛКОМ (а не лишає рядок із порожньою "Підстава для
# виплати", як для решти пунктів) - без задокументованої підстави людина в ці
# секції не потрапляє взагалі, з попередженням у термінал про кожного виключеного
# (generate_report_for_get_money._build_categories). Той самий набір пунктів так
# само визначає РЕТРОАКТИВНЕ додавання в "Прошу внести зміни..." за попередній
# місяць, коли підстава з'явилась лише в actual-файлі, а не в prev
# (content.report_changes._basis_appeared_add_rows).
#
# Керується САМЕ тут (звичайний Python-набір, а не десь у коді генераторів) -
# щоб додати/прибрати пункт з цієї поведінки, досить дописати/видалити його
# назву тут (ключ, що вже є в MONEY_REPORT_CATEGORIES чи COMMANDER_MONEY_REPORT_CATEGORIES
# вище) - жодних змін коду не потрібно; пункту, якого немає в переданому
# categories конкретного рапорту, цей набір просто не стосується.
BASIS_REQUIRED_POINTS = {"100_СПЕЦКОНТИНГЕНТ", "100_БПШП", "100_ВПБП"}

# Кожен пункт BASIS_REQUIRED_POINTS - ВЛАСНА колонка ПІДСТАВИ в ОБЛІК.xlsx (а НЕ
# спільна "ПІДСТАВИ"/"ПІДСТАВА" - раніше всі три статуси помилково ділили ОДНУ
# колонку, тож поява підстави в одного статусу могла хибно "розблокувати" зовсім
# іншого) - підтверджено користувачем. Ключі - РІВНО ті самі, що й у
# BASIS_REQUIRED_POINTS вище (жоден код тут не покладається на набір - лише на
# відповідність ключів); додаючи новий пункт у BASIS_REQUIRED_POINTS, дописуйте
# сюди й назву його колонки.
# "100_БПШП" - ДВА написання ("ПІДСТАВИ ШПБП" першим - САМЕ так називається
# колонка в реальному ОБЛІК.xlsx, "ПІДСТАВИ БПШП" - запасний варіант, якщо
# колонку колись перейменують на назву, що збігається з самим ключем пункту) -
# підтверджено користувачем на реальному випадку (людину з підставою
# 100_БПШП): помилково виключений з рапорту, бо код шукав ЛИШЕ "ПІДСТАВИ БПШП",
# а реальна колонка в файлі - "ПІДСТАВИ ШПБП".
BASIS_REQUIRED_POINT_COLUMN_NAMES = {
    "100_СПЕЦКОНТИНГЕНТ": ("ПІДСТАВИ СПЕЦКОНТИНГЕНТУ",),
    "100_ВПБП": ("ПІДСТАВИ ВПБП",),
    "100_БПШП": ("ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП"),
}

# Загальна підстава (build_pidstavy_extra_grounds_by_person, content/money_report_helpers.py) -
# додається до "Підстава для виплати" БУДЬ-ЯКОГО пункту, крім BASIS_REQUIRED_POINTS
# (ті - лише зі своєї ВЛАСНОЇ колонки, BASIS_REQUIRED_POINT_COLUMN_NAMES вище).
# Стара єдина "ПІДСТАВИ"/"ПІДСТАВА" лишена як запасний варіант (файли, де колонку
# ще не перейменували на нову назву).
GENERAL_PIDSTAVY_COLUMN_NAMES = ("ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ", "ПІДСТАВИ", "ПІДСТАВА")

# Пункт "NOT_PAID" ("Не виплачувати додаткову винагороду...") - деякі дні-статуси
# ІГНОРУЮТЬСЯ ПОВНІСТЮ (не з'являються в рапорті взагалі - ні рядком у документі,
# ні попередженням "Немає підстави для не виплати" в термінал,
# generate_report_for_get_money._log_missing_legal_basis), якщо для ТІЄЇ САМОЇ
# людини десь протягом місяця трапляється ХОЧ ОДИН із "підтверджуючих" статусів
# з відповідного набору - НЕЗАЛЕЖНО від їхнього взаємного порядку в датах.
# Керовано САМЕ ЗВІДСИ (формат {ігнорований_статус: {підтверджуючий_статус, ...}}) -
# щоб додати/прибрати пару, досить дописати/видалити запис тут, без змін коду.
#
# Реальний випадок, підтверджений користувачем: "СЗЧ" ігнорується, якщо людину
# ОФІЦІЙНО оголошено в розшуку ("РОЗП") - сам цей факт уже достатня підстава.
# Порядок появи в датах НЕ має значення (реальні випадки: СЗЧ, потім РОЗП без
# повернення; і СЗЧ, потім РОЗП, потім знову СЗЧ (позначка на межі місяця)) -
# обидва варіанти ігноруються однаково, незалежно від того, що "РОЗП" не
# завжди лишається ОСТАННІМ хронологічно статусом.
NOT_PAID_IGNORED_WHEN_CONFIRMED_BY = {
    "СЗЧ": {"РОЗП"},
}

# Та сама структура, що й MONEY_REPORT_CATEGORIES вище (ті самі точки/назви категорій -
# _build_categories працює з обома однаково), але ОКРЕМА й незалежно редагована - підстави
# для рапорту на командира/ТВО (generate_report_for_commander_money.py) можуть відрізнятись
# від тих, що застосовуються до підлеглих у головному рапорті, тож зміна одного словника
# не повинна впливати на інший. За замовчуванням - без власних статичних "grounds"/"general"
# (use_brs=True скрізь), тож підстава будується з реальних БР, у яких згадано командира/ТВО.
COMMANDER_MONEY_REPORT_CATEGORIES = {
    10: {
        "general": [],
        "ЖИТТЄДІЯЛЬНІСТЬ": {
            "grounds": [], 
            "use_brs": False, 
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": False,
        },
    },
    30: {
        "general": [],
        "ЖИТТЄДІЯЛЬНІСТЬ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    50: {
        "general": [],
        "ПУ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": False,
            "required_basis": False,
        },
    },
    100: {
        "general": [{"id": "1", "start": "01.07.2026", "end": datetime.now().strftime("%d.%m.%Y"), "lines": [f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №2217 від 03.07.2026", f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №1 від 08.07.2026", f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №114 11.07.2026", f"БР {SHORT_UNIT_BRIGADE} №2400 13.07.2026", f"БН {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №1 від 18.07.2026", f"БР {SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE} №131 19.07.2026", "БН-9·АК-УВ (с) -\"Схід\" від-11.05.2026 №119/1/2474т/9", "БР-9·АК №119/1/2500т/9-від-12.05.2026", f"БН командира {SHORT_UNIT_BRIGADE} №-2т від-15.05.2026", f"БР·{SHORT_UNIT_BRIGADE} ·№-1622 від-16.05.2026", f"ПозБД {SHORT_UNIT_BRIGADE} №242 від 18.06.2026", f"БР {SHORT_UNIT_BRIGADE}·№-2217 від-02.07.2026", f"БН командира·{SHORT_UNIT_BRIGADE} ·№-3т·від-02.07.2026", f"БН командира {SHORT_UNIT_BRIGADE}·№-4т·від 12.07.2026", f"БР {SHORT_UNIT_BRIGADE}·№-2400 від-13.07.2026", "ЖБД (Справа·№2т)"]}],
        "ОБОРОНА": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    170: {
        "general": [],
        "ОБОРОНА": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    70: {
        "general": [],
        "ОБОРОНА": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": True,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    '100_БПШП': {
        "general": [],
        "ЛІКУВАННЯ_ПО_ПОРАНЕННЮ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": True,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
    "NOT_PAID": {
        "СЗЧ": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
        "Задув.|задув|задут": {
            "grounds": [],
            "use_brs": False,
            "use_brs_from_selected_folder": False,
            "exclude_general": [],
            "default": False,
            "generate_br": False,
            "include_to_report": True,
            "required_basis": True,
        },
    },
}

# Директорія для збереження вихідних документів
OUTPUT_DIR = 'output'   
OUTPUT_DIR_BR = 'output/br'
OUTPUT_DIR_BR_SAVE = 'output/br_save'
OUTPUT_DIR_EXTRACT_BR = 'output/extractbr'
OUTPUT_DIR_COMBAT_LOG_EXTRACT_WAR = 'output/logwar'

# Робота з папкою етапу (опційно): якщо обрано — resources/output/constants.py.bak
# беруться з обраної папки-етапу замість локальних папок проєкту.
STAGE_DIR = None
USING_STAGE_FOLDER = False
STAGE_CONSTANTS_BACKUP_PATH = None

if not _SKIP_INTERACTIVE_PROMPTS and prompt_should_open_stage_folder():
    _project_dir = os.path.dirname(os.path.abspath(__file__))
    _month_folder, _stage_dir, _is_last_stage = prompt_stage_selection(_project_dir)

    if _stage_dir:
        STAGE_DIR = _stage_dir
        USING_STAGE_FOLDER = True

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
            os.path.abspath(__file__),
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

HIGHER_COMMANDER_TITLE = "Командир батальйону"
COMMANDER_TITLE = "Начальник штабу-заступник командира батальйону"

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

SEQUENCE_HIGHER_COMMANDER = ['ЗКБ', 'ЗКБзА', 'ЗКБзППП', 'ЗКБзО', 'ЗКБзТ', 'НШ-ЗКБ', 'ГС', "НГБС", "НГВП"]

# Назви рот/взводів/батарей тощо - разом з готовими текстами завдань для БР
# (PARAGRAPH_MAP/PARAGRAPH_MAP_FOR_EXTRACT) - у resources/data.json (ключ
# "company_sections"), щоб перейменування чи додавання підрозділу відбувалось
# РІВНО в одному місці (а не тут і окремо в кожному тексті) - той самий принцип,
# що й для "unit"/"br_task_order" вище.
_company_sections = _DATA["company_sections"]
SEQUENCE = _company_sections["SEQUENCE"]
PARAGRAPH_MAP = _company_sections["PARAGRAPH_MAP"]
PARAGRAPH_MAP_FOR_EXTRACT = _company_sections["PARAGRAPH_MAP_FOR_EXTRACT"]

# Ключі - ті самі, що й у PARAGRAPH_MAP - секція "упр, штаб" єдина складена,
# представлена кортежем ('упр', 'штаб'), а не рядком (generate_content_br_general/
# generate_content_br перевіряють належність ПІДРОЗДІЛ через "s in k", де k -
# кортеж саме для цієї секції).
SECTION_LISTS = {
    (tuple(key.split(", ")) if ", " in key else key): []
    for key in PARAGRAPH_MAP
}
