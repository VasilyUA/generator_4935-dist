import json
import os

SIGNAL_GROUP_FOR_UAV = "Кречет 2.0"
SIGNAL_GROUP_FOR_ACTIVITY = "Діяльність"
SIGNAL_GROUP_FOR_ARTILLERY = "Basketball"
SIGNAL_GROUP_FOR_SHELLING = "Пейнтбол"
SIGNAL_GROUP_FOR_COMMISSION  = "КОМІСІЇ"
SIGNAL_GROUPS = [SIGNAL_GROUP_FOR_UAV, SIGNAL_GROUP_FOR_ACTIVITY, SIGNAL_GROUP_FOR_ARTILLERY, SIGNAL_GROUP_FOR_SHELLING, SIGNAL_GROUP_FOR_COMMISSION]

FILE_PATH_DATA_IN_FOLDER_UAV = f"signal_json/{SIGNAL_GROUP_FOR_UAV.replace(' ', '').replace('.', '')}/data.json"
FILE_PATH_DATA_IN_FOLDER_ACTIVITY = f"signal_json/{SIGNAL_GROUP_FOR_ACTIVITY.replace(' ', '').replace('.', '')}/data.json"
FILE_PATH_DATA_IN_FOLDER_ARTILLERY = f"signal_json/{SIGNAL_GROUP_FOR_ARTILLERY.replace(' ', '').replace('.', '')}/data.json"
FILE_PATH_DATA_IN_FOLDER_SHELLING = f"signal_json/{SIGNAL_GROUP_FOR_SHELLING.replace(' ', '').replace('.', '')}/data.json"
FILE_PATH_DATA_IN_FOLDER_COMMISSION = f"signal_json/{SIGNAL_GROUP_FOR_COMMISSION.replace(' ', '').replace('.', '')}/data.json"
FILES_PATHS = [FILE_PATH_DATA_IN_FOLDER_UAV, FILE_PATH_DATA_IN_FOLDER_ACTIVITY, FILE_PATH_DATA_IN_FOLDER_ARTILLERY, FILE_PATH_DATA_IN_FOLDER_SHELLING, FILE_PATH_DATA_IN_FOLDER_COMMISSION]

OUTPUT_DIR = 'output'
RESOURCES_DIR = 'resources'

FILE_PATH_STATIC_DATA = os.path.join(RESOURCES_DIR, "data.json")

PERSONEL_LIST_FILE_NAME = os.path.join(RESOURCES_DIR, "СПИСОК.xlsx")

# абсолютний шлях від розташування цього файлу, а не від cwd процесу - інакше
# завантаження при імпорті constants.py мовчки давало б FileNotFoundError при
# запуску (тестів) не з кореня final_combat_report (напр. з кореня репозиторію
# через pytest.ini)
def _resource_path(file_name):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), RESOURCES_DIR, file_name)


# unit/ambiguous_callsign_overrides/bk_synonym_crew_overrides - усі об'єднані
# в ОДИН resources/data.json (поряд зі щоденним текстовим вмістом звіту), тож
# читаємо його один раз тут, при імпорті модуля. get_data.get_static_data_json()
# пізніше читає той самий файл ЩЕ РАЗ (відносний шлях, у момент генерації
# звіту) для щоденних параграфів - окремі незалежні точки входу, тож простіше
# прочитати малий файл двічі, ніж ділити стан між ними.
with open(_resource_path("data.json"), encoding="utf-8") as _data_file:
    _data = json.load(_data_file)

_unit = _data["unit"]

PERSONEL_LIST_SHEET_NAME = _unit["PERSONEL_LIST_SHEET_NAME"]
# P, AA - П.І.Б./Позивний (пошук імені за позивним); K, N, O - Посада/звання за штатом/звання фактичне (підпис у кінці звіту)
PERSONEL_LIST_COLUMNS_LETTERS = ['P', 'AA', 'K', 'N', 'O']

# один файл без дати в імені - кожен день зберігається як окремий стовпець
# "кількість\nза {дата}", а не окремий файл (щоб не плодити копії щодня)
BK_UAV_FILE_PATH = os.path.join(RESOURCES_DIR, "БК ВБАК.xlsx")

# посади і повні назви для підпису в кінці звіту - звання і ПІБ підтягуються
# з СПИСКУ динамічно (get_signature_officer), самі назви посад в СПИСКУ фіксовані
COMMANDER_POSADA = "Командир батальйону"
CHIEF_OF_STAFF_POSADA = "Начальник штабу-заступник командира батальйону"
COMMANDER_TITLE = _unit["COMMANDER_TITLE"]
CHIEF_OF_STAFF_TITLE = _unit["CHIEF_OF_STAFF_TITLE"]

# UNITS
UNIT_BRIGADE = _unit["UNIT_BRIGADE"]
# Варіант написання з великими літерами ("ОБрМП"), який трапляється в сирих
# повідомленнях і нормалізується REPLACEMENTS до UNIT_BRIGADE - експортується
# окремо, щоб тести могли перевіряти саме цей шлях нормалізації.
UNIT_BRIGADE_MIXED_CASE = _unit["BRIGADE_UPPERCASE_VARIANTS"][0]
UNIT_BATTALION = _unit["UNIT_BATTALION"]
UNIT_COMPANY_ONE = _unit["UNIT_COMPANY_ONE"]
UNIT_COMPANY_TWO = _unit["UNIT_COMPANY_TWO"]
UNIT_COMPANY_DSHR = _unit["UNIT_COMPANY_DSHR"]
UNIT_COMPANY_RVP = _unit["UNIT_COMPANY_RVP"]
UNIT_COMPANY_ARTILLERY = _unit["UNIT_COMPANY_ARTILLERY"]
UNIT_COMPANY_UAV = _unit["UNIT_COMPANY_UAV"]
UNIT_COMPANY_UGV = _unit["UNIT_COMPANY_UGV"]
UNIT_COMPANY_RECONNAISSANCE = _unit["UNIT_COMPANY_RECONNAISSANCE"]
UNIT_COMPANY_ISV = _unit["UNIT_COMPANY_ISV"]
UNIT_COMPANY_COMMUNICATION = _unit["UNIT_COMPANY_COMMUNICATION"]
UNIT_COMPANY_EQUIPMENT_SUPPORT_PLATOON = _unit["UNIT_COMPANY_EQUIPMENT_SUPPORT_PLATOON"]
UNIT_COMPANY_SUPPORT_PLATOON = _unit["UNIT_COMPANY_SUPPORT_PLATOON"]
UNIT_COMPANY_INFIRMARY = _unit["UNIT_COMPANY_INFIRMARY"]

# один файл без дати в імені - кожен день зберігається як окремий стовпець
# "кількість\nза {дата}", а не окремий файл (щоб не плодити копії щодня)
REPORT_FILE_NAME_TEMPLATE = f"Додаток №1 ПБД_{UNIT_BATTALION} {{date}}.docx"


UNITS = [
    UNIT_COMPANY_ONE, UNIT_COMPANY_TWO, UNIT_COMPANY_DSHR, UNIT_COMPANY_RVP, UNIT_COMPANY_ARTILLERY,
    UNIT_COMPANY_UAV, UNIT_COMPANY_UGV, UNIT_COMPANY_RECONNAISSANCE, UNIT_COMPANY_ISV, UNIT_COMPANY_COMMUNICATION,
    UNIT_COMPANY_EQUIPMENT_SUPPORT_PLATOON, UNIT_COMPANY_SUPPORT_PLATOON, UNIT_COMPANY_INFIRMARY,
]

# У таблиці 5.1 підрозділи підписуються без пробілу (UNIT_COMPANY_ONE/TWO без
# пробілу), на відміну від UNITS (з пробілом), що використовується для пошуку
# в тексті повідомлень.
TABLE_5_1_ROWS = [
    UNIT_COMPANY_ONE.replace(" ", ""), UNIT_COMPANY_TWO.replace(" ", ""), UNIT_COMPANY_DSHR, UNIT_COMPANY_RVP, UNIT_COMPANY_ARTILLERY,
    UNIT_COMPANY_UAV, UNIT_COMPANY_UGV, UNIT_COMPANY_RECONNAISSANCE, UNIT_COMPANY_ISV, UNIT_COMPANY_COMMUNICATION,
    UNIT_COMPANY_EQUIPMENT_SUPPORT_PLATOON, UNIT_COMPANY_SUPPORT_PLATOON, UNIT_COMPANY_INFIRMARY,
]

# приданий підрозділ - виноситься окремим рядком під заголовком "Придані підрозділи" в табл. 5.1
UNIT_ATTACHED_TR = "тр"
ATTACHED_UNITS = [UNIT_ATTACHED_TR]

# категорії ОВТ для таблиці 5.2 "ВТРАТИ ОВТ"
OVT_CATEGORIES = [
    "Танків", "ББМ", "ГіМ", "РСЗВ", "ПТ засоби", "ППО", "АТТ", "ВКК",
    "Засоби зв’язку", "Спеціальна техніка", "БпЛА (НРК)", "Склад БК/медичних засобів",
]

# Словник замін: що міняємо -> на що міняємо
REPLACEMENTS = {
    # --- пробіли та спецсимволи ---
    "\n": " ",
    "\r": " ",
    "\t": " ",
    "\u00a0": " ",
    "\u202f": " ",
    "\u2009": " ",
    "\u200b": "",
    "\u200c": "",
    "\u200d": "",
    "\ufeff": "",
    "\\": "/",
    "*": "x",

    # --- лапки (фігурні -> прямі; прямі лапки лишаємо - вони позначають
    #     позивні, напр. ТЗ "Газда", ПВ "Шпонка"; апостроф лишаємо для імен) ---
    "\u2018": "'",
    "\u2019": "'",
    "\u201c": '"',
    "\u201d": '"',
    "\u00ab": '"',
    "\u00bb": '"',

    # --- тире та дефіси ---
    "\u2013": "-",
    "\u2014": "-",
    "\u2212": "-",

    # --- крапки та коми ---
    "⁩⁩.": ".",
    "⁩⁩,": ",",
    "⁩.": ".",
    "⁩,": ",",
    "⁩": "",
    "...": ".",
    "..": ".",
    ",,": ", ",
    ",.": ".",

    # --- пунктуація з пробілом ---
    " ,": ",",
    " :": ":",
    " ;": ";",
    " !": "!",
    " ?": "?",

    # --- граматика ---
    "Кізомиса": "Кізомис",
    "знищенно": "знищено",
    ",Софіївка": ", Софіївка",
    "Паук": "Павук",
    "Стафф": "Staff",
    "лупоніс": "лупиніс",
    "стік ф-0225": "стік уф-0225",

    # --- нормалізація підрозділів ---
    "Лупоніс 10 день": "Лупиніс 10",
    "Лупеніс 10 день": "Лупиніс 10",
    "Лупеніс": "Лупиніс",
    "Лупоніс": "Лупиніс",
    "Колібрі 10 ніч": "Колібрі 10",
}
for _brigade_variant in _unit["BRIGADE_UPPERCASE_VARIANTS"]:
    REPLACEMENTS[_brigade_variant] = UNIT_BRIGADE
REPLACEMENTS[f"{UNIT_BRIGADE} {UNIT_BATTALION}"] = f"{UNIT_BATTALION} {UNIT_BRIGADE}"
REPLACEMENTS[UNIT_BATTALION.upper()] = UNIT_BATTALION
REPLACEMENTS[UNIT_BATTALION.upper().replace(" ", "")] = UNIT_BATTALION

# Скорочення позивних, якими інколи підписуються оператори в сирих
# повідомленнях - не збігаються буквально з повним позивним у СПИСКУ
# (напр. "Пух" замість "Вінні-Пух") і тому не резолвились у ПІБ через fmt().
CALLSIGN_ALIASES = {
    "пух": "вінні-пух",
    "косий": "косой",
    "мпсяня": "масяня",
    "мурчік": "мурчик",
}

def _parse_pipe_keyed_overrides(raw):
    """Розбирає dict {"ключ1|ключ2": значення, ...} (вже завантажений із
    resources/data.json) у dict із ключем (ключ1, ключ2) - спільний формат
    для *_overrides-розділів цього файлу. Рядки, що починаються з "_" (напр.
    "_comment"), ігноруються."""
    return {
        tuple(key.split("|", 1)): value
        for key, value in raw.items()
        if not key.startswith("_")
    }


# Позивні, для яких резолв за замовчуванням (точний збіг чи CALLSIGN_ALIASES)
# дає НЕПРАВИЛЬНУ людину для конкретного екіпажу - перевіряється ПЕРШИМ, ще до
# прямого пошуку в СПИСКУ. Дані (з коментарем-описом обох випадків) винесені
# в ресурс - див. розділ "ambiguous_callsign_overrides" у resources/data.json.
# Ключ розділу - "екіпаж|позивний" (нижній регістр); тут розпаковується в
# (екіпаж, позивний).
AMBIGUOUS_CALLSIGN_OVERRIDES = _parse_pipe_keyed_overrides(_data.get("ambiguous_callsign_overrides", {}))

# Синонім з "Витрата:" (напр. "осколок"), для якого В ТАБЛИЦІ БК ВБАК.xlsx
# ОДНАКОВО збігаються десятки рядків (усі з ідентичним широким списком
# синонімів) - без цього словника resolve_bk_synonyms() обирав би серед них
# просто той, у якого найбільший ЗАЛИШОК, що не обов'язково те, що реально
# використовує саме цей екіпаж/засіб. Ключ розділу - "екіпаж|синонім" (нижній
# регістр); значення - ТОЧНА офіційна назва (колонка D). Див. розділ
# "bk_synonym_crew_overrides" у resources/data.json.
BK_SYNONYM_CREW_OVERRIDES = _parse_pipe_keyed_overrides(_data.get("bk_synonym_crew_overrides", {}))

# не окремі БК-позиції - або компоненти, що завжди йдуть в одній "Витрата:"
# РАЗОМ з основним боєприпасом і списуються як один акт (носій-платформа
# "Лупиніс"/"Вирій", детонатор ЕД-8, плата ініціації), або слово-опис
# результату ("розрив"), що регекс через "N шт" одразу після нього хибно хапає
# як назву виробу (напр. "Моа-400 (1 шт) - розрив 1 шт"). У "Витрата:" і
# зведенні ВБпАК окремим рядком не показуються (перевірка - startswith по
# назві, вже після clean_text+lower)
UAV_SPEND_COMPONENT_PREFIXES = [
    "лупиніс", "вирій", "плата ініціації", "плата інціації", "ед-8", "ед 8", "ед8", "розрив",
]

# Неофіційні написання виробу з "Витрата:", що в сирих повідомленнях завжди
# позначають ОДИН і той самий виріб з таблиці БК ВБАК (однозначно, без
# залежності від залишків чи типу БпЛА) -> офіційна назва для відображення
# (без службового префікса "Боєприпас"). Ключ - clean_text().lower() назви.
UAV_MUNITION_DISPLAY_NAMES = {
    "ко пузатий змій": 'КО 1,3 "Пузатий змій"',
}

# Окремо — заміни через regex щоб не ламати слова
REGEX_REPLACEMENTS = [
    # прибираємо ". " або "," на початку рядка або після ", "
    (r'(?:^|(?<=, ))\.\s+', ""),
    # "о." тільки як окреме слово
    (r'\bо\.(?=\s)', "о. "),
    # "н.п." лишаємо скороченням, лише нормалізуємо пробіл після нього
    # (у сирих повідомленнях трапляється "н.п.Федорівське" без пробілу)
    (r'\bн\.п\.\s*', "н.п. "),
    # так само "с." (село) - "с.Філія" -> "с. Філія"; лукахед на велику
    # літеру, щоб не займати ініціали чи інші скорочення на "с."
    (r'\bс\.(?=[А-ЯІЇЄ])', "с. "),
    # " ." якщо після пробілу крапка не є частиною числа
    (r' \.(?!\d)', "."),
    # два й більше ком підряд (можливо з пробілом між ними, напр. з сирого
    # "обрмп, ,(319-272)" - порожнє поле між двома комами) -> одна кома й
    # пробіл; ловимо ПІСЛЯ інших замін розділових знаків (вище є " ," -> ","
    # для прибирання пробілу ПЕРЕД комою, що з "обрмп, ,(...)" саме й дає
    # "обрмп,,(...)" - без цього кроку так і лишалось би дві коми поспіль)
    (r',\s*,+', ", "),
    # подвійні пробіли
    (r' {2,}', " "),
    # "Айкос" без закінчення -> "Айкоси" (\b після "с" не спрацює всередині вже
    # правильного "Айкоси", бо там за "с" йде "и" - словесний символ, без межі)
    (r'\bАйкос\b', "Айкоси"),
]