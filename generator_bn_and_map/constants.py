import os

# Шляхи прив'язано до розташування цього файлу (__file__), а не до поточної
# робочої директорії - той самий підхід, що й у generator_timesheet/constants.py
# (під час збирання тестів кореневим conftest.py cwd не гарантовано є цією
# текою). Сам data.json тут НЕ читається: його вміст (позначення підрозділу,
# тексти БН, кеш позицій з ПБД) передається генераторам параметром - див.
# content/data_store.py.
_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCES_DIR = os.path.join(_MODULE_DIR, "resources")
# data.json - лише те, що штаб вводить руками (підрозділ, номер/час БН, назва
# КСП, коментар карти); data_auto.json - усе, що веде програма (текст-шаблон
# БН, штабні рішення, коди SIDC, координати попереднього наказу, кеш ПБД).
DATA_FILE_NAME = os.path.join(RESOURCES_DIR, "data.json")
DATA_AUTO_FILE_NAME = os.path.join(RESOURCES_DIR, "data_auto.json")
MAP_ICONS_DIR = os.path.join(RESOURCES_DIR, "map_icons")
# Прототипи форматування зі зразка (generators/style_template.py) - без нього
# БН збирається з базовим форматуванням.
STYLE_FILE_NAME = os.path.join(RESOURCES_DIR, "Стиль БН.docx")
REVIEW_FILE_NAME = "Потребує_ручної_перевірки.txt"

OUTPUT_DIR = os.path.join(_MODULE_DIR, "output")
MAP_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "map")

# {battalion} - SHORT_UNIT_BATTALION з data.json великими літерами, {date} - дд.мм.рррр.
BN_FILE_NAME_TEMPLATE = "Бойовий наказ КБ {battalion} {date}.docx"
# Файл для Кропиви (KMZ - Кропива імпортує його; власний .kme вона шифрує
# паролем при експорті, тож згенерувати .kme неможливо). Назву можна змінити
# в data_auto.json: map.KMZ_FILE_NAME.
KMZ_FILE_NAME_TEMPLATE = "ТАКТИКА {battalion}.kmz"
DELTA_CSV_FILE_NAME = "БП дельта.csv"
# Коди SIDC без справжніх значків і точки з ними (output/map/) - які значки дати.
ICONS_TODO_FILE_NAME = "Потрібні значки.txt"

# Бойове розпорядження старшого командира лежить прямо в resources/.
BR_FILE_GLOB = "БР*.docx"
# Розпорядження з розвідувального забезпечення (зони РОУ/РПР для ВБАК).
RR_FILE_GLOB = "РР*.docx"
# Позатермінові донесення з переліком перейменованих позицій (може бути кілька).
RENAME_FILE_GLOB = "*ПОЗАТЕРМІНОВ*.docx"
# Таблиця сигналів бойового управління батальйону.
SIGNAL_TABLE_GLOB = "Таблиця сигналів*.docx"
# Додатки донесення (ПБД, РОП_ВОП) - у теці resources/ПБД/ або, як раніше, у
# теці, назва якої починається з дати "ДД.ММ.РРРР" (решта назви довільна - у
# код не потрапляє). Береться найновіший ПБД - за датою в назві файлу чи теки.
PBD_FOLDER_NAMES = ("ПБД",)
SOURCE_FOLDER_DATE_RE = r"^(\d{2}\.\d{2}\.\d{4})"
PBD_FILE_MARKERS = ("ПБД",)
ROP_VOP_FILE_MARKERS = ("РОП", "ВОП")
# Останній CSV-експорт карти з Дельти (штабні об'єкти, SIDC): resources/Дельта/*.csv
# або resources/*дельта*.csv.
DELTA_FOLDER_NAME = "Дельта"
DELTA_FILE_MARKER = "дельта"

CLASSIFICATION_MARKING = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
