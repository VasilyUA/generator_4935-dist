import os

# Шляхи прив'язано до розташування цього файлу (__file__), а не до поточної
# робочої директорії - той самий підхід, що й у generator_timesheet/constants.py
# (під час збирання тестів кореневим conftest.py cwd не гарантовано є цією
# текою). Сам data.json тут НЕ читається: його вміст (позначення підрозділу,
# тексти БН, кеш позицій з ПБД) передається генераторам параметром - див.
# content/data_store.py.
_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCES_DIR = os.path.join(_MODULE_DIR, "resources")
DATA_FILE_NAME = os.path.join(RESOURCES_DIR, "data.json")
MAP_ICONS_DIR = os.path.join(RESOURCES_DIR, "map_icons")
# Прототипи форматування зі зразка (generators/style_template.py) - без нього
# БН збирається з базовим форматуванням.
STYLE_FILE_NAME = os.path.join(RESOURCES_DIR, "Стиль БН.docx")
REVIEW_FILE_NAME = "Потребує_ручної_перевірки.txt"

OUTPUT_DIR = os.path.join(_MODULE_DIR, "output")
MAP_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "map")

# {battalion} - SHORT_UNIT_BATTALION з data.json великими літерами, {date} - дд.мм.рррр.
BN_FILE_NAME_TEMPLATE = "Бойовий наказ КБ {battalion} {date}.docx"
KMZ_FILE_NAME = "БН - бп кропива.kmz"
DELTA_CSV_FILE_NAME = "БП - дельта.csv"

# Бойове розпорядження старшого командира лежить прямо в resources/.
BR_FILE_GLOB = "БР*.docx"
# Теки з додатками донесень: назва починається з дати "ДД.ММ.РРРР", решта
# назви довільна - у код не потрапляє.
SOURCE_FOLDER_DATE_RE = r"^(\d{2}\.\d{2}\.\d{4})"
PBD_FILE_MARKERS = ("ПБД",)
ROP_VOP_FILE_MARKERS = ("РОП", "ВОП")

CLASSIFICATION_MARKING = "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"
