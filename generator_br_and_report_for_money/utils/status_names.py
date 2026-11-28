"""Статуси, перейменовані в табелі: назву статусу змінюють в одному місці -
колонка "Статус" таблиці "Статуси" книги налаштувань generator_timesheet.

Табель після перейменування пише в ОБЛІК.xlsx нову назву, а налаштування й
код цього проєкту (пункти й категорії рапорту, винятки невиплати) знають
статус за незмінним кодом (колонка "Код" тієї самої таблиці). Тож усе, що
читається з ОБЛІК.xlsx (utils/excel_reader.py::get_data), спершу
перекладається з назви в код - і нічого в цьому проєкті після перейменування
міняти не треба.

Книга табеля читається як файл, без імпорту коду того проєкту (пакети двох
проєктів навмисно розділені) - так само, як табель читає таблицю "Підрозділи"
книги цього проєкту. Таблиця шукається за смугою "■ Статуси" на будь-якому
аркуші (чи - книга давнішого вигляду - аркуш "Статуси"). Немає книги табеля
(машина, де генерують лише рапорти) чи в ній немає таблиці "Статуси" -
перекладати нічого, значення лишаються як є."""
import os

from utils import settings_book as sb
from utils.logging_utils import print_red

_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TIMESHEET_SETTINGS_PATH = os.path.join(_PROJECT_DIR, "..", "generator_timesheet", "resources", sb.SETTINGS_FILE_NAME)
STATUSES_SHEET = "Статуси"
_STATUS_COLUMNS = (
    sb.Column("status", "Статус", sb.NUMBER_TEXT, nonempty=True),
    sb.Column("code", "Код", sb.NUMBER_TEXT, required=False),
)


def _norm(text):
    return " ".join(str(text).split()).upper()


def read_status_codes(path):
    """{назва перейменованого статусу (нормалізована): код} з таблиці "Статуси"
    книги табеля path; статуси з назвою, що дорівнює коду, не потрапляють."""
    with sb.open_for_reading(path) as workbook:
        if not sb.has_table(workbook, None, STATUSES_SHEET):
            return {}
        rows = sb.read_rows(workbook, None, _STATUS_COLUMNS, section=STATUSES_SHEET)
    return {
        _norm(row["status"]): row["code"]
        for _row_number, row in rows
        if row["code"] and _norm(row["status"]) != _norm(row["code"])
    }


_CODES = None


def _codes():
    global _CODES
    if _CODES is None:
        _CODES = {}
        if os.path.isfile(TIMESHEET_SETTINGS_PATH):
            try:
                _CODES = read_status_codes(TIMESHEET_SETTINGS_PATH)
            except sb.SettingsError as error:
                print_red(f"Назви статусів з {TIMESHEET_SETTINGS_PATH} не прочитано ({error}) - статуси ОБЛІК.xlsx беруться як є.")
    return _CODES


def to_code(value):
    """Назва статусу, перейменованого в табелі, -> його код; будь-яке інше
    значення - як є."""
    if isinstance(value, str):
        codes = _codes()
        if codes:
            return codes.get(_norm(value), value)
    return value
