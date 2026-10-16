import os
import sys
from datetime import datetime
from pathlib import Path

_generator_br_and_report_for_money_DIR = Path(__file__).resolve().parent.parent
sys.path.append(str(_generator_br_and_report_for_money_DIR))


def _oblik_best_available_month():
    """(МІСЯЦЬ, РІК) - той, за який у справжньому resources/ОБЛІК.xlsx (аркуш
    ТРАВЕНЬ) РЕАЛЬНО є найбільше заповнених дат - незалежно від сьогоднішньої
    календарної дати. Реальний випадок (2026-08-19): файл мав ЛИШЕ дані за
    травень (аркуш так і лишився "ТРАВЕНЬ", нові місяці в нього ще не
    дописані) - "поточний"/"попередній" ВІДНОСНО СЬОГОДНІ (серпень/липень)
    обидва давали 0 заповнених клітинок для _oblik_populated_cell_count вище,
    тож старий запасний варіант ("Попередній") теж вів на порожній місяць, і
    constants.py (PERSONEL_LIST_COLUMNS_LETTERS, детекція колонок-дат)
    падав SystemExit ще на ІМПОРТІ - зупиняючи ВЕСЬ прогін тестів, не лише
    цього проєкту (корінний conftest.py збирає всі п'ять проєктів в ОДНІЙ
    pytest-сесії). Ця функція натомість сканує ВСІ колонки-дати аркуша й сама
    визначає, який місяць/рік у них дійсно переважає - працює однаково, який
    би місяць не був "сьогодні" і який би місяць не вела людина в файлі.
    None, якщо файл відсутній/пошкоджений чи в ньому взагалі нема заповнених
    дат жодного місяця (тоді execute() нижче падає назад на старий варіант)."""
    try:
        import openpyxl
        from utils.excel_reader import find_file_with_any_extension

        file_path = find_file_with_any_extension(os.path.join("resources", "ОБЛІК.xlsx"))
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_name = "ТРАВЕНЬ" if "ТРАВЕНЬ" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        date_columns = [(i, value) for i, value in enumerate(header) if isinstance(value, datetime)]
        if not date_columns:
            return None

        counts = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            for i, date_value in date_columns:
                if i < len(row) and row[i] is not None and row[i] != "":
                    key = (date_value.month, date_value.year)
                    counts[key] = counts.get(key, 0) + 1
        if not counts:
            return None

        (month_int, year_int), _count = max(counts.items(), key=lambda item: item[1])
        return f"{month_int:02d}", str(year_int)
    except Exception:
        return None


def pytest_sessionstart(session):
    # Скрипти проєкту (constants.py, index.py) читають resources/... відносно
    # generator_br_and_report_for_money/, тож cwd під час збору й виконання тестів має бути цією
    # директорією — незалежно від того, звідки саме запущено pytest.
    # Зроблено через pytest_sessionstart (а не на рівні модуля), щоб це
    # відбувалось ПІСЛЯ того, як pytest-cov у pytest_configure вже визначив
    # джерело покриття (--cov=generator_br_and_report_for_money у pytest.ini з кореня репозиторію,
    # де корінний conftest.py вже виставив правильний cwd).
    os.chdir(_generator_br_and_report_for_money_DIR)


import pytest
import InquirerPy.prompts.list as _list_mod
import InquirerPy.prompts.input as _input_mod


class FakePrompt:
    """Заглушка InquirerPy для тестового середовища без реальної консолі (Win32 screen buffer).

    За замовчуванням повертає default, переданий у конструктор — це те саме,
    що користувач просто натиснув Enter. Тести, яким потрібні конкретні
    "введені" значення (наприклад перевірка циклу валідації), можуть заповнити
    чергу через фікстуру inquirer_inputs.
    """

    _queue = []

    def __init__(self, *args, **kwargs):
        self._kwargs = kwargs

    def execute(self):
        if FakePrompt._queue:
            return FakePrompt._queue.pop(0)

        # Виняток саме для промпту вибору місяця (constants.py): "default" там -
        # ЗАВЖДИ "поточний місяць" (відносно реального "сьогодні"), а resources/
        # ОБЛІК.xlsx - "живий" файл, який користувач редагує паралельно з
        # роботою над проєктом (не застиглий знімок за один конкретний місяць,
        # і не обов'язково взагалі веде саме ПОТОЧНИЙ реальний місяць - див.
        # _oblik_best_available_month вище) - тож не можна просто взяти
        # "поточний"/"попередній" відносно сьогоднішньої дати як є.
        choices = self._kwargs.get("choices")
        message = self._kwargs.get("message", "")
        if choices and "місяця" in message:
            # Спершу - який місяць РЕАЛЬНО є в самому файлі (працює навіть
            # коли жоден з "Поточний"/"Попередній" відносно сьогодні туди не
            # потрапляє - див. _oblik_best_available_month вище).
            best_available = _oblik_best_available_month()
            if best_available is not None:
                return best_available[0]

            # Файл відсутній/пошкоджений/зовсім без дат - запасний варіант,
            # як і раніше: "Попередній місяць" відносно сьогодні.
            for choice in choices:
                if isinstance(choice, dict) and "Попередній" in choice.get("name", ""):
                    return choice["value"]

        return self._kwargs.get("default")


_list_mod.ListPrompt = FakePrompt
_input_mod.InputPrompt = FakePrompt


def _reassert_fake_prompt():
    # InquirerPy.prompts.list/.input - модулі СПІЛЬНІ для всього pytest-процесу,
    # а не копіюються per-проєкт (на відміну від "constants"/"helpers"/... у
    # _COLLIDING_MODULES кореневого conftest.py). report_generator (і будь-який
    # інший проєкт з власною заглушкою) так само глобально підміняє ListPrompt/
    # InputPrompt на СВІЙ клас - який саме лишиться діючим після того, як усі
    # conftest.py прогону завантажаться, залежить від порядку завантаження
    # (останній виграє), а НЕ від того, чий тест зараз збирається. Якщо
    # переможе чужий клас, constants.py (from InquirerPy.prompts.list import
    # ListPrompt - див. вище) підхопить ЙОГО, а не FakePrompt звідси - і
    # промпт місяця поверне звичайний "default" замість
    # _oblik_best_available_month(), без жодної помилки імпорту (падає вже
    # пізніше й незрозуміліше - SystemExit детекції колонок).
    _list_mod.ListPrompt = FakePrompt
    _input_mod.InputPrompt = FakePrompt


def pytest_collectstart(collector):
    # Хук, як і в кореневому conftest.py (активний лише для файлів У ВЛАСНОМУ
    # піддереві - тут це tests/), перевстановлює ці два атрибути перед КОЖНИМ
    # збором файлу цього проєкту - тож навіть якщо інший проєкт перезаписав їх
    # останнім, до ПЕРШОГО імпорту constants.py (на збиранні файлу) вони знову
    # вказують на FakePrompt звідси.
    _reassert_fake_prompt()


def pytest_runtest_setup(item):
    # Той самий self-heal, але перед ВИКОНАННЯМ кожного тесту, не лише перед
    # збором файлів: constants.py тут ІМПОРТУЄТЬСЯ ПОВТОРНО не лише на рівні
    # файлу (top-level `import constants` у тестових модулях), а й УСЕРЕДИНІ
    # окремих тестових функцій (напр. test_money_report.py робить `import
    # generators.generate_report_for_get_money`, який транзитивно тягне СВІЖИЙ
    # `from constants import (...)`, УЖЕ ПІД ЧАС ВИКОНАННЯ тесту - не збору).
    # Якщо між збором файлів ЦЬОГО проєкту й виконанням ЦЬОГО конкретного
    # тесту встиг виконатись бодай один тест ІНШОГО проєкту (звична
    # черговість при спільній pytest-сесії), report_generator/тощо міг
    # ЗНОВУ перезаписати ListPrompt/InputPrompt своєю заглушкою - лише
    # collectstart тут (як було раніше) цього вже не ловить, бо збір усіх
    # файлів завершується ЗАДОВГО до виконання будь-якого тесту.
    _reassert_fake_prompt()


@pytest.fixture
def inquirer_inputs():
    """Дозволяє заскриптувати послідовність "введених" значень для InquirerPy-промптів."""
    FakePrompt._queue = []
    yield FakePrompt._queue
    FakePrompt._queue = []
