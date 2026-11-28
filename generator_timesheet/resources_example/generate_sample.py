"""Створює мінімальний приклад resources/ОБЛІК.xlsx - кілька вигаданих рядків,
щоб було видно структуру, - і resources/НАЛАШТУВАННЯ.xlsx з вигаданими
назвами й прикладом рядка в кожному списку. Наявні файли не чіпає.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Щоденних рапортів resources/report/*.docx тут немає: це вільний текст, і
заготовка не допоможе - беріть справжній шаблон рапорту свого підрозділу."""
import os
import sys
from datetime import datetime, timedelta

from openpyxl import Workbook

# Скрипт запускається напряму (python resources_example/generate_sample.py),
# тож тека проєкту - не в sys.path сама собою, а книга налаштувань пишеться
# тим самим кодом, що її й читає (utils/project_settings.py).
_PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

from utils import project_settings  # noqa: E402
from utils.settings_book import SETTINGS_FILE_NAME  # noqa: E402

RESOURCES_DIR = os.path.join(_PROJECT_DIR, "resources")

# Вигадані люди для прикладу. Справжніх ПІБ тут бути не може.
PEOPLE = [
    ("Підрозділ 1", "Командир взводу", "лейтенант", "ПЕРШИЙ Перший Першович"),
    ("Підрозділ 2", "Стрілець", "солдат", "ДРУГИЙ Другий Другович"),
]


def _month_day(day):
    return datetime.now().replace(day=day).strftime("%d.%m.%Y")


def settings_example():
    """Вигадані налаштування (та сама форма, що й словник
    project_settings.read_settings). Справжні назви підрозділу, групи Signal
    і ПІБ - лише у вашому resources/НАЛАШТУВАННЯ.xlsx, якого немає в git.
    У кожному списку - по одному прикладу з вигаданими людьми з PEOPLE, щоб
    було видно формат; заміните їх своїми."""
    first_name, second_name = PEOPLE[0][3], PEOPLE[1][3]
    return {
        "unit": {
            "SHORT_UNIT_BATTALION": "0 бат",
            "SHORT_UNIT_BRIGADE": "00 обрбр",
            "FULL_UNIT_BUT": "0 батальйону морської піхоти",
            "FULL_MILITARY_UNIT": "військової частини А0000",
        },
        "groups": {
            "SPECIAL_TASK_GROUP": "РТГР «Назва»",
            "BCHS_SIGNAL_GROUP_NAME": "Група БЧС",
            "REPORTS_SIGNAL_GROUP_NAME": "Група рапортів",
        },
        "INFORMATION_UNIT_KEEP_FILE_NAME": "Управління.xlsx",
        "company_sections": {"SEQUENCE": [pidrozdil for pidrozdil, *_rest in PEOPLE]},
        "CONST_PAYMENT_CATEGORY_100_MAX_DAYS": 27,
        "EXCLUDE_NAMES": [{"name": second_name, "days": 31, "color": True}],
        "EXCLUDE_ABSENT_FROM_INFORMATION_UNIT": [second_name],
        "EXCLUDE_FROM_PAYMENTS_REVIEW": [{"date": _month_day(1), "name": first_name}],
        "EXCLUDE_FROM_REVIEW": [{"date": _month_day(1), "name": first_name}],
        "IGNORE_LISTS": {
            "BOOK_61": [],
            "CVS": [],
            "COMBAT_TRAINING_JOURNAL": [],
            "TREATMENT_JOURNAL": [{"period": f"{_month_day(1)}-{_month_day(2)}", "name": first_name}],
        },
        "BRIGADE_MILITARY_UNITS": ["А0001"],
    }


def write_oblik(path):
    today = datetime.now()
    first_of_month = today.replace(day=1)
    baseline = first_of_month - timedelta(days=1)  # базовий стан "напередодні", генератор його не перезаписує
    dates = [baseline, first_of_month, first_of_month + timedelta(days=1)]

    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Табель"
    ws.append(["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", *dates])
    for pidrozdil, posada, zvannya, pib in PEOPLE:
        ws.append([pidrozdil, posada, zvannya, pib, "РВЗ", None, None])  # лише базовий день заповнено
    for col in range(5, 5 + len(dates)):
        ws.cell(row=1, column=col).number_format = "dd.MM.yyyy"  # див. constants.DATE_NUMBER_FORMAT
    wb.save(path)
    return baseline


def main(resources_dir=RESOURCES_DIR):
    os.makedirs(resources_dir, exist_ok=True)

    oblik_path = os.path.join(resources_dir, "ОБЛІК.xlsx")
    if os.path.exists(oblik_path):
        print(f"{oblik_path} вже існує - не чіпаю (заберіть/перейменуйте файл, якщо хочете новий приклад).")
    else:
        baseline = write_oblik(oblik_path)
        print(f"Створено {oblik_path} (аркуш 'Табель', базовий день {baseline.strftime('%d.%m.%Y')} заповнено).")
        print("Замініть вигаданих людей на реальних і додайте resources/report/*.docx для наступних днів.")

    settings_path = os.path.join(resources_dir, SETTINGS_FILE_NAME)
    if os.path.exists(settings_path):
        print(f"{settings_path} вже існує - не чіпаю.")
    else:
        project_settings.write_settings(settings_path, settings_example())
        print(f"Створено {settings_path} (вигадані назви) - впишіть свій підрозділ, групи Signal і списки людей.")


if __name__ == "__main__":
    main()
