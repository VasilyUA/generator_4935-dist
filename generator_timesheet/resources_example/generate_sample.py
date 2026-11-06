"""Створює мінімальний приклад resources/ОБЛІК.xlsx - кілька вигаданих рядків,
щоб було видно структуру. Наявний файл не чіпає.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Щоденних рапортів resources/report/*.docx тут немає: це вільний текст, і
заготовка не допоможе - беріть справжній шаблон рапорту свого підрозділу."""
import os
from datetime import datetime, timedelta

from openpyxl import Workbook

RESOURCES_DIR = os.path.join(os.path.dirname(__file__), "..", "resources")
OUT_PATH = os.path.join(RESOURCES_DIR, "ОБЛІК.xlsx")

# Вигадані люди для прикладу. Справжніх ПІБ тут бути не може.
PEOPLE = [
    ("Підрозділ 1", "Командир взводу", "лейтенант", "ПЕРШИЙ Перший Першович"),
    ("Підрозділ 2", "Стрілець", "солдат", "ДРУГИЙ Другий Другович"),
]


def main():
    if os.path.exists(OUT_PATH):
        print(f"{OUT_PATH} вже існує - нічого не роблю (заберіть/перейменуйте файл, якщо хочете новий приклад).")
        return

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

    os.makedirs(RESOURCES_DIR, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Створено {OUT_PATH} (аркуш 'Табель', базовий день {baseline.strftime('%d.%m.%Y')} заповнено).")
    print("Замініть вигаданих людей на реальних і додайте resources/report/*.docx для наступних днів.")


if __name__ == "__main__":
    main()
