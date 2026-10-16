"""Генерує МІНІМАЛЬНИЙ приклад resources/ОБЛІК.xlsx (кілька вигаданих рядків,
щоб було видно точну структуру) - не чіпає файл, якщо він уже існує.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Файли resources/report/*.docx (щоденні рапорти) сюди НЕ входять - це вільний
текст конкретного рапорту, приклад-заготовка тут не буде корисним; беріть
реальний шаблон рапорту вашого підрозділу."""
import os
from datetime import datetime, timedelta

from openpyxl import Workbook

RESOURCES_DIR = os.path.join(os.path.dirname(__file__), "..", "resources")
OUT_PATH = os.path.join(RESOURCES_DIR, "ОБЛІК.xlsx")

# Вигадані люди для прикладу - НІКОЛИ не вставляйте сюди реальні ПІБ.
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
    baseline = first_of_month - timedelta(days=1)  # "напередодні" - базовий стан, ніколи не перезаписується
    dates = [baseline, first_of_month, first_of_month + timedelta(days=1)]

    wb = Workbook()
    ws = wb.active
    ws.title = "Табель"
    ws.append(["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", *dates])
    for pidrozdil, posada, zvannya, pib in PEOPLE:
        ws.append([pidrozdil, posada, zvannya, pib, "РВЗ", None, None])  # лише базовий день заповнено
    for col in range(5, 5 + len(dates)):
        ws.cell(row=1, column=col).number_format = "DD.MM.YYYY"

    os.makedirs(RESOURCES_DIR, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Створено {OUT_PATH} (аркуш 'Табель', базовий день {baseline.strftime('%d.%m.%Y')} заповнено).")
    print("Замініть вигаданих людей на реальних і додайте resources/report/*.docx для наступних днів.")


if __name__ == "__main__":
    main()
