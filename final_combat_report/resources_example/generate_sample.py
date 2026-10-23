"""Генерує МІНІМАЛЬНИЙ приклад resources/СПИСОК.xlsx (кілька вигаданих
рядків, щоб було видно точну структуру) - не чіпає файл, якщо він уже існує.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Основний ЩОДЕННИЙ текстовий вміст resources/data.json (координати КСП,
положення військ...) тут НЕ генерується - він заповнюється вручну щодня, а не
структура з фіксованими колонками, тож приклад-заготовка не буде корисною.
Розділ "unit" усередині ТОГО САМОГО файлу - інша річ: constants.py читає ЦІ
ключі ще до появи будь-якого питання (без розділу чи без хоч одного ключа в
ньому скрипт не запуститься взагалі), тож цей скрипт додає розділ "unit" у
data.json (створюючи файл, якщо його ще нема), не чіпаючи решту вмісту."""
import json
import os

from openpyxl import Workbook

RESOURCES_DIR = os.path.join(os.path.dirname(__file__), "..", "resources")
OUT_PATH = os.path.join(RESOURCES_DIR, "СПИСОК.xlsx")
DATA_JSON_PATH = os.path.join(RESOURCES_DIR, "data.json")

# constants.py читає ЦІ ключі з розділу "unit" у resources/data.json ще до
# появи будь-якого питання - без розділу (чи без хоч одного ключа) скрипт не
# запуститься взагалі.
_UNIT_JSON_SAMPLE = {
    "PERSONEL_LIST_SHEET_NAME": "Аркуш1",
    "COMMANDER_TITLE": "Командир батальйону військової частини А0000",
    "CHIEF_OF_STAFF_TITLE": "Начальник штабу-заступник командира батальйону військової частини А0000",
    "UNIT_BRIGADE": "00 обрбр",
    "BRIGADE_UPPERCASE_VARIANTS": ["00 ОБрБР"],
    "UNIT_BATTALION": "0 бат",
    "UNIT_COMPANY_ONE": "Підрозділ 1", "UNIT_COMPANY_TWO": "Підрозділ 2", "UNIT_COMPANY_DSHR": "Підрозділ 3",
    "UNIT_COMPANY_RVP": "Підрозділ 4", "UNIT_COMPANY_ARTILLERY": "Підрозділ 5", "UNIT_COMPANY_UAV": "Підрозділ 6",
    "UNIT_COMPANY_UGV": "Підрозділ 7", "UNIT_COMPANY_RECONNAISSANCE": "Підрозділ 8", "UNIT_COMPANY_ISV": "Підрозділ 9",
    "UNIT_COMPANY_COMMUNICATION": "Підрозділ 10", "UNIT_COMPANY_EQUIPMENT_SUPPORT_PLATOON": "Підрозділ 11",
    "UNIT_COMPANY_SUPPORT_PLATOON": "Підрозділ 12", "UNIT_COMPANY_INFIRMARY": "Підрозділ 13",
}

# Вигадані люди для прикладу - НІКОЛИ не вставляйте сюди реальні ПІБ/позивні.
# Колонки, які реально читає код: P (ПІБ), AA (позивний), K (посада),
# N (звання за штатом), O (звання фактичне) - решта літер лишаються порожніми.
# Двоє останніх - на посадах "Командир батальйону"/"Начальник штабу-заступник
# командира батальйону" (constants.COMMANDER_POSADA/CHIEF_OF_STAFF_POSADA,
# фіксований текст) - без них підпис наприкінці звіту не сформується.
PEOPLE = [
    ("ПЕРШИЙ Перший Першович", "Сокіл", "Стрілець", "солдат", "солдат"),
    ("ДРУГИЙ Другий Другович", "Беркут", "Командир батальйону", "майор", "майор"),
    ("ТРЕТІЙ Третій Третьович", "Гриф", "Начальник штабу-заступник командира батальйону", "капітан", "капітан"),
]


def _col_index(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def main():
    os.makedirs(RESOURCES_DIR, exist_ok=True)

    data = {}
    if os.path.exists(DATA_JSON_PATH):
        with open(DATA_JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)

    if "unit" not in data:
        # мерджимо лише розділ "unit" - не чіпаємо решту вмісту (щоденні
        # параграфи звіту), яка вже може бути в реальному data.json
        data["unit"] = _UNIT_JSON_SAMPLE
        with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Додано розділ \"unit\" у {DATA_JSON_PATH} (заповніть реальними даними).")

    sheet_name = data["unit"]["PERSONEL_LIST_SHEET_NAME"]

    if os.path.exists(OUT_PATH):
        print(f"{OUT_PATH} вже існує - нічого не роблю (заберіть/перейменуйте файл, якщо хочете новий приклад).")
        return

    width = _col_index("AA") + 1
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header = [""] * width
    header[_col_index("P")] = "ПІБ"
    header[_col_index("K")] = "Посада"
    header[_col_index("N")] = "звання за штатом"
    header[_col_index("O")] = "звання фактичне"
    header[_col_index("AA")] = "позивний"
    ws.append(header)

    for pib, callsign, posada, rank_state, rank_fact in PEOPLE:
        row = [""] * width
        row[_col_index("P")] = pib
        row[_col_index("K")] = posada
        row[_col_index("N")] = rank_state
        row[_col_index("O")] = rank_fact
        row[_col_index("AA")] = callsign
        ws.append(row)

    os.makedirs(RESOURCES_DIR, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Створено {OUT_PATH}")
    print("Замініть вигаданих людей на реальних.")


if __name__ == "__main__":
    main()
