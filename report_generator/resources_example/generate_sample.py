"""Генерує МІНІМАЛЬНИЙ приклад resources/СПИСОК.xlsm + resources/ЗАДАЧІ.xlsm +
mass_movement.xlsx (у КОРЕНІ проєкту, не в resources/) - кілька вигаданих
рядків, щоб було видно точну структуру. Не чіпає файли, які вже існують.
Запуск: python resources_example/generate_sample.py (з кореня проєкту)."""
import json
import os

from openpyxl import Workbook

ROOT_DIR = os.path.join(os.path.dirname(__file__), "..")
RESOURCES_DIR = os.path.join(ROOT_DIR, "resources")
DATA_JSON_PATH = os.path.join(RESOURCES_DIR, "data.json")
SPISOK_PATH = os.path.join(RESOURCES_DIR, "СПИСОК.xlsm")
ZADACHI_PATH = os.path.join(RESOURCES_DIR, "ЗАДАЧІ.xlsm")
MASS_MOVEMENT_PATH = os.path.join(ROOT_DIR, "mass_movement.xlsx")

# Вигадані люди для прикладу (№ посади - наскрізний ID, той самий у обох файлах)
# - НІКОЛИ не вставляйте сюди реальні ПІБ.
PEOPLE = [
    # №, підрозділ, ВОС, звання за штатом, звання фактичне, ПІБ (називний)
    (1, "Підрозділ 1", "0000", "солдат", "солдат", "ПЕРШИЙ Перший Першович"),
    (2, "управління", "0000", "майор", "майор", "ДРУГИЙ Другий Другович"),
]


def _col_index(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _row(width, values_by_letter):
    row = [""] * width
    for letter, value in values_by_letter.items():
        row[_col_index(letter)] = value
    return row


def _write_if_missing(path, build_fn, label=""):
    if os.path.exists(path):
        print(f"{path} вже існує - пропускаю.")
        return
    build_fn()
    print(f"Створено {path}{label}")


def main():
    os.makedirs(RESOURCES_DIR, exist_ok=True)

    if not os.path.exists(DATA_JSON_PATH):
        with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "unit": {
                    "PERSONEL_LIST_SHEET_NAME": "Штат",
                    "FULL_MILITARY_UNIT": "військової частини А0000",
                }
            }, f, ensure_ascii=False, indent=2)
        print(f"Створено {DATA_JSON_PATH} (заповніть реальними даними).")
    sheet_name = json.load(open(DATA_JSON_PATH, encoding="utf-8"))["unit"]["PERSONEL_LIST_SHEET_NAME"]

    def build_spisok():
        width = _col_index("U") + 1
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(_row(width, {
            "N": "підрозділ повністю", "O": "№", "R": "ВОС",
            "S": "звання за штатом", "T": "звання фактичне", "U": "ПІБ",
        }))
        for pos_id, pidrozdil, vos, rank_state, rank_fact, pib in PEOPLE:
            ws.append(_row(width, {
                "N": pidrozdil, "O": pos_id, "R": vos, "S": rank_state, "T": rank_fact, "U": pib,
            }))
        wb.save(SPISOK_PATH)

    def build_zadachi():
        wb = Workbook()
        ws = wb.active
        ws.title = "Аркуш1"
        ws.append([
            "№ посади", "Підрозділ", "Посада називний", "посада називний", "ПІБ називний",
            "Посада родовий", "посада родовий", "ПІБ родовий",
            "Посада давальний", "посада давальний", "ПІБ давальний",
        ])
        for pos_id, pidrozdil, _vos, _rs, _rf, pib in PEOPLE:
            ws.append([pos_id, pidrozdil, "Стрілець", "стрільця", pib, "стрільця", "стрільця", pib, "стрільцю", "стрільцю", pib])
        wb.save(ZADACHI_PATH)

    def build_mass_movement():
        wb = Workbook()
        ws = wb.active
        ws.title = "Аркуш1"
        ws.append([
            "НОМЕР ПОСАДИ З ЯКОЇ ПЕРЕМІЩУЄТЬСЯ ОСОБА",
            "НОМЕР ПОСАДИ НА ЯКУ ПЕРЕМІЩУЄТЬСЯ ОСОБА",
            "КОМАНДИР ВЗВОДУ ЧИ РОТИ АБО ЇХ ТВО ЯКІ КЛОПОЧУТЬ КОМАНДИРУ БАТАЛЬЙОНУ",
            "КОМАНДИР БАТАЛЬЙОНУ АБО ЙОГО ТВО",
        ])
        ws.append([1, 2, 1, 2])  # приклад: переміщення з посади 1 на посаду 2
        wb.save(MASS_MOVEMENT_PATH)

    _write_if_missing(SPISOK_PATH, build_spisok, f" (аркуш '{sheet_name}').")
    _write_if_missing(ZADACHI_PATH, build_zadachi)
    _write_if_missing(MASS_MOVEMENT_PATH, build_mass_movement, " (лише для 'Рапорти на масове переміщення').")
    print("Замініть вигаданих людей на реальних.")


if __name__ == "__main__":
    main()
