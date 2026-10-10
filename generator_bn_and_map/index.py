import os
import re
import sys
from datetime import date

from InquirerPy.prompts.list import ListPrompt

from constants import (
    DATA_FILE_NAME, MAP_ICONS_DIR, MAP_OUTPUT_DIR, OUTPUT_DIR, RESOURCES_DIR, REVIEW_FILE_NAME, STYLE_FILE_NAME,
)
from content.data_store import load_data, save_data
from generators.generate_bn import generate_bn, review_lines
from generators.generate_map import generate_map
from generators.pipeline import SourceError, load_latest_br, refresh_pbd_cache
from utils.logging_utils import print_green, print_purple, print_red
from utils.output_folder import clear_output_directory, open_file
from utils.prompts import ask_date, ask_text, ask_yes_no

GENERATE_ALL = "all"
GENERATE_BN = "bn"
GENERATE_MAP = "map"

_GENERATE_CHOICES = [
    {"name": "Бойовий наказ + карта (Кропива/Дельта)", "value": GENERATE_ALL},
    {"name": "Лише бойовий наказ", "value": GENERATE_BN},
    {"name": "Лише карта", "value": GENERATE_MAP},
]


def _print_warnings(title, warnings):
    if warnings:
        print_purple(title)
        for warning in warnings:
            print_purple(f"  - {warning}")


def _is_time(text):
    return bool(re.fullmatch(r"\d{1,2}[.:]\d{2}", (text or "").strip()))


def _update_pbd_cache(data):
    try:
        cache, warnings, used_files = refresh_pbd_cache(data, RESOURCES_DIR)
    except SourceError as e:
        print_red(str(e))
        return False
    data["pbd"] = cache
    try:
        save_data(DATA_FILE_NAME, data)
    except PermissionError as e:
        print_red(f"Не вдалось записати {DATA_FILE_NAME}: {e}")
        return False
    print_green(
        f"data.json оновлено з: {', '.join(os.path.basename(path) for path in used_files)} - "
        f"точок {len(cache['points'])}, ліній {len(cache['lines'])}."
    )
    _print_warnings("Під час оновлення з ПБД:", warnings)
    return True


def _generate_bn(data, pbd_cache):
    bn = data.get("bn") or {}
    bn_date = ask_date("Дата бойового наказу (дд.мм.рррр):", default=date.today())
    bn_time = ask_text("Час (ГГ.ХХ):", default=bn.get("TIME", "06.00"), validate=_is_time,
                       invalid_message="❌ Час має бути у форматі ГГ.ХХ.").replace(":", ".")
    bn_number = ask_text("Номер бойового наказу:", default=bn.get("NUMBER", ""))

    br = load_latest_br(RESOURCES_DIR)
    if br is None:
        print_red("У resources/ немає файлу БР*.docx - блоки з бойового розпорядження будуть порожні.")
    else:
        print_green(f"БР: {os.path.basename(br.path)} (№{br.number} від {br.date})")
    try:
        result = generate_bn(data, br, pbd_cache, bn_date, bn_time, bn_number, OUTPUT_DIR, STYLE_FILE_NAME)
    except (PermissionError, ValueError) as e:
        print_red(str(e))
        return None
    print_green(result.path)
    print_green(
        f"Автоматично написано {result.automatic_percent}% тексту бойового наказу; перевірити вручну - "
        f"{result.staff_values} штабних значень (координати й сигнали без джерела) у {len(result.manual)} абзацах."
    )
    _print_warnings("Бойовий наказ - перевірте:", result.warnings)
    return result


def _generate_map(data, pbd_cache):
    try:
        result = generate_map(data, pbd_cache, MAP_OUTPUT_DIR, MAP_ICONS_DIR)
    except PermissionError as e:
        print_red(str(e))
        return None
    print_green(f"{result.kmz_path} - точок {len(result.points)}, ліній {len(result.lines)}")
    print_green(result.csv_path)
    _print_warnings("Карта - перевірте:", result.warnings)
    return result


def _write_review_file(bn_result, map_result):
    """Усе, що варто перевірити очима, - одним файлом поруч з результатом
    (як "Потребує_ручної_перевірки.txt" у generator_timesheet)."""
    lines = review_lines(bn_result) if bn_result else []
    if map_result and map_result.warnings:
        lines += ["", "КАРТА", *[f"  - {w}" for w in map_result.warnings]]
    if not lines:
        return None
    path = os.path.join(OUTPUT_DIR, REVIEW_FILE_NAME)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    data = load_data(DATA_FILE_NAME)
    if not data:
        print_red(f"Немає {DATA_FILE_NAME} - див. README.md, розділ generator_bn_and_map.")
        return 1

    choice = ListPrompt(message="Що згенерувати?", choices=_GENERATE_CHOICES, default=GENERATE_ALL).execute()
    if ask_yes_no("Оновити всі позиції, ХАБи та інше з ПБД у data.json?", default=True):
        if not _update_pbd_cache(data):
            return 1
    pbd_cache = data.get("pbd")
    if not pbd_cache:
        print_red("У data.json ще немає даних з ПБД - запустіть ще раз і дайте відповідь «Так» на оновлення.")
        return 1

    # Як в інших проєктах репозиторію - output/ щоразу очищається, щоб не
    # змішувати результати різних запусків.
    clear_output_directory(OUTPUT_DIR)
    bn_result = _generate_bn(data, pbd_cache) if choice in (GENERATE_ALL, GENERATE_BN) else None
    map_result = _generate_map(data, pbd_cache) if choice in (GENERATE_ALL, GENERATE_MAP) else None
    review_path = _write_review_file(bn_result, map_result)
    if review_path:
        print_purple(f"Що перевірити - {review_path}")
    if bn_result and ask_yes_no(f"Відкрити {bn_result.path}?", default=True):
        open_file(bn_result.path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
