import os
import re
import sys
from datetime import date

from InquirerPy.prompts.list import ListPrompt

from constants import (
    DATA_AUTO_FILE_NAME, DATA_FILE_NAME, MAP_ICONS_DIR, MAP_OUTPUT_DIR, OUTPUT_DIR, RESOURCES_DIR, REVIEW_FILE_NAME,
    STYLE_FILE_NAME,
)
from content.data_store import PREVIOUS_FIELDS, load_data, migrate_layout, save_auto_data
from content.staff import apply_delta, document_placeholders, load_staff, staff_to_data
from generators.generate_bn import (
    SOURCE_BR, SOURCE_PBD, generate_bn, priority_order, priority_overview, review_lines,
)
from generators.generate_map import generate_map
from generators.pipeline import (
    SourceError, load_extras, load_latest_br, refresh_pbd_cache, remove_backup_files, save_missing_icons,
)
from utils.backup import backup_resources_and_output, snapshot_folder_name
from utils.folder_picker import pick_folder
from utils.logging_utils import print_green, print_purple, print_red
from utils.output_folder import clear_output_directory, open_file
from utils.prompts import ask_date, ask_source_priority, ask_text, ask_yes_no

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


def _save_auto(data):
    try:
        save_auto_data(DATA_AUTO_FILE_NAME, data)
    except PermissionError as e:
        print_red(f"Не вдалось записати {DATA_AUTO_FILE_NAME}: {e}")
        return False
    return True


def _remember_run(data, bn_result, extras):
    """Що програма запам'ятовує в data_auto.json після генерації: штабні
    рішення з координатами, оновленими з карти Дельти, і значення цього наказу,
    взяті з БР чи ПБД (координати КСП/ЗКП, н.п. КСП, видання карти) - для
    наступного наказу це "попередній бойовий наказ"."""
    entries = load_staff(data)
    apply_delta(entries, extras.delta)
    if entries:
        data["staff"] = staff_to_data(entries, data.get("staff"))
    if bn_result is not None:
        for key, (value, source) in bn_result.priority_values.items():
            if value and source in (SOURCE_BR, SOURCE_PBD):
                data.setdefault("bn", {})[PREVIOUS_FIELDS[key]] = value
    _save_auto(data)


def _ask_priority(data, br, pbd_cache):
    """Показує, що є в кожному джерелі, і питає, що головніше - БР чи ПБД."""
    print_purple("Дані, що можуть бути і в БР, і в ПБД:")
    for line in priority_overview(br, pbd_cache, data.get("bn")):
        print_purple(f"  - {line}")
    return priority_order(ask_source_priority(default=SOURCE_BR))


def _update_pbd_cache(data):
    try:
        cache, warnings, used_files = refresh_pbd_cache(data, RESOURCES_DIR)
    except SourceError as e:
        print_red(str(e))
        return False
    data["pbd"] = cache
    if not _save_auto(data):
        return False
    print_green(
        f"Кеш ПБД (data_auto.json) оновлено з: {', '.join(os.path.basename(path) for path in used_files)} - "
        f"точок {len(cache['points'])}, ліній {len(cache['lines'])}."
    )
    _print_warnings("Під час оновлення з ПБД:", warnings)
    return True


def _generate_bn(data, pbd_cache, extras):
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
    priority = _ask_priority(data, br, pbd_cache)
    try:
        result = generate_bn(data, br, pbd_cache, bn_date, bn_time, bn_number, OUTPUT_DIR, STYLE_FILE_NAME, extras,
                             priority)
    except (PermissionError, ValueError) as e:
        print_red(str(e))
        return None
    print_green(result.path)
    print_green(
        f"Автоматично написано {result.automatic_percent}% тексту бойового наказу; штабні значення: "
        f"{result.structured_values} - з data_auto.json → staff, {result.staff_values} - перенесені з попереднього наказу."
    )
    if result.signals is not None and result.signals.path:
        print_green(f"{result.signals.path} - виправлена таблиця сигналів "
                    f"(додано {len(result.signals.added)}, змінено {len(result.signals.updated)})")
    _print_warnings("Бойовий наказ - перевірте:", result.warnings)
    return result


def _generate_map(data, pbd_cache, extras):
    # На карту - лише штабні рішення, які є в тексті бойового наказу.
    in_order = document_placeholders(data)
    staff_entries = [entry for entry in load_staff(data) if entry.placeholder in in_order]
    apply_delta(staff_entries, extras.delta)
    saved = save_missing_icons(MAP_ICONS_DIR, extras.icons)
    if saved:
        print_green(f"Значки SIDC з KMZ Дельти збережено в resources/map_icons: {saved}")
    try:
        result = generate_map(data, pbd_cache, MAP_OUTPUT_DIR, MAP_ICONS_DIR, extras.delta, staff_entries, extras.icons)
    except PermissionError as e:
        print_red(str(e))
        return None
    print_green(f"{result.kmz_path} - точок {len(result.points)}, ліній {len(result.lines)} "
                f"(закруглено {result.stats.get('smoothed', 0)})")
    print_green(result.csv_path)
    _print_warnings("Карта - перевірте:", result.warnings)
    return result


def _map_review_lines(map_result):
    stats = map_result.stats or {}
    lines = ["", "КАРТА"]
    if stats:
        lines.append(
            f"  Позицій з ПБД: {stats.get('pbd_points', 0)} (з них знайдено на карті Дельти - {stats.get('matched', 0)}); "
            f"штабних рішень з data_auto.json: {stats.get('staff', 0)}; штабних рішень з карти Дельти: "
            f"{stats.get('carried', 0)}; не перенесено з карти Дельти (немає ні в ПБД, ні в наказі): "
            f"{stats.get('skipped', 0)}; прибрано застарілих позицій: {stats.get('stale', 0)}; "
            f"закруглено ліній і районів: {stats.get('smoothed', 0)}."
        )
    return lines + [f"  - {w}" for w in map_result.warnings]


def _write_review_file(bn_result, map_result):
    """Усе, що варто перевірити очима, - одним файлом поруч з результатом
    (як "Потребує_ручної_перевірки.txt" у generator_timesheet)."""
    lines = review_lines(bn_result) if bn_result else []
    if map_result and (map_result.warnings or map_result.stats):
        lines += _map_review_lines(map_result)
    if not lines:
        return None
    path = os.path.join(OUTPUT_DIR, REVIEW_FILE_NAME)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    if os.path.isfile(DATA_FILE_NAME) and migrate_layout(DATA_FILE_NAME, DATA_AUTO_FILE_NAME):
        print_green("data.json розділено: ручні поля лишились у data.json, решту програма веде в data_auto.json.")
    data = load_data(DATA_FILE_NAME, DATA_AUTO_FILE_NAME)
    if not data:
        print_red(f"Немає {DATA_FILE_NAME} - див. README.md, розділ generator_bn_and_map.")
        return 1

    choice = ListPrompt(message="Що згенерувати?", choices=_GENERATE_CHOICES, default=GENERATE_ALL).execute()
    if ask_yes_no("Оновити всі позиції, ХАБи та інше з ПБД?", default=True):
        if not _update_pbd_cache(data):
            return 1
    pbd_cache = data.get("pbd")
    if not pbd_cache:
        print_red("Ще немає даних з ПБД - запустіть ще раз і дайте відповідь «Так» на оновлення.")
        return 1

    extras, used_files, problems = load_extras(RESOURCES_DIR)
    if used_files:
        print_green("Додаткові джерела: " + ", ".join(os.path.basename(path) for path in used_files))
    _print_warnings("Джерела:", problems)

    # Як в інших проєктах репозиторію - output/ щоразу очищається, щоб не
    # змішувати результати різних запусків.
    clear_output_directory(OUTPUT_DIR)
    bn_result = _generate_bn(data, pbd_cache, extras) if choice in (GENERATE_ALL, GENERATE_BN) else None
    map_result = _generate_map(data, pbd_cache, extras) if choice in (GENERATE_ALL, GENERATE_MAP) else None
    review_path = _write_review_file(bn_result, map_result)
    if review_path:
        print_purple(f"Що перевірити - {review_path}")
    _remember_run(data, bn_result, extras)
    removed = remove_backup_files(RESOURCES_DIR)
    if removed:
        print_green(f"Видалено резервних копій (.bak) з resources/: {len(removed)}")
    if bn_result and ask_yes_no(f"Відкрити {bn_result.path}?", default=True):
        open_file(bn_result.path)
    # Як у generator_timesheet - останнє питання: зберегти поточні resources/ і
    # output/ ще кудись (ручний "знімок" перед наступним плануванням).
    if ask_yes_no("Зберегти resources та output в іншій папці?", default=False):
        _backup(bn_result.bn_date if bn_result else None, pbd_cache)
    return 0


def _backup(bn_date, pbd_cache):
    folder = pick_folder("Оберіть папку, куди зберегти resources та output")
    if not folder:
        print_red("Папку не обрано - resources та output нікуди не збережено.")
        return
    name = snapshot_folder_name(bn_date, (pbd_cache or {}).get("report_datetime", ""))
    try:
        target = backup_resources_and_output(folder, name, RESOURCES_DIR, OUTPUT_DIR)
    except OSError as e:
        print_red(str(e))
    else:
        print_green(f"resources та output збережено в {target}")


if __name__ == "__main__":
    sys.exit(main())
