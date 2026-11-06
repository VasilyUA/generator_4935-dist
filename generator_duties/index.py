import json
import os
import sys
from datetime import date

from InquirerPy.prompts.list import ListPrompt

from constants import DATA_FILE_NAME, DUTIES_BACKUP_FILE_NAME, DUTIES_FILE_NAME, OUTPUT_DIR, RESOURCES_DIR, REVIEW_FILE_NAME
from content.duties_table import TableError, read_table
from content.positions import auto_suffixes, group_positions
from content.prefill import label_names
from content.regulation_draft import ensure_regulation_draft
from content.roster_reader import RosterError, read_roster
from content.sources import load_sources
from generators.generate_books import generate_books, review_lines
from generators.update_table import stale_positions, update_table
from utils.backup import save_output
from utils.folder_picker import pick_folder
from utils.logging_utils import print_green, print_purple, print_red
from utils.output_folder import clear_output_directory, open_file
from utils.prompts import ask_text, ask_yes_no

UPDATE_TABLE = "update"
GENERATE = "generate"

_CHOICES = [
    {"name": "Оновити таблицю обов'язків зі штату (resources/Обов'язки.xlsx)", "value": UPDATE_TABLE},
    {"name": "Згенерувати документи: функціональні + посадові обов'язки", "value": GENERATE},
]


def _load_data():
    if not os.path.isfile(DATA_FILE_NAME):
        print_red(f"Немає {DATA_FILE_NAME} - див. README.md, розділ generator_duties.")
        return None
    with open(DATA_FILE_NAME, encoding="utf-8") as f:
        return json.load(f)


def _print_list(title, items, limit=15):
    if not items:
        return
    print_purple(f"{title} ({len(items)}):")
    for item in items[:limit]:
        print_purple(f"  - {item}")
    if len(items) > limit:
        print_purple(f"  … ще {len(items) - limit}")


def _update(positions, sources, offer_to_open=True):
    table = read_table(DUTIES_FILE_NAME, label_names(sources)) if os.path.isfile(DUTIES_FILE_NAME) else None
    stale = stale_positions(table, positions) if table else []
    remove = False
    if stale:
        _print_list("На аркуші «Посади» є посади, яких більше немає в штаті", stale)
        remove = ask_yes_no(f"Видалити ці {len(stale)} посад з таблиці?", default=False)
    result = update_table(DUTIES_FILE_NAME, DUTIES_BACKUP_FILE_NAME, positions, sources, remove_stale=remove)
    print_green(f"{'Створено' if result.created else 'Оновлено'} {result.path}"
                + ("" if result.created else f" (попередня копія - {os.path.basename(DUTIES_BACKUP_FILE_NAME)})"))
    if result.migrated:
        print_green("Таблицю перетворено на новий вигляд: аркуш «Посади» (кожна посада один раз) + набори обов'язків "
                    "без повторів.")
    print_green("Наборів обов'язків: " + ", ".join(f"{sheet} - {count}" for sheet, count in result.set_counts.items()))
    _print_list("Додано посад", result.added)
    _print_list("Заповнено заново (були лише «ЗАПОВНИТИ»)", result.refilled)
    _print_list("Видалено посад", result.removed)
    for sheet, titles in result.to_fill.items():
        _print_list(f"[{sheet}] Посади з «ЗАПОВНИТИ» - допишіть текст", titles)
    if offer_to_open and ask_yes_no(f"Відкрити {os.path.basename(DUTIES_FILE_NAME)}?", default=True):
        open_file(DUTIES_FILE_NAME)
    return 0


def _save_result(year):
    folder = pick_folder("Оберіть папку, куди зберегти результат")
    if not folder:
        print_red("Папку не обрано - результат лишився лише в output/.")
        return
    try:
        target = save_output(OUTPUT_DIR, folder, f"Обов'язки {year}")
    except OSError as e:
        print_red(str(e))
    else:
        print_green(f"Результат збережено в {target}")


def _generate(positions, sources, data):
    table = read_table(DUTIES_FILE_NAME, label_names(sources)) if os.path.isfile(DUTIES_FILE_NAME) else None
    if table is None or table.legacy:
        print_purple(f"{'Ще немає' if table is None else 'Перетворюю'} {os.path.basename(DUTIES_FILE_NAME)} - "
                     f"{'створюю' if table is None else 'оновлюю'} її зі штату.")
        _update(positions, sources, offer_to_open=False)
        table = read_table(DUTIES_FILE_NAME, label_names(sources))
    year = int(ask_text("На який рік:", default=date.today().year + 1, validate=lambda t: t.strip().isdigit(),
                        invalid_message="❌ Рік - число, напр. 2027."))
    clear_output_directory(OUTPUT_DIR)
    print_green("Генерую книги (Word оновлює ЗМІСТ і рахує аркуші - це може тривати хвилину)…")
    results, signer = generate_books(positions, table, sources, data, year, OUTPUT_DIR)
    for result in results:
        pages = f" - сторінок {result.pages}, аркушів {result.sheets}" if result.pages else ""
        print_green(f"{result.path}{pages}")
    print_green(f"У листі ознайомлення - {results[0].blank_rows} порожніх рядків для нових в/с "
                f"(кількість - data.json → blank_rows).")
    lines = review_lines(results, signer, sources.warnings)
    if lines:
        path = os.path.join(OUTPUT_DIR, REVIEW_FILE_NAME)
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        print_purple(f"Що перевірити ({len(lines)}) - {path}")
    if ask_yes_no("Відкрити теку з результатом?", default=True):
        open_file(OUTPUT_DIR)
    elif ask_yes_no("Зберегти результат в іншій папці?", default=True):
        _save_result(year)
    return 0


def main():
    data = _load_data()
    if data is None:
        return 1
    choice = ListPrompt(message="Що зробити?", choices=_CHOICES, default=GENERATE).execute()
    try:
        roster = read_roster(RESOURCES_DIR, data.get("roster"))
    except RosterError as e:
        print_red(str(e))
        return 1
    suffixes = (data.get("roster") or {}).get("position_suffixes") or auto_suffixes(roster.slots)
    positions = group_positions(roster.slots, suffixes)
    vacant = sum(1 for slot in roster.slots if slot.vacant)
    print_green(f"Штат: {os.path.basename(roster.path)} (аркуш «{roster.sheet}», колонка посади «{roster.headers['position']}») - "
                f"посад {len(roster.slots)}, з них вакантних {vacant}; унікальних посад {len(positions)}.")
    if suffixes:
        print_green("З назв посад прибирається: " + ", ".join(f"«{suffix}»" for suffix in suffixes))
    try:
        draft = ensure_regulation_draft(RESOURCES_DIR, positions, roster.slots, data, suffixes)
    except PermissionError as e:
        print_red(str(e))
        draft = None
    if draft:
        print_green(f"Створено чернетку Положення про батальйон зі структури штату: {draft} - перегляньте й допрацюйте "
                    f"її у Word (програма її більше не перезаписує).")
    sources = load_sources(RESOURCES_DIR)
    if sources.used_files:
        print_green("Джерела: " + ", ".join(os.path.basename(path) for path in sources.used_files))
    for warning in sources.warnings:
        print_purple(f"  - {warning}")
    try:
        if choice == UPDATE_TABLE:
            return _update(positions, sources)
        return _generate(positions, sources, data)
    except (TableError, PermissionError) as e:
        print_red(str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
