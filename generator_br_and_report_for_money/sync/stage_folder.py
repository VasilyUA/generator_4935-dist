import os
import re
import shutil

from InquirerPy.prompts.list import ListPrompt

from utils.logging_utils import print_red, print_green
from utils.folder_picker import pick_folder
from utils.prompts import ask_yes_no as _ask_yes_no
from sync.constants_sync import load_number_dicts_from_legacy_backup, sync_document_numbers_from_folder
from utils.project_settings import read_numbers_workbook, write_numbers_workbook
from utils.settings_book import SETTINGS_FILE_NAME, SettingsError

_STAGE_FOLDER_RE = re.compile(r'^(\d+)\s*етап$', re.IGNORECASE)

# Номери документів етапу (щоденні БР, ЗАВДАННЯ, БЗ) - окрема книга в папці етапу
# з тими самими трьома аркушами номерів, що й resources/НАЛАШТУВАННЯ.xlsx. Решта
# налаштувань спільна для всіх етапів і читається з resources/ проєкту.
STAGE_NUMBERS_FILE_NAME = "НОМЕРИ ДОКУМЕНТІВ.xlsx"
# Колишній формат: копія constants.py з номерами прямо в коді.
LEGACY_STAGE_BACKUP_FILE_NAME = "constants.py.bak"


def prompt_should_open_stage_folder():
    return _ask_yes_no("Відкрити папку проєкту (з папками етапів)?", False)


def find_stage_folders(parent_folder):
    """Повертає папки виду '1 етап', '2 етап', ... з parent_folder,
    відсортовані від більшого номера етапу до меншого."""
    stages = []
    for entry in os.listdir(parent_folder):
        full_path = os.path.join(parent_folder, entry)
        if not os.path.isdir(full_path):
            continue
        match = _STAGE_FOLDER_RE.match(entry)
        if match:
            stages.append((int(match.group(1)), entry, full_path))

    stages.sort(key=lambda stage: stage[0], reverse=True)
    return stages


def prompt_stage_selection(initial_dir):
    """Відкриває діалог вибору папки (за замовчуванням — initial_dir): або одразу папки
    самого етапу (тоді її номер визначається з назви й додатковий вибір не пропонується),
    або папки місяця з кількома '1 етап', '2 етап', ... — тоді пропонується обрати один з
    них (від більшого номера етапу до меншого).

    Повертає (шлях_до_папки_місяця, шлях_до_обраного_етапу, це_останній_етап) або
    (None, None, None), якщо папку/етап не обрано."""
    picked_folder = pick_folder("Оберіть папку місяця з етапами (або одразу папку етапу)", initial_dir)
    if not picked_folder:
        print_red("Папку не обрано.")
        return None, None, None

    picked_folder = picked_folder.rstrip("\\/")
    direct_match = _STAGE_FOLDER_RE.match(os.path.basename(picked_folder))

    if direct_match:
        # Обрано саму папку етапу — номер беремо з її назви, вибір етапу не потрібен.
        month_folder = os.path.dirname(picked_folder)
        stage_path = picked_folder
        stage_number = int(direct_match.group(1))
    else:
        month_folder = picked_folder
        stages = find_stage_folders(month_folder)
        if not stages:
            print_red(f"У папці {month_folder} не знайдено папок етапів (напр. '1 етап').")
            return None, None, None

        stage_choice = ListPrompt(
            message="Оберіть етап:",
            choices=[{"name": name, "value": (number, full_path)} for number, name, full_path in stages],
            default=(stages[0][0], stages[0][2]),
        ).execute()

        if stage_choice is None:
            print_red("Етап не обрано.")
            return None, None, None

        stage_number, stage_path = stage_choice

    # Визначаємо, чи це останній етап, за сусідніми папками в папці місяця
    # (працює однаково незалежно від того, яку саме папку обрали вище).
    stages_in_month_folder = find_stage_folders(month_folder)
    last_stage_number = stages_in_month_folder[0][0] if stages_in_month_folder else stage_number

    return month_folder, stage_path, stage_number == last_stage_number


def bootstrap_empty_stage_folder(stage_dir, project_dir):
    """Якщо stage_dir ЦІЛКОМ порожня (щойно створена папка етапу, без жодного файлу
    чи підпапки) - копіює туди resources/ (ОБЛІК.xlsx/ПРИДАНІ.xlsx тощо, бо режим
    етапу надалі читає їх САМЕ звідти - PERSONEL_LIST_FILE_NAME/DOWRIES_LIST_FILE_NAME,
    constants.py) - підтверджено користувачем: без цього щойно створена папка етапу
    лишається без жодних ресурсів, і наступне ж читання ОБЛІК.xlsx звідти одразу
    падає. Книга налаштувань (з її копією "(до оновлення)") і колишній data.json не копіюються: налаштування
    спільні для всіх етапів і читаються з resources/ проєкту, а номери документів
    етапу живуть в окремій книзі STAGE_NUMBERS_FILE_NAME (sync_stage_constants_backup).

    Папка НЕ порожня (навіть один файл чи підпапка) - нічого не робить і не
    друкує: лише ЦІЛКОМ порожня, щойно створена папка вважається "новою"."""
    if os.listdir(stage_dir):
        return

    project_resources_dir = os.path.join(project_dir, "resources")
    if os.path.isdir(project_resources_dir):
        shutil.copytree(
            project_resources_dir, os.path.join(stage_dir, "resources"),
            ignore=shutil.ignore_patterns(f"{os.path.splitext(SETTINGS_FILE_NAME)[0]}*.xlsx", "data.json"),
        )

    print_green(f"Папка етапу {stage_dir} була порожня - скопійовано resources/ з {project_dir}.")


def _write_stage_numbers(path, number_dicts):
    try:
        write_numbers_workbook(path, number_dicts)
    except SettingsError as error:
        print_red(f"Не вдалось записати номери етапу: {error}")
        return
    print_green(f"Номери документів етапу записано в {path}.")


def _replace_numbers(number_dicts, loaded):
    for name, number_dict in number_dicts.items():
        number_dict.clear()
        number_dict.update(loaded[name])


def sync_stage_constants_backup(
    stage_dir, month, month_folder, is_last_stage, settings_file_path,
    day_keys, week_keys, save_keys,
    number_of_documents_brs_every_day, number_of_documents_brs_every_week, number_of_documents_brs_save,
):
    """Вирішує, чи оновлювати книгу номерів етапу (STAGE_NUMBERS_FILE_NAME у папці
    етапу), чи прочитати наявну.

    Питання "Зчитати обрану папку з документами?" пропонується лише коли папка місяця
    (батьківська щодо етапу) називається так само, як обраний місяць генерації, І обраний
    етап — останній за нумерацією в цій папці, АБО коли книги номерів етапу ще немає.

    - Якщо на це питання відповісти "Так" і обрати папку — номери синхронізуються в
      книгу налаштувань проєкту (як і в звичайному сценарії без етапів), і лише після
      цього поточний стан записується в книгу номерів етапу (перезаписуючи попередню).
    - Якщо відповісти "Ні" (або питання не задавалось) — книга налаштувань не чіпається:
      якщо книга номерів етапу вже є, номери просто зчитуються з неї; якщо її ще немає,
      але є constants.py.bak колишнього формату - номери беруться з нього (і далі живуть
      у книзі етапу); інакше книга етапу створюється з поточних номерів.

    Повертає шлях до книги номерів етапу (constants.STAGE_CONSTANTS_BACKUP_PATH)."""
    backup_path = os.path.join(stage_dir, STAGE_NUMBERS_FILE_NAME)
    legacy_backup_path = os.path.join(stage_dir, LEGACY_STAGE_BACKUP_FILE_NAME)
    number_dicts = {
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY": number_of_documents_brs_every_day,
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK": number_of_documents_brs_every_week,
        "NUMBER_OF_DOCUMENTS_BRS_SAVE": number_of_documents_brs_save,
    }
    keys_by_name = {
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY": day_keys,
        "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK": week_keys,
        "NUMBER_OF_DOCUMENTS_BRS_SAVE": save_keys,
    }
    is_current_month_folder = os.path.basename(month_folder.rstrip("\\/")) == month
    backup_exists = os.path.isfile(backup_path)
    should_offer_sync = (is_current_month_folder and is_last_stage) or not backup_exists

    if should_offer_sync and _ask_yes_no("Зчитати обрану папку з документами (для номерів БАТ / БЗ)?", False):
        chosen_folder = pick_folder("Оберіть папку з документами")
        if chosen_folder:
            sync_document_numbers_from_folder(
                chosen_folder, settings_file_path,
                number_of_documents_brs_every_day, day_keys,
                number_of_documents_brs_every_week, week_keys,
                number_of_documents_brs_save, save_keys,
            )
        else:
            print_red("Папку не обрано — синхронізація номерів пропущена.")

        _write_stage_numbers(backup_path, number_dicts)
        return backup_path

    if backup_exists:
        try:
            loaded = read_numbers_workbook(backup_path, keys_by_name)
        except SettingsError as error:
            print_red(f"Номери етапу не прочитано ({error}) - беру номери з книги налаштувань проєкту.")
            return backup_path
        _replace_numbers(number_dicts, loaded)
        print_green(f"Номери БАТ/БЗ зчитано з {backup_path}.")
        return backup_path

    legacy = load_number_dicts_from_legacy_backup(legacy_backup_path, keys_by_name)
    if legacy is not None:
        _replace_numbers(number_dicts, legacy)
        print_green(f"Номери БАТ/БЗ зчитано з {legacy_backup_path} (колишній формат).")
    _write_stage_numbers(backup_path, number_dicts)
    return backup_path
