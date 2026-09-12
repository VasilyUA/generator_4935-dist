import os
import re
import shutil

from InquirerPy.prompts.list import ListPrompt

from utils.logging_utils import print_red, print_green
from utils.folder_picker import pick_folder
from utils.prompts import ask_yes_no as _ask_yes_no
from sync.constants_sync import load_number_dicts_from_backup, sync_document_numbers_from_folder

_STAGE_FOLDER_RE = re.compile(r'^(\d+)\s*етап$', re.IGNORECASE)


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


def sync_stage_constants_backup(
    stage_dir, month, month_folder, is_last_stage, constants_file_path,
    day_keys, week_keys, save_keys,
    number_of_documents_brs_every_day, number_of_documents_brs_every_week, number_of_documents_brs_save,
):
    """Вирішує, чи оновлювати constants.py.bak у папці етапу, чи прочитати наявний.

    Питання "Зчитати обрану папку з документами?" пропонується лише коли папка місяця
    (батьківська щодо етапу) називається так само, як обраний місяць генерації, І обраний
    етап — останній за нумерацією в цій папці, АБО коли бекапу ще немає взагалі.

    - Якщо на це питання відповісти "Так" і обрати папку — номери синхронізуються в сам
      constants.py (як і в звичайному сценарії без етапів), і лише після цього поточний
      стан записується в constants.py.bak етапу (перезаписуючи попередній).
    - Якщо відповісти "Ні" (або питання не задавалось) — constants.py не чіпається: якщо
      constants.py.bak вже є, номери просто зчитуються з нього; якщо його ще немає —
      і тільки тоді він створюється (копією поточного constants.py)."""
    backup_path = os.path.join(stage_dir, "constants.py.bak")
    is_current_month_folder = os.path.basename(month_folder.rstrip("\\/")) == month
    backup_exists = os.path.isfile(backup_path)
    should_offer_sync = (is_current_month_folder and is_last_stage) or not backup_exists

    if should_offer_sync and _ask_yes_no("Зчитати обрану папку з документами (для номерів БАТ / БЗ)?", False):
        chosen_folder = pick_folder("Оберіть папку з документами")
        if chosen_folder:
            sync_document_numbers_from_folder(
                chosen_folder, constants_file_path,
                number_of_documents_brs_every_day, day_keys,
                number_of_documents_brs_every_week, week_keys,
                number_of_documents_brs_save, save_keys,
            )
        else:
            print_red("Папку не обрано — синхронізація номерів пропущена.")

        shutil.copyfile(constants_file_path, backup_path)
        print_green(f"constants.py.bak згенеровано у {backup_path}.")
        return backup_path

    if backup_exists:
        day_dict, week_dict, save_dict = load_number_dicts_from_backup(backup_path, day_keys, week_keys, save_keys)
        if day_dict is not None and week_dict is not None and save_dict is not None:
            number_of_documents_brs_every_day.clear()
            number_of_documents_brs_every_day.update(day_dict)
            number_of_documents_brs_every_week.clear()
            number_of_documents_brs_every_week.update(week_dict)
            number_of_documents_brs_save.clear()
            number_of_documents_brs_save.update(save_dict)
            print_green(f"Номери БАТ/БЗ зчитано з {backup_path}.")
        return backup_path

    shutil.copyfile(constants_file_path, backup_path)
    print_green(f"constants.py.bak згенеровано у {backup_path}.")
    return backup_path
