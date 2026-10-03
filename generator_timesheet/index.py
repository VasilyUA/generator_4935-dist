import os
import warnings

from InquirerPy.prompts.list import ListPrompt

from constants import (
    MONEY_PROJECT_RESOURCES_DIR,
    MONEY_PROJECT_SHEET_NAME,
    OUTPUT_DIR,
    PAYMENTS_REVIEW_FILE_NAME,
    REPORT_DIR,
    RESOURCES_DIR,
    REVIEW_FILE_NAME,
)
from content.oblik_timesheet import export_for_money_project
from generators.generate_payments_timesheet import generate_payments_timesheet
from generators.generate_rop_vop_statement import generate_rop_vop_statement
from generators.generate_timesheet import generate_timesheet
from utils.backup import backup_resources_and_output
from utils.folder_picker import pick_folder
from utils.logging_utils import print_green, print_red
from utils.output_folder import clear_output_directory, open_file
from utils.prompts import ask_yes_no

# За прямою вказівкою користувача - реальні resources/ОБЛІК.xlsx і resources/
# information_unit/*.xlsx мають Excel-функції (розширений формат умовного
# форматування/перевірки даних, "extLst" - типово для ОБЛІК.xlsx, і фігури/
# малюнки в одному з файлів information_unit), яких openpyxl НЕ вміє
# розібрати - друкує про це попередження при КОЖНОМУ читанні цих файлів. Не
# впливає на результат (цей проєкт читає/пише лише значення клітинок і власні
# кольори заливки - STATUS_COLORS - а не ці Excel-функції), лише "засмічує"
# термінал - тож приховано ТОЧКОВО (за текстом повідомлення), а не всі
# попередження взагалі (щоб не сховати випадково якесь ІНШЕ, дійсно корисне).
for _message in (
    r"Conditional Formatting extension is not supported.*",
    r"Data Validation extension is not supported.*",
    r"DrawingML support is incomplete.*",
):
    warnings.filterwarnings("ignore", message=_message, category=UserWarning)


def _open_if_confirmed(file_path):
    """Питає в терміналі, чи відкрити file_path, перш ніж це робити - за
    прямою вказівкою користувача: раніше файли з output/ відкривались
    автоматично, без запитання. default=True - Enter лишає попередню
    поведінку (відкрити), користувач має свідомо обрати "Ні", щоб
    пропустити."""
    if ask_yes_no(f"Відкрити {file_path}?", default=True):
        open_file(file_path)

GENERATE_PERSONEL_LIST = "personel_list"
GENERATE_PAYMENTS_LIST = "payments_list"
GENERATE_ROP_VOP_STATEMENT = "rop_vop_statement"
GENERATE_ALL = "generate_all"

_GENERATE_CHOICES = [
    {"name": "Файл обліку особового складу (ОБЛІК.xlsx)", "value": GENERATE_PERSONEL_LIST},
    {"name": "Файл ОБЛІК для виплат з БЧС", "value": GENERATE_PAYMENTS_LIST},
    {"name": "Відомість РОП/ВОП", "value": GENERATE_ROP_VOP_STATEMENT},
    {"name": "Згенерувати все", "value": GENERATE_ALL},
]

_HANDLERS = {
    GENERATE_PERSONEL_LIST: generate_timesheet,
    GENERATE_PAYMENTS_LIST: generate_payments_timesheet,
}

# За прямою вказівкою користувача - лише для пункту "Згенерувати все": ОБИДВА
# generate_timesheet/generate_payments_timesheet за замовчуванням пишуть у
# ТОЙ САМИЙ файл "ОБЛІК.xlsx" (навмисно - див. PAYMENTS_OUTPUT_FILE_NAME у
# constants.py, "обидва прогони завжди окремі, ніколи одночасно") - тут вони
# ЖИВУТЬ РАЗОМ у ТІЙ САМІЙ теці за ОДИН прогін, тож потребують РІЗНИХ імен:
# "ОБЛІК по стройовій.xlsx" (особовий склад) лишає "ОБЛІК.xlsx" (виплати з
# БЧС) під його ЗВИЧНОЮ назвою. Відомість РОП/ВОП теж отримує фіксовану
# назву (замість динамічної "підрозділ+місяць") - за прямою вказівкою
# користувача.
_ALL_PERSONEL_LIST_OUTPUT_FILE_NAME = os.path.join(OUTPUT_DIR, "ОБЛІК по стройовій.xlsx")
_ALL_ROP_VOP_STATEMENT_OUTPUT_FILE_NAME = os.path.join(OUTPUT_DIR, "Відомість 170_70.xlsx")

# Той самий review_file_name, що й ЗА ЗАМОВЧУВАННЯМ використовує відповідний
# generate_*_timesheet() у _HANDLERS вище (обидва викликаються тут БЕЗ
# аргументів) - потрібен тут окремо, щоб знати, ЯКИЙ саме файл відкрити.
_REVIEW_FILE_NAME_BY_CHOICE = {
    GENERATE_PERSONEL_LIST: REVIEW_FILE_NAME,
    GENERATE_PAYMENTS_LIST: PAYMENTS_REVIEW_FILE_NAME,
}

if __name__ == "__main__":
    choice = ListPrompt(message="Що зробити?", choices=_GENERATE_CHOICES, default=GENERATE_PERSONEL_LIST).execute()
    clear_output_directory(OUTPUT_DIR)

    if choice == GENERATE_ALL:
        # За прямою вказівкою користувача - усі ТРИ генератори ПОСПІЛЬ, у
        # ТІЙ САМІЙ теці, з ІМЕНАМИ файлів, підібраними вище (без цього ОБЛІК
        # від generate_timesheet і generate_payments_timesheet писали б у
        # ТОЙ САМИЙ "ОБЛІК.xlsx" і перезаписували б один одного). Жодного
        # запитання "відкрити файл?"/"записати кудись іще?" тут НЕМАЄ (за
        # прямою вказівкою користувача - лише шляхи, кожен генератор ВЖЕ
        # друкує свій власний результат кольором, print_green/print_red,
        # усередині _finish/generate_rop_vop_statement). ОКРЕМИЙ try/except
        # на КОЖЕН генератор - файл, заблокований в іншій програмі, зупиняє
        # ЛИШЕ ЦЕЙ ОДИН результат, а не решту двох.
        for _generate_one in (
            lambda: generate_timesheet(output_file_name=_ALL_PERSONEL_LIST_OUTPUT_FILE_NAME),
            generate_payments_timesheet,
            lambda: generate_rop_vop_statement(output_file_name=_ALL_ROP_VOP_STATEMENT_OUTPUT_FILE_NAME),
        ):
            try:
                _generate_one()
            except PermissionError as e:
                print_red(str(e))
    else:
        try:
            if choice == GENERATE_ROP_VOP_STATEMENT:
                # generate_rop_vop_statement() повертає (output_path,
                # people_count, mismatch_report_path) - ІНША форма, ніж у
                # решти 2 пунктів (немає unresolved/review-файлу - формальний
                # документ, зазвичай завжди зберігається, навіть якщо
                # people_count == 0), тож ОКРЕМА гілка, а не спільний
                # _HANDLERS[choice](). mismatch_report_path - НЕ None лише
                # якщо звірка з resources/schedule знайшла розбіжність - у
                # цьому разі output_path/people_count - None (сам файл
                # НЕ збережено, звірка блокує генерацію - див. докстрінг
                # generate_rop_vop_statement).
                _output_path, _people_count, _schedule_mismatch_path = generate_rop_vop_statement()
            else:
                # generate_timesheet() повертає (output_path, unresolved);
                # generate_payments_timesheet() - (output_path, unresolved,
                # mismatch_report_path) - третій елемент є ЛИШЕ для GENERATE_PAYMENTS_LIST
                # і ЛИШЕ якщо звіт розбіжностей (error_mis_statuses.xlsx) містить хоч
                # один запис (інакше None - нема сенсу відкривати порожній звіт).
                _output_path, _unresolved, *_extra = _HANDLERS[choice]()
        except PermissionError as e:
            # Файл заблокований іншою програмою (напр. ОБЛІК.xlsx відкритий у
            # Excel) - зрозуміле повідомлення замість traceback, автоматично
            # нічого не закриваємо (ризик втратити незбережені зміни).
            print_red(str(e))
        else:
            if choice == GENERATE_ROP_VOP_STATEMENT:
                # _output_path - None, якщо звірка з resources/schedule
                # знайшла розбіжність (генерацію заблоковано) - тоді
                # пропонується відкрити ЗВІТ розбіжностей замість самої
                # Відомості (яку в цьому разі взагалі не збережено).
                if _output_path is not None:
                    _open_if_confirmed(_output_path)
                else:
                    _open_if_confirmed(_schedule_mismatch_path)
            elif _unresolved:
                # Є що перевірити - пропонується відкрити файл "Потребує ручної
                # перевірки", а не сам ОБЛІК.xlsx (спершу варто прочитати
                # зауваження).
                _open_if_confirmed(_REVIEW_FILE_NAME_BY_CHOICE[choice])
            else:
                # Жодних зауважень - пропонується відкрити сам згенерований файл.
                _open_if_confirmed(_output_path)
            # Решта - error_mis_statuses.xlsx/експорт у money-проєкт - стосується
            # ЛИШЕ GENERATE_PAYMENTS_LIST (_extra/_unresolved там і взагалі не
            # визначені для GENERATE_ROP_VOP_STATEMENT - у нього немає ні того, ні
            # іншого, вище вже все зроблено).
            if choice != GENERATE_ROP_VOP_STATEMENT:
                # error_mis_statuses.xlsx ПРОПОНУЄТЬСЯ ВІДКРИТИ ДОДАТКОВО (незалежно
                # від того, який з файлів вище щойно відкрився) - за прямою вказівкою
                # користувача.
                if _extra and _extra[0]:
                    _open_if_confirmed(_extra[0])
                # За прямою вказівкою користувача - коли прогін "Файл ОБЛІК для
                # виплат з БЧС" ПОВНІСТЮ чистий (ні unresolved-записів, ні розбіжностей
                # error_mis_statuses.xlsx - _extra тут або порожній, або None) -
                # пропонується ТАКОЖ записати щойно згенерований файл кудись іще:
                # ПРОСТЕ підтвердження (без цільового шляху в самому запитанні), а
                # ЦІЛЬОВУ папку користувач обирає ОКРЕМО, через діалог - за прямою
                # вказівкою користувача (раніше шлях був зафіксований на сусідньому
                # проєкті generator_br_and_report_for_money - тепер користувач сам
                # вирішує щоразу, MONEY_PROJECT_RESOURCES_DIR лише підказка, звідки
                # діалог стартує).
                if choice == GENERATE_PAYMENTS_LIST and not _unresolved and not (_extra and _extra[0]):
                    if ask_yes_no(f"Записати {_output_path}?", default=False):
                        chosen_folder = pick_folder("Оберіть папку, куди записати файл", initial_dir=MONEY_PROJECT_RESOURCES_DIR)
                        if not chosen_folder:
                            print_red("Папку не обрано - файл нікуди не записано.")
                        else:
                            target_path = os.path.join(chosen_folder, os.path.basename(_output_path))
                            try:
                                export_for_money_project(_output_path, target_path, MONEY_PROJECT_SHEET_NAME)
                            except PermissionError as e:
                                print_red(str(e))
                            else:
                                print_green(target_path)

    # За прямою вказівкою користувача - ОСТАННЄ питання скрипта, БЕЗУМОВНО
    # (для КОЖНОГО пункту меню, і НЕЗАЛЕЖНО від того, чи генерація вище
    # вдалась - код нижче поза try/except/else вище) - чи зберегти ПОТОЧНІ
    # resources/ і output/ кудись іще (ручний "знімок" стану, напр. перед
    # початком нового місяця). "Ні" (значення за замовчуванням) - нічого не
    # змінюється, як і раніше.
    if ask_yes_no("Зберегти resources та output в іншій папці?", default=False):
        _chosen_backup_folder = pick_folder("Оберіть папку, куди зберегти resources та output")
        if not _chosen_backup_folder:
            print_red("Папку не обрано - resources та output нікуди не збережено.")
        else:
            try:
                _backup_target_dir = backup_resources_and_output(_chosen_backup_folder, REPORT_DIR, RESOURCES_DIR, OUTPUT_DIR)
            except OSError as e:
                print_red(str(e))
            else:
                if _backup_target_dir is None:
                    print_red(f"У {REPORT_DIR} не знайдено жодного рапорту з розпізнаваною датою - не вдалось назвати нову підтеку.")
                else:
                    print_green(_backup_target_dir)
