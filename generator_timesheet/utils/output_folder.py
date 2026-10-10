import os
import shutil


def clear_output_directory(directory):
    """Очищує (чи створює, якщо ще нема) directory - перед КОЖНИМ запуском
    (index.py), щоб файли попереднього запуску (можливо, вже неактуальний
    статус/інша генерація) не змішувались із щойно сформованими.

    Файл, який наразі відкритий в іншій програмі (напр. ОБЛІК.xlsx у Excel -
    сам файл чи його тимчасовий "~$..." lock-файл), НЕ видаляється - без
    падіння з PermissionError, просто лишається як є до наступного запуску,
    коли користувач його закриє (сам скрипт нічого не закриває автоматично -
    свідоме рішення, щоб не втратити НЕЗБЕРЕЖЕНІ зміни, якщо вони там є). Так
    само толерується "тека не спорожніла" (OSError) - ПРЯМИЙ наслідок такого
    пропущеного файлу всередині: сама тека, отже, теж лишається як є."""
    if os.path.isdir(directory):
        shutil.rmtree(directory, onexc=_ignore_os_error)
    os.makedirs(directory, exist_ok=True)


def _ignore_os_error(_function, _path, exc):
    """Ця тека лише "по можливості" чиститься перед новим запуском - не
    критично для коректності (якщо цільовий файл ЗАЛИШИТЬСЯ заблокованим і на
    запис - про це чітко повідомить Timesheet.save(), а не ця функція)."""
    if not isinstance(exc, OSError):
        raise exc


def clear_information_unit_directory(directory, keep_file_name):
    """На відміну від clear_output_directory вище - прибирає З directory
    усі файли, КРІМ keep_file_name (за прямою вказівкою користувача: цей
    файл не має щоденної дати й не надходить через Signal, тож
    перезавантаження файлів БЧС його не має торкатись). Підпапок тут не
    очікується (information_unit - "пласка" тека) - якщо трапиться, вона
    просто пропускається, а не видаляється. Файл, заблокований в іншій
    програмі (PermissionError) чи взагалі будь-яка інша OSError при спробі
    видалення - лишається як є; це НЕ зупиняє видалення решти файлів (на
    відміну від shutil.rmtree, тут немає "усе або нічого")."""
    if not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)
        return
    for name in os.listdir(directory):
        if name == keep_file_name:
            continue
        path = os.path.join(directory, name)
        if not os.path.isfile(path):
            continue
        try:
            os.remove(path)
        except OSError:
            pass


def open_file(file_path):
    """Відкриває file_path у типовій програмі ОС (той самий підхід, що й у
    generator_br_and_report_for_money/utils/output_folder.py) - AttributeError
    (os.startfile є лише на Windows) чи будь-яка інша помилка відкриття
    (немає типової програми тощо) не повинні зривати сам запуск генерації."""
    if file_path and os.path.isfile(file_path):
        try:
            os.startfile(file_path)
        except (AttributeError, OSError):
            pass
