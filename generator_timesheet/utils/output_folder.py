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
