import os
import shutil


def clear_output_directory(directory):
    """Очищає directory (або створює, якщо її немає) перед кожним запуском
    index.py, щоб файли попереднього прогону не змішувались із новими.

    Файл, відкритий в іншій програмі (напр. ОБЛІК.xlsx в Excel чи його
    lock-файл "~$..."), не видаляється і лишається до наступного запуску:
    скрипт навмисно нічого не закриває, щоб не втратити незбережені зміни.
    Через такий файл і сама тека може лишитися - це теж не помилка."""
    if os.path.isdir(directory):
        shutil.rmtree(directory, onexc=_ignore_os_error)
    os.makedirs(directory, exist_ok=True)


def _ignore_os_error(_function, _path, exc):
    """Очищення output/ - "по змозі": якщо потрібний файл так і лишиться
    заблокованим, про це зрозуміло скаже вже Timesheet.save()."""
    if not isinstance(exc, OSError):
        raise exc


def clear_information_unit_directory(directory, keep_file_name):
    """Видаляє з directory усі файли, крім keep_file_name: цей файл не має
    щоденної дати й не приходить через Signal, тож перезавантаження файлів БЧС
    його не чіпає. Підтеки пропускаються. Файл, який не вдалося видалити
    (заблокований чи інша OSError), лишається, а решта видаляється далі -
    на відміну від shutil.rmtree, тут не "все або нічого"."""
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
    """Відкриває file_path у типовій програмі. Якщо не вийшло (os.startfile є
    лише на Windows, немає типової програми тощо), генерація від цього не
    зривається."""
    if file_path and os.path.isfile(file_path):
        try:
            os.startfile(file_path)
        except (AttributeError, OSError):
            pass


def open_folder(directory):
    """Як open_file, але для теки (відкривається Провідник)."""
    if directory and os.path.isdir(directory):
        try:
            os.startfile(directory)
        except (AttributeError, OSError):
            pass
