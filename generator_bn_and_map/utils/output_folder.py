import os
import shutil


def clear_output_directory(directory):
    """Очищує (чи створює) directory перед кожним запуском - як у
    generator_timesheet/utils/output_folder.py. Файл, відкритий в іншій
    програмі (напр. попередній БН у Word), не видаляється і не зупиняє
    запуск - лишається до наступного разу."""
    if os.path.isdir(directory):
        shutil.rmtree(directory, onexc=_ignore_os_error)
    os.makedirs(directory, exist_ok=True)


def _ignore_os_error(_function, _path, exc):
    if not isinstance(exc, OSError):
        raise exc


def open_file(file_path):
    """Відкриває file_path у типовій програмі ОС (копія з
    generator_timesheet/utils/output_folder.py) - будь-яка помилка відкриття не
    зриває сам запуск."""
    if file_path and os.path.isfile(file_path):
        try:
            os.startfile(file_path)
        except (AttributeError, OSError):
            pass
