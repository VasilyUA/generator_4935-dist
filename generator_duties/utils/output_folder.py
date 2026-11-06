import os
import shutil


def clear_output_directory(directory):
    """Очищує (чи створює) directory перед кожною генерацією - як у
    generator_bn_and_map/utils/output_folder.py. Файл, відкритий в іншій
    програмі (напр. попередня книга у Word), не видаляється і не зупиняє запуск."""
    if os.path.isdir(directory):
        shutil.rmtree(directory, onexc=_ignore_os_error)
    os.makedirs(directory, exist_ok=True)


def _ignore_os_error(_function, _path, exc):
    if not isinstance(exc, OSError):
        raise exc


def open_file(file_path):
    """Відкриває file_path (файл чи теку) у типовій програмі ОС - помилка
    відкриття не зриває сам запуск."""
    if file_path and os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except (AttributeError, OSError):
            pass
