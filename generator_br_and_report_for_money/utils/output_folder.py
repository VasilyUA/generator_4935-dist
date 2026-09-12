import os


def open_output_folder(directory):
    if os.path.isdir(directory):
        try:
            os.startfile(directory)
        except AttributeError:
            pass


def open_file(file_path):
    if file_path and os.path.isfile(file_path):
        try:
            os.startfile(file_path)
        except AttributeError:
            pass
