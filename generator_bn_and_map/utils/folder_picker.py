"""Вікно вибору теки (копія generator_timesheet/utils/folder_picker.py)."""
from tkinter import Tk, filedialog


def pick_folder(title, initial_dir=None):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title, initialdir=initial_dir)
    root.destroy()
    return folder
