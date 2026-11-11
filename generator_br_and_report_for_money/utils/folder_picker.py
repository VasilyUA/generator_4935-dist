from tkinter import Tk, filedialog


def pick_folder(title, initial_dir=None):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title, initialdir=initial_dir)
    root.destroy()
    return folder


def pick_file(title, initial_dir=None, filetypes=(("Word documents", "*.docx"), ("Усі файли", "*.*"))):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(title=title, initialdir=initial_dir, filetypes=filetypes)
    root.destroy()
    return file_path
