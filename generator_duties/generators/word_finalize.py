"""Доведення книги через Word (COM, pywin32): оновити поля (номери сторінок у
ЗМІСТі), порахувати сторінки й вписати кількість аркушів у напис про зшивання.
python-docx сам сторінок не рахує - це вміє лише Word.

Без Word (немає pywin32/Word, помилка COM) - finalize_without_word: маркер
замінюється на "____ (____) аркушів", а Word запропонує оновити ЗМІСТ під час
відкриття (Ctrl+A, F9 - вручну)."""
import os

from docx import Document

from constants import SHEETS_BLANK, SHEETS_MARKER
from utils.docx_utils import enable_update_fields_on_open
from utils.number_words import sheets_text

WD_STATISTIC_PAGES = 2
WD_FIND_CONTINUE = 1
WD_REPLACE_ALL = 2


class FinalizeError(Exception):
    pass


def sheets_for_pages(pages, duplex):
    """Аркуші: одна сторінка на аркуш, або дві при друці з двох боків."""
    return (pages + 1) // 2 if duplex else pages


def _open_word():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    # Окремий екземпляр Word - не чіпає документів, відкритих користувачем.
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word, pythoncom.CoUninitialize


def finalize_with_word(path, duplex=False, open_word=None):
    """→ (сторінок, аркушів). FinalizeError, якщо Word недоступний чи впав."""
    try:
        word, release = (open_word or _open_word)()
    except Exception as e:
        raise FinalizeError(f"Word недоступний ({e})") from e
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(path), AddToRecentFiles=False)
        doc.Repaginate()
        doc.Fields.Update()
        pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        sheets = sheets_for_pages(pages, duplex)
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        # Лише позиційно: іменовані ReplaceWith/Replace через пізнє зв'язування
        # COM Word мовчки ігнорує (знаходить, але не замінює).
        find.Execute(SHEETS_MARKER, True, False, False, False, False, True, WD_FIND_CONTINUE, False,
                     sheets_text(sheets), WD_REPLACE_ALL)
        doc.Fields.Update()
        doc.Save()
        return pages, sheets
    except Exception as e:
        raise FinalizeError(f"Word не зміг оновити '{os.path.basename(path)}' ({e})") from e
    finally:
        try:
            if doc is not None:
                doc.Close(False)
            word.Quit()
        except Exception:
            pass
        release()


def finalize_without_word(path):
    """Маркер → прочерки; ЗМІСТ оновиться, коли Word відкриє файл."""
    doc = Document(path)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if SHEETS_MARKER in run.text:
                run.text = run.text.replace(SHEETS_MARKER, f"{SHEETS_BLANK} аркушів")
    enable_update_fields_on_open(doc)
    doc.save(path)
