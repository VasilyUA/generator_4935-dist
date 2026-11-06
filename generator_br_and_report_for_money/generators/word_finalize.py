"""Доведення рапорту через Word (COM, pywin32): порахувати сторінки й вписати
їх кількість у "Рапорт подається на ... аркушах." python-docx сам сторінок не
рахує (немає реальної пагінації) - це вміє лише сам Word. Той самий принцип,
що й generator_duties/generators/word_finalize.py (інша фраза - родовий, а не
називний відмінок, тож своя таблиця чисел - utils/number_words.py; і без
"аркушів на двох сторонах" - тут аркуш завжди рівно одна сторінка).

Без Word (немає pywin32/Word, помилка COM) - finalize_report_without_word:
маркер замінюється на REPORT_SHEETS_BLANK ("____") - той самий текст, що тут
був статично ДО цієї функції, тож рапорт і далі генерується, просто без
автоматичного числа аркушів (дозаповнюється вручну, як і раніше)."""
import os

from docx import Document

from utils.logging_utils import print_yellow
from utils.number_words import sheets_locative_phrase

REPORT_SHEETS_MARKER = "{{АРКУШІВ_РАПОРТУ}}"
REPORT_SHEETS_BLANK = "____"

WD_STATISTIC_PAGES = 2
WD_FIND_CONTINUE = 1
WD_REPLACE_ALL = 2


class FinalizeError(Exception):
    pass


def _open_word():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    # Окремий екземпляр Word - не чіпає документів, відкритих користувачем.
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word, pythoncom.CoUninitialize


def finalize_report_with_word(path, open_word=None):
    """-> кількість сторінок. FinalizeError, якщо Word недоступний чи впав."""
    try:
        word, release = (open_word or _open_word)()
    except Exception as e:
        raise FinalizeError(f"Word недоступний ({e})") from e
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(path), AddToRecentFiles=False)
        doc.Repaginate()
        pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        # Лише позиційно: іменовані ReplaceWith/Replace через пізнє зв'язування
        # COM Word мовчки ігнорує (знаходить, але не замінює) - той самий
        # нюанс, що й у generator_duties/generators/word_finalize.py.
        find.Execute(REPORT_SHEETS_MARKER, True, False, False, False, False, True, WD_FIND_CONTINUE, False,
                     sheets_locative_phrase(pages), WD_REPLACE_ALL)
        doc.Save()
        return pages
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


def finalize_report_without_word(path):
    """Маркер -> прочерк (REPORT_SHEETS_BLANK)."""
    doc = Document(path)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if REPORT_SHEETS_MARKER in run.text:
                run.text = run.text.replace(REPORT_SHEETS_MARKER, REPORT_SHEETS_BLANK)
    doc.save(path)


def finalize_report_sheet_count(path):
    """Пробує Word (точна кількість сторінок); якщо недоступний чи впав - не
    ламає генерацію рапорту, лише лишає прочерк (finalize_report_without_word)
    з жовтим попередженням у термінал."""
    try:
        finalize_report_with_word(path)
    except FinalizeError as e:
        print_yellow(f"⚠ {e} - кількість аркушів у підписі лишається порожньою, заповніть вручну.")
        finalize_report_without_word(path)
