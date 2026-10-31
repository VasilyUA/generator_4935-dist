import os
import time

from utils.logging_utils import print_red


def _close_excel_workbook_if_open(file_path):
    """Якщо САМЕ цей файл зараз відкритий у Excel (напр. користувач переглядав
    попередній результат) - закриває його БЕЗ збереження, щоб openpyxl зміг
    перезаписати файл. Той самий підхід, що й formatting.docx_utils.
    _close_word_document_if_open для .docx (win32com.client.GetActiveObject) -
    тут для Excel.Application. Excel не запущено або pywin32 недоступний -
    тихо виходить (як і докс-версія)."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return

    abs_path = os.path.abspath(file_path)
    try:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            workbooks = list(excel.Workbooks)
        except Exception:
            # Excel не запущено, або запущений екземпляр у нестабільному стані
            # (напр. модальний діалог, захищений перегляд) - COM-виклик до
            # excel.Workbooks може впасти з AttributeError, не лише
            # com_error - той самий "тихо виходь" підхід, що й вище.
            return
        for workbook in workbooks:
            try:
                full_name = os.path.abspath(workbook.FullName)
            except Exception:
                continue
            if full_name == abs_path:
                print_red(f"Закриваю відкритий у Excel файл: {os.path.basename(full_name)}")
                workbook.Close(SaveChanges=False)
    finally:
        pythoncom.CoUninitialize()


def save_workbook_safely(workbook, file_path, attempts=5, delay=0.5):
    """workbook.save(), але спершу закриває цей самий файл у Excel, якщо він
    там відкритий (напр. користувач переглядав попередній результат цього ж
    інструменту), і повторює спробу, якщо файл лишається заблокованим ще
    коротку мить (антивірус/OneDrive) - той самий підхід, що й
    formatting.docx_utils.save_docx_safely, для .xlsx замість .docx."""
    _close_excel_workbook_if_open(file_path)
    for attempt in range(attempts):
        try:
            workbook.save(file_path)
            return
        except PermissionError as e:
            if attempt == attempts - 1:
                raise PermissionError(
                    f"Не вдалось зберегти '{file_path}': файл заблоковано іншою програмою. "
                    f"Закрийте його (напр. Excel) і спробуйте ще раз. ({e})"
                ) from e
            time.sleep(delay)
