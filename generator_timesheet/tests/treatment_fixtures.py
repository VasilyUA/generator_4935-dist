"""Синтетичний журнал обліку військовослужбовців на лікуванні поза межами
військової частини (resources/addons) для тестів - заголовки й рядок номерів
колонок як у реальній книзі."""

from openpyxl import Workbook

TREATMENT_HEADER = [
    "№ з/п", "військова частина", "Військове звання", "Прізвище, ім’я та по батькові", "ІПН", "Первинний діагноз",
    "Дата госпіталізації", "Заклад охорони здоров'я до якого було госпіталізовано",
    "До якого закладу було переведено та дата переводу", "Дата виписки",
]


def write_treatment_journal(path, rows_by_sheet):
    """rows_by_sheet: {аркуш: [(ПІБ, дата госпіталізації, дата виписки), ...]}."""
    wb = Workbook()
    wb.remove(wb.worksheets[0])
    for sheet_name, rows in rows_by_sheet.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(TREATMENT_HEADER)
        ws.append(list(range(1, len(TREATMENT_HEADER) + 1)))
        for number, (pib_raw, start, end) in enumerate(rows, start=1):
            ws.append([number, "A0000", "матрос", pib_raw, None, "діагноз", start, "лікарня", None, end])
    wb.save(str(path))
    return str(path)
