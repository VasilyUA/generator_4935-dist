"""Синтетична книга ЦВС (resources/addons) для тестів - заголовки як у
реальній книзі; на аркуші безвісті над заголовком є рядок-назва (заголовок у
рядку 2), на аркуші загиблих під заголовком - рядок номерів колонок."""

from openpyxl import Workbook

VLK_SHEET = "ВЛК "
VLK_HEADER = ["№ З/П", "ПІБ", "Звання", "Дата Направлення", "Дата проходження ВЛК",
              "Номер ВЛК/Номер протоколу/Свідоцтво про хворобу", "Рішення влк"]
ZVFN_SHEET = "Звільнення від фіз навантажень "
ZVFN_HEADER = ["№З/п", "ПІБ", "Дата початку", "Дата завершення", "№ довідки", "Примітки"]
MEDICAL_SHEET = "СОМАТИКА-ПОРАНЕНІ"
MEDICAL_HEADER = [None, "№ П/П", "Підрозділ", "Військове званння", "Прізвище ім'я по батькові",
                  "соматика / травмування / поранення", "Дата Надходження в мед заклад", "Лікувальний заклад",
                  "Підстава( № Госпіталізації )", "Дата виписки"]
MISSING_SHEET = "БЕЗВЕСТІ"
MISSING_HEADER = ["№ Посади", "Посада", "Підрозділ", "Звання", "ПІБ", "Дата зарахування в частини", "Статус",
                  "Дата зникнення"]
CAPTIVITY_SHEET = "В полоні"
CAPTIVITY_HEADER = ["№З/п", "Звання", "ПІБ", "Рік народження", "Дата в полоні", "Дата звільнення з полону"]
KILLED_SHEET = "Загиблі-200"
KILLED_HEADER = [None, "№ з/п", "Військове звання", "ПІБ", "Дата втрат (смерті) особового складу",
                 "лікарське свідоцтво про смерть (номер, дата)"]


def write_cvs(path, rows_by_sheet):
    """rows_by_sheet: {назва аркуша: [рядок даних, ...]} - кожен аркуш з
    заголовком у тій самій позиції, що й у реальній книзі; аркуші, яких немає
    в rows_by_sheet, не створюються (крім "map")."""
    headers = {
        VLK_SHEET: VLK_HEADER, ZVFN_SHEET: ZVFN_HEADER, MEDICAL_SHEET: MEDICAL_HEADER,
        MISSING_SHEET: MISSING_HEADER, CAPTIVITY_SHEET: CAPTIVITY_HEADER, KILLED_SHEET: KILLED_HEADER,
    }
    wb = Workbook()
    wb.worksheets[0].title = "map"
    wb.worksheets[0].append(['"ВЛК" — Військово лікарська комісія', "ПІБ, Дата проходження ВЛК"])
    for sheet_name, rows in rows_by_sheet.items():
        ws = wb.create_sheet(sheet_name)
        if sheet_name == MISSING_SHEET:
            ws.append([None, None, None, None, "Корпус A0000"])
        ws.append(headers[sheet_name])
        if sheet_name == KILLED_SHEET:
            ws.append([None] + list(range(1, len(KILLED_HEADER))))
        for row in rows:
            ws.append(row)
    wb.save(str(path))
    return str(path)
