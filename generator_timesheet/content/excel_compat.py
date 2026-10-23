"""load_workbook (openpyxl), стійкий до текстових значень користувацьких
фільтрів - усі книги проєкту відкриваються звідси.

openpyxl 3.1.2 приймає в <customFilter val="..."> автофільтра лише число чи
шаблон із "*" і на будь-якому іншому тексті відмовляється відкривати книгу
ВЗАГАЛІ. А саме так Excel зберігає звичайний фільтр "(Непусті)" (notEqual
" ") чи "дорівнює/не дорівнює <текст>". Реальний випадок: фільтр на колонці
аркуша "Табель" resources/ОБЛІК.xlsx - ValueError "Value must be either
numerical or a string containing a wildcard" ще до читання жодної клітинки.
Такі значення лишаються текстом - фільтр зберігається у файлі як був."""

from openpyxl import load_workbook
from openpyxl.worksheet.filters import CustomFilter, CustomFilterValueDescriptor

__all__ = ["load_workbook"]


class _TextTolerantFilterValue(CustomFilterValueDescriptor):
    def __set__(self, instance, value):
        if isinstance(value, str) and not self.pattern.match(value):
            instance.__dict__[self.name] = value
            return
        super().__set__(instance, value)


CustomFilter.val = _TextTolerantFilterValue(name="val")
