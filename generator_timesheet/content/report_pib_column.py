"""Колонка ПІБ у звітах розбіжностей (mis_match.xlsx, addons_mis_match.xlsx,
error_mis_statuses.xlsx, error_mis_statuses_Відомість.xlsx): ПІБ має бути
видно повністю. Раніше колонка мала ширину 30, текст стояв по центру без
переносу, і довгі ПІБ обрізались з обох боків - прізвища не було видно.

Ширина рахується за найдовшим ПІБ звіту: Times New Roman 14 ширший за шрифт,
у символах якого Excel міряє колонку, а великі літери (прізвище) ще ширші за
малі. Вирівнювання ліворуч без переносу (перенос у рядках висотою 20
користувач відхилив - текст накладався): якщо щось і не вміститься,
обріжеться кінець по батькові, а не прізвище."""

import math

from openpyxl.styles import Alignment

PIB_ALIGNMENT = Alignment(horizontal="left", vertical="center", wrap_text=False)

# Одиниць ширини колонки Excel на символ Times New Roman 14.
_UPPER_WIDTH = 1.8
_LOWER_WIDTH = 1.35
_OTHER_WIDTH = 0.7
_PADDING = 2
_MIN_WIDTH = 30


def _text_width(text):
    return sum(_UPPER_WIDTH if ch.isupper() else _LOWER_WIDTH if ch.isalpha() else _OTHER_WIDTH for ch in text)


def pib_column_width(names):
    """Ширина колонки, у яку вміщається найдовше з names (порожні
    пропускаються), але не менша за _MIN_WIDTH."""
    widths = [_text_width(" ".join(str(name).split())) for name in names if name]
    return max(_MIN_WIDTH, math.ceil(max(widths, default=0)) + _PADDING)
