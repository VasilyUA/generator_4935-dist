"""Колонка ПІБ у звітах розбіжностей (content/report_pib_column.py)."""

import math

from content.report_pib_column import PIB_ALIGNMENT, pib_column_width


def test_pib_column_fits_the_longest_name_with_wider_capitals():
    """Рішення користувача: ПІБ - повністю: ширина за найдовшим
    ПІБ (великі літери прізвища ширші за малі), не менше 30; вирівнювання
    ліворуч без переносу - обрізатись може хіба кінець по батькові."""
    assert pib_column_width([None, ""]) == pib_column_width(["ПЕРШИЙ П."]) == 30  # мінімум
    assert pib_column_width(["ПЕРШИЙ Перший Перший"]) == math.ceil(8 * 1.8 + 10 * 1.35 + 2 * 0.7) + 2
    longest = "ВІСІМНАДЦЯТИЙ Вісімнадцятий Вісімнадцятий"  # 15 великих, 24 малі, 2 пробіли
    assert pib_column_width(["ПЕРШИЙ Перший Перший", longest]) == math.ceil(15 * 1.8 + 24 * 1.35 + 2 * 0.7) + 2
    assert pib_column_width(["ВІСІМНАДЦЯТИЙ   Вісімнадцятий Вісімнадцятий"]) == pib_column_width([longest])
    assert (PIB_ALIGNMENT.horizontal, PIB_ALIGNMENT.wrap_text) == ("left", False)
