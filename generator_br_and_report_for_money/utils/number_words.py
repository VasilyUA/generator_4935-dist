"""Число словами українською у РОДОВОМУ відмінку (узгоджується з "аркушах" у
фразі "Рапорт подається на ... аркушах" - родовий і місцевий відмінки
числівників тут збігаються: "133" -> "ста тридцяти трьох"). Окрема таблиця
від generator_duties/utils/number_words.py (там - називний відмінок, для
іншої фрази "Зшито та пронумеровано ... аркушів")."""

_UNITS_GENITIVE = ("", "одного", "двох", "трьох", "чотирьох", "п’яти", "шести", "семи", "восьми", "дев’яти")
_TEENS_GENITIVE = (
    "десяти", "одинадцяти", "дванадцяти", "тринадцяти", "чотирнадцяти", "п’ятнадцяти",
    "шістнадцяти", "сімнадцяти", "вісімнадцяти", "дев’ятнадцяти",
)
_TENS_GENITIVE = (
    "", "", "двадцяти", "тридцяти", "сорока", "п’ятдесяти", "шістдесяти", "сімдесяти", "вісімдесяти", "дев’яноста",
)
_HUNDREDS_GENITIVE = (
    "", "ста", "двохсот", "трьохсот", "чотирьохсот", "п’ятисот", "шестисот", "семисот", "восьмисот", "дев’ятисот",
)


def _below_thousand_genitive(number):
    words = []
    hundreds, rest = divmod(number, 100)
    if hundreds:
        words.append(_HUNDREDS_GENITIVE[hundreds])
    if 10 <= rest <= 19:
        words.append(_TEENS_GENITIVE[rest - 10])
    else:
        tens, unit = divmod(rest, 10)
        if tens:
            words.append(_TENS_GENITIVE[tens])
        if unit:
            words.append(_UNITS_GENITIVE[unit])
    return words


def number_to_words_genitive(number):
    """0 < number < 1 000 000 -> "ста тридцяти трьох" (133), "ста вісімнадцяти" (118)."""
    if not 0 < number < 1_000_000:
        raise ValueError(f"Число {number} поза межами 1..999999")
    thousands, rest = divmod(number, 1000)
    words = []
    if thousands:
        if thousands == 1:
            words.append("однієї")
        else:
            words += _below_thousand_genitive(thousands)
        words.append("тисячі" if thousands == 1 else "тисяч")
    words += _below_thousand_genitive(rest)
    return " ".join(words)


def sheets_locative_phrase(number):
    """133 -> "133 (ста тридцяти трьох)" - без самого слова "аркушах": у фразі
    "Рапорт подається на ... аркушах" воно незмінне (завжди ця форма,
    незалежно від числа), тож лишається в шаблоні окремо, а не тут."""
    return f"{number} ({number_to_words_genitive(number)})"
