"""Число словами українською (чоловічий рід - для "аркуш") і правильна форма
слова "аркуш" - для напису про зшивання."""

_UNITS = ("", "один", "два", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
_UNITS_FEMININE = ("", "одна", "дві", "три", "чотири", "п’ять", "шість", "сім", "вісім", "дев’ять")
_TEENS = (
    "десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п’ятнадцять",
    "шістнадцять", "сімнадцять", "вісімнадцять", "дев’ятнадцять",
)
_TENS = ("", "", "двадцять", "тридцять", "сорок", "п’ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев’яносто")
_HUNDREDS = ("", "сто", "двісті", "триста", "чотириста", "п’ятсот", "шістсот", "сімсот", "вісімсот", "дев’ятсот")


def plural_form(number, one, few, many):
    """Форма іменника за числом: 1, 21 → one; 2-4, 22-24 → few; решта → many."""
    if number % 10 == 1 and number % 100 != 11:
        return one
    if 2 <= number % 10 <= 4 and not 12 <= number % 100 <= 14:
        return few
    return many


def _below_thousand(number, units):
    words = []
    hundreds, rest = divmod(number, 100)
    if hundreds:
        words.append(_HUNDREDS[hundreds])
    if 10 <= rest <= 19:
        words.append(_TEENS[rest - 10])
    else:
        tens, unit = divmod(rest, 10)
        if tens:
            words.append(_TENS[tens])
        if unit:
            words.append(units[unit])
    return words


def number_to_words(number):
    """0 < number < 1 000 000 → "сто двадцять п’ять"."""
    if not 0 < number < 1_000_000:
        raise ValueError(f"Число {number} поза межами 1..999999")
    thousands, rest = divmod(number, 1000)
    words = []
    if thousands:
        words += _below_thousand(thousands, _UNITS_FEMININE)
        words.append(plural_form(thousands, "тисяча", "тисячі", "тисяч"))
    words += _below_thousand(rest, _UNITS)
    return " ".join(words)


def sheets_text(number):
    """125 → "125 (сто двадцять п’ять) аркушів"."""
    return f"{number} ({number_to_words(number)}) {plural_form(number, 'аркуш', 'аркуші', 'аркушів')}"
