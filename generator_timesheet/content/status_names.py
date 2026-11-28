"""Перейменовані статуси: код статусу (яким його знають правила програми -
constants.py, utils/default_statuses.py) <-> назва, яку користувач дав йому
на аркуші "Статуси" книги налаштувань.

Правила (розпізнавання рапортів, звірка з книгами addons і файлами БЧС,
виплати) посилаються на статус за кодом, тож уся логіка працює лише з кодами.
Назва з'являється лише на межі програми:
- to_code - на вході: клітинки ОБЛІК.xlsx (Timesheet), значення файлів БЧС
  (content/information_unit_reader.py) - і нова назва, і сам код дають код;
- to_display - на виході: клітинки збереженого ОБЛІК.xlsx (Timesheet.save) і
  звітів (payment_mismatch_checker._display_value).
Нічого не перейменовано - обидві функції повертають значення як є."""
from constants import STATUS_NAMES


def _norm(text):
    return " ".join(text.split()).upper()


# {код: назва} і {нормалізована назва: код} - лише перейменовані статуси.
_DISPLAY = dict(STATUS_NAMES)
_CODES = {_norm(name): code for code, name in _DISPLAY.items()}


def to_code(value):
    """Назва статусу (без урахування регістру й зайвих пробілів) -> його код;
    будь-яке інше значення (зокрема сам код) - як є."""
    if isinstance(value, str) and _CODES:
        return _CODES.get(_norm(value), value)
    return value


def to_display(value):
    """Код статусу -> назва, яку дав йому користувач; складене "А/Б" (кілька
    статусів у звіті) - частинами. Решта значень - як є."""
    if not isinstance(value, str) or not _DISPLAY:
        return value
    if value in _DISPLAY:
        return _DISPLAY[value]
    if "/" in value:
        parts = value.split("/")
        if any(part in _DISPLAY for part in parts):
            return "/".join(_DISPLAY.get(part, part) for part in parts)
    return value
