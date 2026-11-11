from datetime import datetime

from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt


def ask_yes_no(message, default):
    return ListPrompt(
        message=message,
        choices=[
            {"name": "Ні", "value": False},
            {"name": "Так", "value": True},
        ],
        default=default,
    ).execute()


def ask_source_priority(default="br"):
    """Що головніше, якщо значення є і в БР, і в ПБД: "br" чи "pbd".
    Попередній бойовий наказ - завжди третій."""
    return ListPrompt(
        message="Пріоритет даних, що є і в БР, і в ПБД (координати КСП/ЗКП, н.п. КСП, видання карти):",
        choices=[
            {"name": "1) Бойове розпорядження → 2) ПБД → 3) попередній бойовий наказ", "value": "br"},
            {"name": "1) ПБД → 2) Бойове розпорядження → 3) попередній бойовий наказ", "value": "pbd"},
        ],
        default=default,
    ).execute()


def _is_valid_date_text(text):
    try:
        datetime.strptime(text, "%d.%m.%Y")
    except (TypeError, ValueError):
        return False
    return True


def ask_date(message, default):
    """Текстове поле дд.мм.рррр (копія generator_timesheet/utils/prompts.py) -
    default (datetime.date) лишається підказкою, якщо просто натиснути Enter."""
    text = InputPrompt(
        message=message,
        default=default.strftime("%d.%m.%Y"),
        validate=_is_valid_date_text,
        invalid_message="❌ Дата має бути у форматі дд.мм.рррр.",
    ).execute()
    return datetime.strptime(text, "%d.%m.%Y").date()


def ask_text(message, default="", validate=None, invalid_message="❌ Некоректне значення."):
    kwargs = {"message": message, "default": str(default)}
    if validate is not None:
        kwargs.update(validate=validate, invalid_message=invalid_message)
    return InputPrompt(**kwargs).execute().strip()
