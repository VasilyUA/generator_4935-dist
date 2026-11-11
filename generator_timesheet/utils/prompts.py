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


def _is_valid_date_text(text):
    try:
        datetime.strptime(text, "%d.%m.%Y")
    except (TypeError, ValueError):
        return False
    return True


def ask_date(message, default):
    """Текстове поле дд.мм.рррр; Enter лишає default. Повертає datetime.date,
    як і решта дат проєкту."""
    text = InputPrompt(
        message=message,
        default=default.strftime("%d.%m.%Y"),
        validate=_is_valid_date_text,
        invalid_message="❌ Дата має бути у форматі дд.мм.рррр.",
    ).execute()
    return datetime.strptime(text, "%d.%m.%Y").date()
