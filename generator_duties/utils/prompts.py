from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt


def ask_yes_no(message, default):
    """Копія з generator_bn_and_map/utils/prompts.py."""
    return ListPrompt(
        message=message,
        choices=[
            {"name": "Ні", "value": False},
            {"name": "Так", "value": True},
        ],
        default=default,
    ).execute()


def ask_text(message, default="", validate=None, invalid_message="❌ Некоректне значення."):
    kwargs = {"message": message, "default": str(default)}
    if validate is not None:
        kwargs.update(validate=validate, invalid_message=invalid_message)
    return InputPrompt(**kwargs).execute().strip()
