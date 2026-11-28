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
