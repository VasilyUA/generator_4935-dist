import re

_SURNAME_RE = re.compile(r"^[А-ЯІЇЄҐ][А-ЯІЇЄҐ'’ʼ\-]+$")


def signature_name(pib):
    """"ПРІЗВИЩЕ Ім'я По батькові" → "Ім'я ПРІЗВИЩЕ" (рядок підпису, як у
    зразку). Якщо перше слово не написане великими літерами - текст лишається
    як є (порядок невідомий)."""
    parts = (pib or "").split()
    if len(parts) >= 2 and _SURNAME_RE.match(parts[0]):
        return f"{parts[1]} {parts[0]}"
    return " ".join(parts)
