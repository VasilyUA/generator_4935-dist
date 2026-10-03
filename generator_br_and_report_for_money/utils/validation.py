import re
from InquirerPy.prompts.input import InputPrompt

from utils.logging_utils import print_red
from constants import HIGHER_COMMANDER_TITLE


def validate_battalion_commander_exists(rows_with_data, file_path=None):
    """Перевіряє, що серед посад (колонка B, 'ПОСАДА') є 'Командир батальйону' —
    від цього залежить find_higher_commander і, відповідно, вся генерація."""
    if not any(row.get('ПОСАДА') == HIGHER_COMMANDER_TITLE for row in rows_with_data):
        location = f" у файлі {file_path}" if file_path else ""
        raise ValueError(
            f"Не знайдено посаду '{HIGHER_COMMANDER_TITLE}' (колонка B){location}. "
            f"Перевірте, що така людина є в списку особового складу."
        )


def get_validated_ksp_data():
    city_pattern = r"^[А-ЩЬЮЯҐЄІЇа-щьюяґєіїA-Za-z\s\-']+$"
    mgrs_pattern = r"^\d{2}[A-Z]\s[A-Z]{2}\s\d{5}\s\d{5}$"

    while True:
        city = InputPrompt(message="Введіть населений пункт КСП (формат: КИЇВ):", default="КИЇВ").execute().strip()
        if not re.match(city_pattern, city):
            print_red("❌ Некоректна назва. Використовуйте тільки літери.")
            continue

        coordinates = InputPrompt(message="Введіть координати КСП (формат: 36T TT 12345 67890):", default="36T TT 12345 67890").execute().strip()
        if not re.match(mgrs_pattern, coordinates):
            print_red("❌ Некоректні координати. Введіть у форматі: 36T TT 12345 67890.")
            continue

        print(f"✅ КСП {city} {coordinates} — запускаємо скрипт...")
        return city, coordinates
