import re
from datetime import datetime

from constants import FILE_PATH_DATA_IN_FOLDER_COMMISSION
from helpers import get_signal_data, clean_text
from filters import get_filtered_messages_for_date


def get_data_commission(hour_of_report, selected_date):
    """Зчитує JSON групи КОМІСІЇ і повертає список підготовлених абзаців для 3.4."""
    messages = get_signal_data(FILE_PATH_DATA_IN_FOLDER_COMMISSION)
    messages = get_filtered_messages_for_date(messages, hour_of_report, selected_date)

    return get_normalized_data(messages)


def get_normalized_data(messages):
    normalized_list = []

    for msg in messages:
        text = clean_text(_strip_service_prefixes(msg.get("body", "")))
        if not text or not re.search(r"[a-zA-Zа-яА-ЯіІїЇєЄ]", text):
            continue

        dt = _parse_message_date(msg.get("date"))

        # не дублюємо час, якщо повідомлення вже починається з дати/часу
        if not re.match(r"^\d{1,2}[:.]\d{2}", text) and dt is not None:
            text = f"{dt.strftime('%H:%M %d.%m.%Y')} {text}"

        normalized_list.append({"date": dt or datetime.min, "text": text})

    normalized_list.sort(key=lambda x: x["date"])

    return [item["text"] for item in normalized_list]


def _strip_service_prefixes(body):
    if not isinstance(body, str):
        return body

    return re.sub(r"^\s*Пропишіть\s*:?\s*", "", body, flags=re.IGNORECASE)


def _parse_message_date(date_val):
    if isinstance(date_val, datetime):
        return date_val
    if isinstance(date_val, str):
        try:
            return datetime.fromisoformat(date_val)
        except ValueError:
            return None
    return None
