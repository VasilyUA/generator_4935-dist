from datetime import datetime, timedelta
from helpers import clean_text

def get_filtered_messages_for_date(messages, hour_of_report, report_date):
    """
    report_date = '27.04.2026'

    Поверне дані за звітну добу:
    з 26.04.2026 18:00:00
    по 27.04.2026 18:00:00
    """
    if report_date is None:
        report_date = datetime.now().strftime("%d.%m.%Y")

    # захист: беремо тільки перші 10 символів "dd.mm.yyyy"
    report_date = report_date.strip()[:10]

    try:
        report_day = datetime.strptime(report_date, "%d.%m.%Y")
    except ValueError as e:
        raise ValueError(f"Невірний формат дати: '{report_date}'. Очікується 'dd.mm.yyyy'") from e

    # якщо годину задано як 24 - доба закінчується опівночі наступного дня (datetime не приймає hour=24)
    if hour_of_report == 24:
        end = report_day + timedelta(days=1)
    else:
        end = report_day.replace(hour=hour_of_report, minute=0, second=0, microsecond=0)

    start = end - timedelta(days=1)

    filtered = []
    for msg in messages:
        d = msg.get("date")
        if not d:
            continue
        msg_date = datetime.fromisoformat(d) if isinstance(d, str) else d
        if start <= msg_date < end:
            filtered.append(msg)

    return sorted(filtered, key=lambda x: x["date"])

def _nonempty_bodies(messages):
    return [m.get('body') for m in messages if clean_text(m.get('body', "")) != ""]

def get_filtered_body_from_cost_messages(messages, need_spend=True):
    body = _nonempty_bodies(messages)

    if need_spend:
        return [b for b in body if 'Результат' in b or 'Витрата' in b or 'Витрати' in b or 'Знищено' in b]
    return body

def get_filtered_body_from_logistics_uav_messages(messages):
    return [b for b in _nonempty_bodies(messages) if 'логістика' in b.lower() and 'звіт' in b.lower()]

def get_filtered_body_from_loss_uav_messages(messages):
    return [b for b in _nonempty_bodies(messages) if 'виліт' in b.lower() and 'логістичне' in b.lower()]

def get_filtered_body_from_kasper_logistics_messages(messages):
    """Екіпаж "Каспер" звітує іншим форматом ("...бомбер-вампір... працював на
    логістику...") - не містить ні "логістика"+"звіт" (get_filtered_body_from_
    logistics_uav_messages), ні "логістичне" (get_filtered_body_from_loss_uav_
    messages), тому без окремого фільтра ці вильоти повністю випадали з підрахунку."""
    return [b for b in _nonempty_bodies(messages) if 'бомбер-вампір' in b.lower() and 'логістику' in b.lower()]