import re
from datetime import datetime, timedelta
from constants import FILE_PATH_DATA_IN_FOLDER_SHELLING
from helpers import get_signal_data
from filters import get_filtered_messages_for_date


def get_paintball_data(hour_of_report, selected_date):
    """Зчитує JSON і повертає оброблений список повідомлень, пов'язаних з UAV."""
    messages = get_signal_data(FILE_PATH_DATA_IN_FOLDER_SHELLING)
    messages = get_filtered_messages_for_date(messages, hour_of_report, selected_date)

    return get_sorted_messages(messages)


def _normalize_time(t):
    t = t.replace(".", ":")
    return t if len(t) == 5 else "0" + t

def get_sorted_messages(messages):
    temp_list = []

    for msg in messages:
        body = msg.get("body", "").strip()
        raw_date = msg.get("date")

        if not body or not raw_date or "Уточнення" in body:
            continue

        try:
            msg_date = datetime.fromisoformat(raw_date)
        except:
            continue

        # шукаємо час (або діапазон часу) на початку: 7:55 / 07:55 / 14.45 / 09:30-09:35
        time_match = re.match(r"^\s*(\d{1,2}[:.]\d{2})(?:\s*-\s*(\d{1,2}[:.]\d{2}))?", body)

        if time_match:
            start_time = _normalize_time(time_match.group(1))
            end_time = _normalize_time(time_match.group(2)) if time_match.group(2) else None
            time_str = f"{start_time}-{end_time}" if end_time else start_time
            content = body[time_match.end():].strip(" ,.-")
        else:
            start_time = "00:00"
            time_str = "00:00"
            content = body

        try:
            hh, mm = map(int, start_time.split(":"))
        except:
            hh, mm = 0, 0

        # визначаємо реальну дату події:
        # якщо повідомлення прийшло вночі (00:00-05:59)
        # і час події вечірній (18-23) — подія була попереднього дня
        msg_hour = msg_date.hour
        if 0 <= msg_hour <= 5 and 18 <= hh <= 23:
            event_date = msg_date - timedelta(days=1)
        else:
            event_date = msg_date

        formatted_date = event_date.strftime("%d.%m.%Y")

        try:
            sort_key = msg_date.replace(hour=hh, minute=mm, second=0, microsecond=0)
        except:
            sort_key = msg_date.replace(hour=0, minute=0, second=0, microsecond=0)

        formatted_text = (
            f"{time_str} {formatted_date} "
            f"{content}. Втрати уточнюються."
        )

        temp_list.append((sort_key, formatted_text))

    temp_list.sort(key=lambda x: x[0])

    return [x[1] for x in temp_list]


# "бмп" як звичайний підрядок збігається з "{UNIT_BATTALION} {UNIT_BRIGADE}" - штатним
# позначенням СВОГО підрозділу, що є практично в кожному повідомленні цього
# каналу (не має жодного стосунку до обстрілу ворожою БМП). Тому для цієї
# категорії - окремий regex з запереченням "цифра[+пробіл] перед бмп".
_BMP_ATTACK_RE = re.compile(r'(?<!\d)(?<!\d )бмп')

# категорія обстрілу -> ключові слова (рядок - підрядок; скомпільований regex -
# перевіряється через .search), за якими вона розпізнається в тексті повідомлення
_ATTACK_KEYWORDS = [
    ("ракетних_ударів", ("ракет",)),
    ("авіаційних_ударів", ("авіа", "каб")),
    ("артилерійських_обстрілів", ("арт",)),
    ("рсзв", ("рсзв",)),
    ("гранатомети", ("гранатомет",)),
    ("скиди_бпла", ("mavic", "мавік")),
    ("удари_FPV_дронів", ("fpv", "фпв")),
    ("БМП", (_BMP_ATTACK_RE,)),
    ("Підриви", ("свп",)),
    ("танковий", ("танк",)),
    ("ТОС", ("тос",)),
    ("БПЛА_Крило_ЛАНЦЕТ", ("крила", "ланцет")),
    ("снайпер", ("снайпер",)),
    ("стріл_зброя", ("стрілецьк", "зі зброї", "із зброї")),
    ("фосфор", ("фосфор",)),
    ("гранати", ("гранат",)),
]

def _keyword_matches(keyword, text):
    return keyword.search(text) is not None if hasattr(keyword, "search") else keyword in text

# "стрілецький бій" здебільшого - наша власна пошукова група знайшла й
# знищила ворога без власних втрат (це не "обстріл", про які звітує п.3.1.2 -
# там ідеться про удари ПРОТИВНИКА по нас). Тому в "стріл. зброя" рахуємо
# лише випадки, де прямо НЕ сказано "без втрат"/"немає" серед наших військ -
# тобто є реальні втрати (чи вони описані іншим формулюванням, без мітки
# "Втрати наших військ:" узагалі, як у деяких давніших повідомленнях).
_NO_OWN_LOSSES_RE = re.compile(r'втрати\s+наших\s+військ\s*:?\s*(?:без\s+втрат|нема[єc])', re.IGNORECASE)

def _has_no_own_losses_statement(text):
    return _NO_OWN_LOSSES_RE.search(text) is not None

def classify_attacks(data_paintball):
    stats = {key: {"count": 0, "messages": []} for key, _ in _ATTACK_KEYWORDS}

    for msg in data_paintball:
        text = msg.lower()
        if "уточнення" in text:
            continue
        for key, keywords in _ATTACK_KEYWORDS:
            if key == "стріл_зброя" and _has_no_own_losses_statement(text):
                continue
            if any(_keyword_matches(kw, text) for kw in keywords):
                stats[key]["count"] += 1
                stats[key]["messages"].append(msg)

    stats["всього"] = {"count": sum(s["count"] for s in stats.values())}
    return stats

