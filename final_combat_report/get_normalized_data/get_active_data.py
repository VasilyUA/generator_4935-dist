import re
from datetime import datetime
from collections import defaultdict
from constants import FILE_PATH_DATA_IN_FOLDER_ACTIVITY, UNIT_BATTALION, UNIT_COMPANY_ONE, UNIT_COMPANY_TWO, UNIT_COMPANY_DSHR, UNIT_COMPANY_RVP, UNIT_COMPANY_RECONNAISSANCE
from helpers import clean_text, get_signal_data
from filters import get_filtered_messages_for_date


def get_active_data(hour_of_report, selected_date):
    messages = get_signal_data(FILE_PATH_DATA_IN_FOLDER_ACTIVITY)
    messages = get_filtered_messages_for_date(messages, hour_of_report, selected_date)

    bodies = []
    for msg in messages:
        body = msg.get("body", "")
        if clean_text(body) == "":
            continue
        bodies.append(_inject_missing_date(body, msg.get("date")))

    return get_normalized_data(bodies)


def _inject_missing_date(body, msg_date):
    """Підрозділ "мін. батр." у реальних повідомленнях пише лише час, без дати
    (на відміну від інших підрозділів) - get_normalized_data такі рядки просто
    пропускає (немає date-патерну). Підставляємо дату із самого Signal-
    повідомлення (msg["date"]) одразу після часу на початку рядка."""
    if re.search(r"\d{2}[:.]\d{2}\s+\d{2}\.\d{2}\.\d{4}", body):
        return body

    leading_time = re.match(r"^\s*(\d{2}[:.]\d{2})", body)
    if not leading_time or not msg_date:
        return body

    try:
        date_str = datetime.fromisoformat(msg_date).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return body

    return body[:leading_time.end()] + f" {date_str}" + body[leading_time.end():]

def get_normalized_data(data_active):
    normalized_list = []
    
    for text in data_active:
        text = clean_text(text)
        if not text:
            continue

        # шукаємо час з двокрапкою АБО крапкою: "15:49" або "15.49"
        match = re.search(r"(\d{2}[:.]\d{2})\s+(\d{2}\.\d{2}\.\d{4})", text)
        
        if match:
            time_str, date_str = match.groups()
            # нормалізуємо крапку в часі до двокрапки: "15.49" -> "15:49"
            time_str = time_str.replace(".", ":")
            full_date_str = f"{date_str} {time_str}"
            
            try:
                dt_obj = datetime.strptime(full_date_str, "%d.%m.%Y %H:%M")
                normalized_list.append({"date": dt_obj, "text": clean_text(text) })
            except ValueError:
                continue
                
    normalized_list.sort(key=lambda x: x["date"])
    
    return normalized_list

def _normalize_text(text: str) -> str:
    text = clean_text(text).lower()
    text = re.sub(r'(?<=\d)\s*[хХ]\s*(?=\d)', 'x', text)  # кирилична х між цифрами
    text = re.sub(r'(\d)\.(\d)', r'\1,\2', text)            # 5.56 -> 5,56 (тільки між цифрами)
    return text


def _norm(cal: str) -> str:
    # відображення калібру - кирилична "х" і крапка як роздільник, як у реальних
    # донесеннях (5.56х45мм), а не "x"/кома, які використовуються лише для
    # внутрішнього зіставлення з текстом повідомлень (_normalize_text)
    cal = cal.strip()
    cal = re.sub(r'\s+', '', cal)
    cal = cal.replace('мм', '')
    if 'x' not in cal:
        return cal.strip()
    a, b = cal.split('x', 1)
    return f"{a.replace(',', '.')}х{b.replace(',', '.')}мм"


def _unit_in_text(unit_norm: str, text: str) -> bool:
    if unit_norm in text:
        return True
    # реальні повідомлення часто скорочують "мінбатр" як "мін. батр." (крапка +
    # пробіл) - тому прибираємо й крапки, а не лише пробіли, перед порівнянням
    compact = re.sub(r'[\s.]+', '', clean_text(unit_norm))
    text_compact = re.sub(r'[\s.]+', '', text)
    return compact in text_compact


def calculate(data, unit_name):
    ammo_totals = defaultdict(int)
    unit_norm = unit_name.lower().strip()

    for row in data:
        text = _normalize_text(clean_text(row.get("text", "")))

        if not _unit_in_text(unit_norm, text):
            continue

        text_nospace = re.sub(r'\s+', '', text)

        # дшр: рядки РВП містять "дшр/{UNIT_BATTALION без пробілу}" — виключаємо
        if unit_norm == UNIT_COMPANY_DSHR and UNIT_COMPANY_RVP in text:
            continue
        # рв: "рв" - підрядок "рвп", виключаємо щоб не задвоювати з рвп
        if unit_norm == UNIT_COMPANY_RECONNAISSANCE and UNIT_COMPANY_RVP in text:
            continue
        company_two_nospace = UNIT_COMPANY_TWO.replace(" ", "")
        # UNIT_COMPANY_ONE: не брати рядки, що згадують UNIT_COMPANY_TWO
        if unit_norm == UNIT_COMPANY_ONE and company_two_nospace in text_nospace:
            continue
        # UNIT_COMPANY_TWO: брати тільки рядки, що згадують UNIT_COMPANY_TWO
        if unit_norm == UNIT_COMPANY_TWO and company_two_nospace not in text_nospace:
            continue

        found_pairs = []

        # "360 шт (5,56x45мм)" або "360 набоїв (5,56x45мм)"
        for qty, cal in re.findall(
            r'(\d+)\s*(?:шт\.?|набоїв)\s*\(?\s*(\d+[,]\d+\s*x\s*\d+[,]?\d*\s*(?:мм?)?)',
            text
        ):
            found_pairs.append((qty, cal))

        # "5,56x45 мм - 300 набоїв" або "5,56x45 - 300 шт" (мм необов'язково)
        for cal, qty in re.findall(
            r'\(?\s*(\d+[,]\d+\s*x\s*\d+[,]?\d*\s*(?:мм?)?)\)?\s*-\s*(\d+)\s*(?:шт\.?|набоїв)',
            text
        ):
            found_pairs.append((qty, cal))

        # "5,45x39 ПС - 330 шт" — калібр + тип патрона (1-5 букв) + кількість
        for cal, qty in re.findall(
            r'(\d+[,]\d+\s*x\s*\d+[,]?\d*\s*(?:мм?)?)\s+\w{1,5}\s*-\s*(\d+)\s*(?:шт\.?|набоїв)',
            text
        ):
            found_pairs.append((qty, cal))

        # дедуплікація і підрахунок
        used = set()
        for qty, cal in found_pairs:
            normed = _norm(cal)
            key = (normed, qty)
            if key in used:
                continue
            used.add(key)
            ammo_totals[normed] += int(qty)

        # 12x70 окремо
        used_12x70 = set()
        for qty in re.findall(r'12\s*x\s*70[^\d]{0,30}?(\d+)\s*(?:набоїв|шт)', text):
            if qty not in used_12x70:
                used_12x70.add(qty)
                ammo_totals["12х70"] += int(qty)

        # 12 калібр окремо (шт або набоїв, з тире або без)
        used_12 = set()
        for qty in re.findall(
            r'12\s*каліб\w*\s*(?:-\s*)?(?:використано\s*)?(\d+)\s*(?:набоїв|шт)',
            text
        ):
            if qty not in used_12:
                used_12.add(qty)
                ammo_totals["12"] += int(qty)

    if not ammo_totals:
        return ""

    priority = [
        "5.56х45мм",
        "5.45х39мм",
        "12х70",
        "7.62х51мм",
        "7.62х39мм",
        "12",
    ]
    result = sorted(
        ammo_totals.items(),
        key=lambda x: priority.index(x[0]) if x[0] in priority else 999
    )
    return ", ".join(f"{k} – {v} шт." for k, v in result)