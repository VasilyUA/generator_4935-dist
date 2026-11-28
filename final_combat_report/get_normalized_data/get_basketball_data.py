import re
from collections import defaultdict
from constants import FILE_PATH_DATA_IN_FOLDER_ARTILLERY
from helpers import get_signal_data, clean_text
from filters import get_filtered_messages_for_date, get_filtered_body_from_cost_messages


def get_basketball_data(hour_of_report, selected_date):
    messages = get_signal_data(FILE_PATH_DATA_IN_FOLDER_ARTILLERY)
    messages = get_filtered_messages_for_date(messages, hour_of_report, selected_date)
    messages = get_filtered_body_from_cost_messages(messages)
    json_data = [get_parsed(b) for b in messages]

    return normalize_data(json_data)


# прості поля "мітка: значення до кінця рядка" - шукаються однаково, без пост-обробки
# рік у "Дата:" - завжди 2-значний у реальних донесеннях ВГрМ ("Дата: 02.08.26")
_SIMPLE_FIELDS = [
    ("Дата", r'Дата:\s*(\d{1,2}\.\d{2}\.\d{2,4})'),
    ("Час", r'Час:\s*([\d:]+(?:-[\d:]+)?)'),
]

# "ВГрМ 1.1""Пежо {UNIT_BRIGADE}" - позивний позиції в реальних повідомленнях
# набирають двома прямими лапками замість « »/" " навколо назви, без
# закриваючої лапки ("1.1""Пежо" замість "1.1 "Пежо""). Ловимо `""СЛОВО` і
# вставляємо пробіл перед відкриваючою й закриваючу лапку після слова.
_BAD_QUOTE_RE = re.compile(r'""(\S+)')

# той самий "N БМП ... N ОБрМП", що й Підрозділ, часто дублюється в кінці
# першого рядка (разом із позивним позиції) - прибираємо звідти, бо він уже
# показаний окремим полем Підрозділ.
_TRAILING_UNIT_RE = re.compile(r'\s*(?:\d+\s*БМП\s*)?\d+\s*ОБрМП\s*$', re.IGNORECASE)

def get_parsed(text):
    record = {"text": text.strip()}

    lines = [x.strip() for x in text.split("\n") if x.strip()]

    for field, pattern in _SIMPLE_FIELDS:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            record[field] = m.group(1).strip()

    unit_match = re.search(r'(\d+\s*БМП)\s*,?\s*(\d+\s*ОБрМП)', text, re.IGNORECASE)
    if unit_match:
        record["Підрозділ"] = f"{unit_match.group(1).lower()} {unit_match.group(2).lower()}"

    if lines:
        pozytsiya = _TRAILING_UNIT_RE.sub('', lines[0])
        pozytsiya = _BAD_QUOTE_RE.sub(r' "\1"', pozytsiya).strip()
        record["ПОЗИЦІЯ"] = pozytsiya

    # "Характер цілі:" тягнеться до "Ціль:" і по дорозі часто захоплює й
    # координати (реальний формат кладе їх МІЖ цими двома мітками, на
    # окремому рядку) - clean_text згортає перенос рядка в пробіл
    kharakter_match = re.search(r'Характер цілі:\s*(.+?)(?=\n\s*Ціль:|$)', text, re.IGNORECASE | re.DOTALL)
    if kharakter_match:
        record["Характер цілі"] = clean_text(kharakter_match.group(1))

    # "Ціль:" - код і, якщо одразу за ним (можливо з наступного рядка) йде
    # опис району в дужках, додаємо його теж ("020840101 (в р-н н.п. Філія)")
    ciль_match = re.search(r'Ціль:\s*(\d+)\s*\n?\s*(\([^)]+\))?', text, re.IGNORECASE)
    if ciль_match:
        code = ciль_match.group(1).strip()
        area = ciль_match.group(2)
        record["Ціль"] = f"{code} {area}" if area else code

    spend_match = re.search(r'Витрат[аи]:\s*(.+?)(?:Результат:|Стрім:|Стрим:|$)', text, re.IGNORECASE | re.DOTALL)
    if spend_match:
        spend_text = clean_text(spend_match.group(1))

        # реальний формат: "ОФ-843Б, 2шт. зар.6, М12, 2шт. (120мм)" - перше
        # "назва, Nшт." це власне міна, "зар.N" і друге "назва, Nшт." - заряд
        # (метальний), що йде РАЗОМ з тією ж міною й не рахується окремо;
        # калібр у дужках наприкінці додається до назви для відображення
        primary_match = re.match(r'\s*([^\s,]+)\s*,\s*(\d+)\s*шт\.?', spend_text)
        caliber_match = re.search(r'\((\d+\s*мм)\)', spend_text)

        if primary_match:
            name = primary_match.group(1)
            qty = int(primary_match.group(2))
            caliber = caliber_match.group(1).replace(" ", "") if caliber_match else ""
            display_name = f"{name} {caliber}".strip()
            record["Витрата"] = f"{display_name}. -{qty}шт."
            record["spend"] = {display_name: qty}
        else:
            # не розпізнано - лишаємо сирий текст лише для відображення в
            # 3.3.6, у підрахунок (spend) не потрапляє, щоб не рахувати "1"
            # для того, що насправді не вдалося розпарсити
            record["Витрата"] = spend_text

    result_match = re.search(r'Результат:\s*(.+?)(?:Стрім:|Стрим:|$)', text, re.IGNORECASE | re.DOTALL)
    if result_match:
        record["Результат"] = result_match.group(1).strip()

    # "Стрім: Без стріму" - на відміну від УБпАК-формату "Стрим:" тут "без"
    # уже й так означає "немає трансляції" саме такими словами, як написано -
    # переписувати на "Відсутній" не треба, лишаємо як є
    stream_match = re.search(r'(?:Стрім|Стрим|Пілот):\s*(.+)', text, re.IGNORECASE)
    if stream_match:
        record["Стрим"] = stream_match.group(1).strip()

    return record


def get_messages(data):
    parts = []

    time_date = f"{data.get('Час', '')} {data.get('Дата', '')}".strip()
    if time_date:
        parts.append(time_date)

    header = f"{data.get('ПОЗИЦІЯ', '')} {data.get('Підрозділ', '')}".strip()
    if header:
        parts.append(f"{header}.")

    if data.get('Характер цілі'):
        parts.append(f"Характер цілі: {data['Характер цілі']}.")

    if data.get('Ціль'):
        parts.append(f"Ціль: {data['Ціль']}.")

    if data.get('Витрата'):
        parts.append(f"Витрата: {data['Витрата']}")

    if data.get('Результат'):
        parts.append(f"Результат: {data['Результат']}.")

    if data.get('Стрим'):
        parts.append(f"Стрім: {data['Стрим']}.")

    return " ".join(parts).strip()

def normalize_data(json_data):
    return [{
        "time": f"{b.get('Дата', '')} {b.get('Час', '')}".strip(),
        "date": b.get('Дата', ''),
        "spend": b.get('spend', {}),
        "text": get_messages(b)} 
        for b in json_data]

def calculate_basketball_data(data):
    ammo_stats = defaultdict(int)

    for item in data:
        for name, qty in item.get("spend", {}).items():
            ammo_stats[name] += qty

    return ", ".join(f"{ammo}. – {qty} шт." for ammo, qty in ammo_stats.items())
