import re

from constants import FILE_PATH_DATA_IN_FOLDER_UAV, UAV_SPEND_COMPONENT_PREFIXES, UAV_MUNITION_DISPLAY_NAMES, CALLSIGN_ALIASES, AMBIGUOUS_CALLSIGN_OVERRIDES, UNIT_BRIGADE, UNIT_BATTALION
from helpers import get_signal_data, clean_text
from collections import defaultdict
from filters import get_filtered_body_from_loss_uav_messages, get_filtered_messages_for_date, get_filtered_body_from_cost_messages, get_filtered_body_from_logistics_uav_messages, get_filtered_body_from_kasper_logistics_messages

# UNIT_BATTALION/UNIT_BRIGADE ("N бмп"/"N обрмп") зі змінним пробілом між числом
# і абревіатурою (сирі повідомлення пишуть і "1бмп", і "1 бмп") - \ escape-ований
# пробіл замінюємо на \s* замість голого re.escape(...), regex все одно
# case-insensitive (re.I нижче), тож окремий mixed-case варіант тут не потрібен.
_UNIT_BRIGADE_RE_PART = re.escape(UNIT_BRIGADE).replace(r'\ ', r'\s*')
_UNIT_BATTALION_RE_PART = re.escape(UNIT_BATTALION).replace(r'\ ', r'\s*')


def get_UAV_data(hour_of_report, selected_date, rows_with_data):
    """Зчитує JSON і повертає оброблений список повідомлень, пов'язаних з UAV."""
    messages = get_signal_data(FILE_PATH_DATA_IN_FOLDER_UAV)
    messages = get_filtered_messages_for_date(messages, hour_of_report, selected_date)

    logistics_data = get_logistics_data(messages) + get_kasper_logistics_data(messages)
    cost_data = get_cost_data(messages, rows_with_data, selected_date)
    loss_data = get_loss_data(messages, rows_with_data)

    return {
        "cost_data": cost_data,
        "loss_data": loss_data,
        "logistics_data": logistics_data
    }

def get_logistics_data(messages):
    filtered_messages = get_filtered_body_from_logistics_uav_messages(messages)
    results = []

    for msg in filtered_messages:
        text = msg.get("text", "") if isinstance(msg, dict) else str(msg)
        text = text.strip()

        # --- Екіпаж ---
        crew = ""
        m = re.search(r'Звіт по роботі екіпажу\s+"?([^\n"]+)"?', text, re.I)
        if m:
            crew = m.group(1).strip()

        # --- Підрозділ --- clean_text ДО lower(): REPLACEMENTS перетворює написання
        # ВЕЛИКИМИ літерами на UNIT_BATTALION, шукаючи точну (не приведену до
        # нижнього регістру) заміну, тож застосована ПІСЛЯ .lower() вона б просто
        # не спрацювала і лишила б написання злипленим.
        unit = ""
        m = re.search(rf"(ВБНК.*?{_UNIT_BRIGADE_RE_PART}|{_UNIT_BRIGADE_RE_PART}.*?{_UNIT_BATTALION_RE_PART})", text, re.I)
        if m:
            unit = clean_text(m.group(1).strip()).lower()

        # --- Тип НРК / Засіб ---
        nrk_type = ""
        m = re.search(r"Тип НРК:\s*(.+)", text, re.I)
        if m:
            nrk_type = m.group(1).strip()
        else:
            m = re.search(r"Засіб:\s*(.+)", text, re.I)
            if m:
                nrk_type = m.group(1).strip()

        # --- Маршрут ---
        route = ""
        m = re.search(r"Маршрут:\s*(.*?)\n", text, re.S | re.I)
        if m:
            route = clean_text(m.group(1))

        # --- Задачі (напр. "логістика та перегонка борта") ---
        tasks = ""
        m = re.search(r"Задачі:\s*(.+)", text, re.I)
        if m:
            tasks = m.group(1).strip()

        # --- Дата (підтримка 1.05.2026 і 01.05.2026, "План роботи НРК: X-Y",
        # "за X.Y.YY" (однозначна дата, рік може бути 2-значним), і - як
        # останній варіант - дата окремим рядком без жодної мітки) ---
        date = ""

        def pad_date(d):
            parts = d.split(".")
            year = parts[2] if len(parts[2]) == 4 else f"20{parts[2]}"
            return f"{int(parts[0]):02d}.{parts[1]}.{year}"

        m = re.search(r"(?:План роботи НРК|за)\s*:?\s*(\d{1,2}\.\d{2}\.\d{4})\s*[-–]\s*(\d{1,2}\.\d{2}\.\d{4})", text, re.I)
        if m:
            date = f"{pad_date(m.group(1))}-{pad_date(m.group(2))}"
        else:
            m = re.search(r"Дата:\s*(\d{1,2}\.\d{2}\.\d{4})", text)
            if m:
                date = pad_date(m.group(1))
            else:
                m = re.search(r"\bза\s+(\d{1,2}\.\d{2}\.\d{2,4})\b", text, re.I)
                if m:
                    date = pad_date(m.group(1))
                else:
                    m = re.search(r"(?:^|\n)\s*(\d{1,2}\.\d{2}\.\d{4})\s*(?:\n|$)", text)
                    if m:
                        date = pad_date(m.group(1))

        # --- Час виїзду --- (обидві межі окремо доповнюємо нулем до 2 цифр:
        # "6:20-7:10" -> "06:20-07:10")
        departure_time = ""
        m = re.search(r"Час виїзду:\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})", text)
        if m:
            def pad_time(t):
                hh, mm = t.split(":")
                return f"{int(hh):02d}:{mm}"
            departure_time = f"{pad_time(m.group(1))}-{pad_time(m.group(2))}"

        # --- Виїзд / Виліт ---
        departure = ""
        m = re.search(r"(Виїзд\s*\d+\s*(?:\(доставки\s*\d+\))?)", text, re.I)
        if m:
            departure = m.group(1).strip()
        else:
            m = re.search(r"(Виліт\s*\d+)", text, re.I)
            if m:
                departure = m.group(1).strip()

        # --- Вага - зберігаємо оригінальну мітку ("Загальна вага:" чи просто
        # "Вага:") такою, як у сирому повідомленні ---
        weight = ""
        weight_label = "Загальна вага"
        m = re.search(r"Загальна\s*вага:\s*([\d.,]+\s*кг)", text, re.I)
        if m:
            weight = m.group(1).strip()
        else:
            m = re.search(r"Вага:\s*([\d.,]+\s*кг)", text, re.I)
            if m:
                weight = m.group(1).strip()
                weight_label = "Вага"

        # --- Відстань ---
        distance = ""
        m = re.search(r"Відстань:\s*([\d.,]+\s*км)", text, re.I)
        if m:
            distance = m.group(1).strip()

        # --- Логістика (нумеровані пункти) ---
        logistics_items = []
        block = re.search(
            r"Логістика\s*(.*?)(?:Час виїзду:|Загальна вага:|Відстань:|Логістика успішно|$)",
            text,
            re.S | re.I
        )
        if block:
            for line in block.group(1).split("\n"):
                line = line.strip()
                if re.match(r"^\d+\.", line):
                    logistics_items.append(line)

        logistics_text = " ".join(logistics_items).strip()

        # --- Успішність - "Логістика успішно." показуємо лише коли сире
        # повідомлення справді про це звітує ---
        success = "Логістика успішно." if re.search(r"Логістика\s+успішно", text, re.I) else ""

        # --- Формування результату - "Звіт по роботі екіпажу {crew} {unit}."
        # замість окремих міток "Підрозділ:"/"Екіпаж:", бо саме так виглядають
        # реальні звіти ВБНК ("Звіт по роботі екіпажу Змій\nВБНК {UNIT_BATTALION без пробілу, ВЕЛИКИМИ} {UNIT_BRIGADE}...")
        res_str = (
            f"{departure_time} {date} Звіт по роботі екіпажу {crew} {unit}. "
            f"{departure} {logistics_text} "
            f"Тип НРК: {nrk_type}. "
            f"Маршрут: {route}. "
            + (f"Задачі: {tasks}. " if tasks else "")
            + f"{weight_label}: {weight}. "
            f"Відстань: {distance}. "
            + success
        )

        res_str = re.sub(r"\s+", " ", res_str).strip()
        res_str = re.sub(r"\.+", ".", res_str)
        res_str = res_str.replace(" .", ".")

        if res_str and res_str not in results:
            results.append({
                "date": f"{departure_time} {date}",
                "text": res_str
            })

    return results

# Екіпаж "Каспер" звітує повністю іншим шаблоном ("Звіт {UNIT_BATTALION} {UNIT_BRIGADE}:
# бомбер-вампір(Старлінк) / Екіпаж "Каспер" працював на логістику. / дата /
# Виліт N  час_початку/час_кінця  ( позиція ) / Успішно доставлен(о|а) - вантаж").
# В одному Signal-повідомленні буває КІЛЬКА таких блоків підряд (оператор
# копіює попередній звіт і дописує новий) - тому шукаємо ВСІ входження
# "Екіпаж "Каспер"" через finditer, а не лише перше.
_KASPER_BLOCK_RE = re.compile(
    r'Екіпаж\s*"Каспер"[^\n]*\n\s*'
    r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*\n\s*'
    r'Виліт\s*(\d+)\s+(\d{1,2}:\d{2})\s*[/\\]\s*(\d{1,2}:\d{2})\s*\(\s*([^)]+?)\s*\)\s*\n\s*'
    r'(Успішно\s+достав\w*)\s*[-–—]\s*(.+)',
    re.IGNORECASE
)

def get_kasper_logistics_data(messages):
    filtered_messages = get_filtered_body_from_kasper_logistics_messages(messages)
    results = []

    for body in filtered_messages:
        for m in _KASPER_BLOCK_RE.finditer(body):
            day, month, year, flight_num, start_time, end_time, location, verb, cargo = m.groups()
            year4 = year if len(year) == 4 else f"20{year}"
            date_str = f"{int(day):02d}.{int(month):02d}.{year4}"
            cargo = re.sub(r'\s+', ' ', cargo).strip()
            location = location.strip()

            text = (
                f"{start_time}-{end_time} {date_str} Звіт {UNIT_BATTALION} {UNIT_BRIGADE}: бомбер-вампір (Старлінк). "
                f'Екіпаж "Каспер" працював на логістику. Виліт {flight_num}. Позиція: {location}. '
                f"{verb.strip()} - {cargo}"
            )

            results.append({
                "date": f"{start_time} {date_str}",
                "text": clean_text(text)
            })

    return results

def get_loss_data(messages, rows_with_data):
    filtered_messages = get_filtered_body_from_loss_uav_messages(messages)
    result = []

    for text in filtered_messages:
        # \s+ замість пробілу дозволяє часу бути на новому рядку; \d{1,2} для
        # години допускає однозначний запис без нуля попереду ("Дата: ... 7:21"),
        # який раніше не проходив і повідомлення мовчки відкидалось.
        date_match = re.search(r'Дата:\s*(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2})', text)
        if not date_match:
            continue

        date_str = date_match.group(1)
        hour_str, minute_str = date_match.group(2).split(":")
        time_str = f"{int(hour_str):02d}:{minute_str}"

        # Назва екіпажу зі "Звіт по роботі екіпажу "X"" - раніше була жорстко
        # зашита як "Посіпаки", хоча реальні звіти бувають і від інших
        # екіпажів (напр. "Карлсон")
        crew_label_match = re.search(r'Звіт по роботі екіпажу\s+"?([^\n"]+)"?', text, re.I)
        crew_label = crew_label_match.group(1).strip() if crew_label_match else ""

        # Позивні екіпажу - DOTALL + lookahead на наступне поле, бо позивні
        # часто продовжуються на НАСТУПНОМУ рядку ("Позивний: Росомаха,\nМакро,
        # Дед, Пепс") - без DOTALL продовження мовчки губилось.
        crew_match = re.search(
            r'Позивний\s*:\s*(.+?)(?=\n\s*(?:Водій|Засіб|Виліт|Логістичне|Позиція|Результат|Події)\s*[:\-–]|$)',
            text, re.IGNORECASE | re.DOTALL
        )
        crew_raw = re.sub(r'\s+', ' ', crew_match.group(1)).strip() if crew_match else ""
        crew = fmt(crew_raw, rows_with_data, crew_label)

        # Водій
        driver_match = re.search(r'Водій\s*:\s*(.+)', text)
        driver = f"Водій: {fmt(driver_match.group(1).strip(), rows_with_data, crew_label)}" if driver_match else ""

        # Засіб
        vehicle_match = re.search(r'Засіб\s*:\s*(.+)', text)
        vehicle = vehicle_match.group(1).strip() if vehicle_match else ""

        # Номер вильоту
        SUPERSCRIPT_MAP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
        flight_num_match = re.search(r'Виліт\s*([\d⁰¹²³⁴⁵⁶⁷⁸⁹]+)', text, re.I)
        flight_num = flight_num_match.group(1).translate(SUPERSCRIPT_MAP) if flight_num_match else ""

        # Позиція: "Позиція: НАЗВА" (окремим рядком) і опис у дужках на
        # наступному рядку - об'єднуємо в один вираз "НАЗВА (опис)"; раніше
        # бралась лише дужкова частина, а сама назва позиції губилась.
        position_match = re.search(r'Позиція\s*:\s*(.+)', text, re.I)
        position_name = position_match.group(1).strip() if position_match else ""
        location_match = re.search(r'\(([^)]+)\)', text)
        location_desc = location_match.group(1).strip() if location_match else ""
        if position_name and location_desc:
            location = f"{position_name} ({location_desc})"
        elif position_name:
            location = position_name
        elif location_desc:
            location = f"({location_desc})"
        else:
            location = ""

        # результат
        result_match = re.search(r'Результат[:\-–]\s*(.+)', text)
        result_str = result_match.group(1).strip() if result_match else ""

        # події
        events_match = re.search(r'Події:\s*(.+)', text)
        events = events_match.group(1).strip() if events_match else ""

        crew_and_driver = crew
        if driver:
            crew_and_driver = f"{crew_and_driver}. {driver}" if crew_and_driver else driver

        line = (
            f"{time_str} {date_str} Звіт по роботі екіпажу \"{crew_label}\" {UNIT_BATTALION} {UNIT_BRIGADE}. "
            f"Позивний: {crew_and_driver}. "
            f"Засіб: {vehicle}. "
            f"Виліт {flight_num}. Логістичне забезпечення. "
            f"Позиція: {location}. "
            f"Результат — {result_str}. "
            f"Події: {events}."
        )
        line = re.sub(r"\s+", " ", line).strip()
        line = re.sub(r"\.+", ".", line)
        line = line.replace(" .", ".")

        result.append({"date": f"{time_str} {date_str}", "text": line})

    return result

def get_cost_data(messages, rows_with_data, selected_date):
    messages = get_filtered_body_from_cost_messages(messages)
  

    json_data = [get_parsed(b) for b in messages]

    text =  get_text(json_data, rows_with_data)

    # today_short = selected_date[:6]
    # yesterday_short = f"{str(int(selected_date[:2]) - 1).zfill(2)}{selected_date[2:6]}"
    # text = [t for t in text if not re.search(r'доставлен', t.get("text", ""), re.IGNORECASE) and t.get("text", "").strip() and (yesterday_short in t.get("text", "") or today_short in t.get("text", ""))]
    text = [t for t in text if not re.search(r'доставлен', t.get("text", ""), re.IGNORECASE) and t.get("text", "").strip()]
           
    return text

# прості поля "мітка: значення до кінця рядка" - без пост-обробки, крім .strip().
# \s* перед двокрапкою - деякі екіпажі пишуть "Екіпаж : FPV Дзиґа" (пробіл
# перед двокрапкою); без цього допуску поле не розпізнавалось узагалі, і тип
# БпЛА (FPV) зникав із тексту звіту - через що й синоніми БК не резолвились
# (немає "fpv" в тексті - не проходить перевірка типу БпЛА).
_SIMPLE_FIELDS = [
    ("Екіпаж", r'Екіпаж\s*:\s*(.+)'),
    ("Ціль", r'Ціль\s*:\s*(.+)'),
    ("Результат", r'Результат[:\s]*([^\n]+)'),
    ("Засіб", r'Засіб\s*:\s*(.+)'),
    ("Пілот", r'Пілот[:\s]*([^\n]+)'),
    ("Штурман", r'Штурман[:\s]*([^\n]+)'),
    # Споряджаюч(ий|ій) - деякі екіпажі пишуть з "і" на кінці ("Споряджаючій:")
    # замість "и" ("Споряджаючий:") - без цього допуску поле не розпізнавалось
    # узагалі, і споряджаючий зникав із тексту звіту.
    ("Споряджаючий", r'Споряджаюч(?:ий|ій)[:\s]*([^\n]+)'),
]

# для збереження офіційного написання (з великими літерами не лише на
# першому слові) при форматуванні зведення ВБпАК - див. calculate_uav_data
_MUNITION_DISPLAY_NAMES_BY_LOWER = {v.lower(): v for v in UAV_MUNITION_DISPLAY_NAMES.values()}


def _is_component_item(name):
    """Носій-платформа ("Лупиніс"/"Вирій"), детонатор ("ЕД 8") і плата ініціації
    завжди йдуть в одній "Витрата:" РАЗОМ з основним боєприпасом і списуються як
    ЄДИНИЙ акт - тому окремим рядком у "Витрата:" і зведенні ВБпАК не показуються."""
    normalized = name.lower()
    return any(normalized.startswith(prefix) for prefix in UAV_SPEND_COMPONENT_PREFIXES)


def _normalize_munition_display_name(name):
    """Неофіційне написання відомого виробу з "Витрата:" -> офіційна назва
    (UAV_MUNITION_DISPLAY_NAMES), якщо є однозначний збіг; інакше без змін."""
    return UAV_MUNITION_DISPLAY_NAMES.get(clean_text(name).lower(), name)


def get_parsed(text):
    record = {}

    # \d{1,2} для години допускає однозначний запис без нуля попереду ("26.07.2026
    # 6:20") - раніше з \d{2} така година взагалі не проходила, і повідомлення
    # лишалось без часу (запис показувався лише з датою).
    dt_match = re.search(r'(\d{2}\.\d{2}\.\d{2,4})\s*(\d{1,2}:\d{2})?', text)
    if dt_match:
        record['Дата'] = dt_match.group(1)
        time_part = dt_match.group(2)
        if time_part:
            hh, mm = time_part.split(':')
            time_part = f"{int(hh):02d}:{mm}"
        record['час'] = time_part

    # підрозділ
    unit = re.search(r'Підрозділ\s*:\s*(.+)', text)
    if unit:
        record['Підрозділ'] = clean_text(unit.group(1))

    for field, pattern in _SIMPLE_FIELDS:
        m = re.search(pattern, text)
        if m:
            record[field] = m.group(1).strip()

    # координати - DOTALL + lookahead на наступне поле (а не лише поточний рядок),
    # бо трапляється формат "Координати: околиці нп. Філія\n37U CP 31973 27452"
    # (назва н.п. і сама MGRS-сітка на РІЗНИХ рядках) - без цього MGRS-частина
    # мовчки губилась, лишалась тільки назва населеного пункту.
    target = re.search(
        r'Координати\s*:\s*(.+?)(?=\n\s*(?:Результат|Витрат[аи]|Пілот|Штурман|Споряджаю|Стрім|Стрим|Водій|Події)\s*[:\-–]|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    if target:
        coords_raw = re.sub(r'\s+', ' ', target.group(1)).strip()

        # Регулярний вираз для форматування:
        # (\d{2}[A-Z]) - 2 цифри + літера (Зона: 36T)
        # ([A-Z]{2})   - 2 літери (Квадрат: VS)
        # (\d{5})      - 5 цифр (Схід)
        # (\d{5})      - 5 цифр (Північ)
        # \s* між частинами (а не жорстка відсутність пробілу) - трапляється
        # частково розділений запис "37U CP 3200127426" (зона й квадрат з
        # пробілами, а схід+північ злиплись в одне число без пробілу).
        formatted_coords = re.sub(r'(\d{2}[A-Z])\s*([A-Z]{2})\s*(\d{5})\s*(\d{5})', r'\1 \2 \3 \4', coords_raw)

        # Зберігаємо результат з вашими замінами
        record['Координати'] = formatted_coords

    spend_match = re.search(
        r'Витрат[аи]\s*:?\s*(.+?)(?=\n(?:Пілот|Штурман|Споряджаю|Стрім|Водій|Події)\s*:|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )

    if spend_match:
        spend_text = spend_match.group(1).strip()
        spend_text = re.sub(r'\s+', ' ', spend_text)

        # "розрив" - слово-опис результату детонації поруч із вже врахованою
        # кількістю ("Nшт розрив" або "розрив N шт"), не нова позиція БК.
        # Прибираємо його РАЗОМ із сусідньою кількістю ще до розбору патернів -
        # інакше воно "приклеюється" до сусідньої реальної назви при
        # нежадібному пошуку (напр. "Моа-400 (1 шт) - розрив 1 шт Термобар
        # (1 шт)" без цього кроку недожадібно захоплював би "розрив 1 шт
        # Термобар" одним шматком, ховаючи "Термобар" від патерну 3).
        spend_text = re.sub(r'[-–—]?\s*\bрозрив\b\s*\d*\s*шт?\.?', ' ', spend_text, flags=re.IGNORECASE)
        spend_text = re.sub(r'\d+\s*шт\.?\s*[-–—]?\s*\bрозрив\b', ' ', spend_text, flags=re.IGNORECASE)

        # Патерн 1: "Назва - 1 шт" або "Назва – 1 шт" (з тире). Назва може містити
        # ВНУТРІШНІ дефіси ("МОА-400", "ОГ-Б1", "Стік уф-0225") і кому як десятковий
        # роздільник ("КО 1,3") - тому обидва символи в класі назви, а роздільник
        # значення/кількості вимагає пробіл З ОБОХ БОКІВ (" - " чи " — "): інакше
        # regex вважав би внутрішній дефіс роздільником і обрізав назву (напр.
        # "ОГ-Б1" ставало "Б1", а "МОА-400" не знаходилось узагалі, бо після
        # дефіса-роздільника без пробілу йшло "400", і "шт" далі не було). Тире-
        # роздільник трапляється і як "-", і як "–", і як em-dash "—".
        p_with_dash = re.compile(
            r'([a-zA-Zа-яА-ЯіїєІЇЄ][a-zA-Zа-яА-ЯіїєІЇЄ0-9"\'\.\\\/\s,-]*?)'
            r'\s+[-–—]\s+'
            r'(\d+)\s*шт\.?',
            re.IGNORECASE
        )

        # Патерн 2: "МОА400 4шт" або "Лом 9,5кг 1шт" (без тире). Назва тут теж
        # може мати кілька слів і кому ("Лом 9,5кг") - без пробілу/коми в класі
        # назва обривалась на останньому слові перед кількістю (напр. "Лом 9,5кг
        # 1шт" розпізнавалось як назва "кг", бо "кг" - це те, що прямо стоїть
        # перед "1шт").
        p_without_dash = re.compile(
            r'([a-zA-Zа-яА-ЯіїєІЇЄ][a-zA-Zа-яА-ЯіїєІЇЄ0-9\s,-]*?)'
            r'\s+'
            r'(\d+)\s*шт\.?',
            re.IGNORECASE
        )

        # Патерн 3: "Назва (1 шт)" або "Назва (1шт)" - кількість у дужках, як
        # часто пишуть при мінуванні ("стік уф-0225 (24шт)", "Моа-400 (1 шт)").
        p_parenthesized = re.compile(
            r'([a-zA-Zа-яА-ЯіїєІЇЄ][a-zA-Zа-яА-ЯіїєІЇЄ0-9\s,-]*?)'
            r'\s*\(\s*(\d+)\s*шт\.?\s*\)',
            re.IGNORECASE
        )

        # Патерн 4: "1шт осколок" - кількість ПЕРЕД назвою, злита з "шт" без
        # пробілу (зворотний порядок, трапляється при перерахуванні кількох
        # позицій підряд: "1шт осколок\n1шт фугас"). Без цього патерну
        # p_without_dash хибно хапав "шт" (лишок від "1шт") як початок назви
        # ("шт осколок"), а другу позицію ("фугас") губив повністю - тому цей
        # патерн застосовується РАНІШЕ за p_without_dash.
        p_qty_before_name = re.compile(
            r'(\d+)\s*шт\.?\s+'
            r'([a-zA-Zа-яА-ЯіїєІЇЄ][a-zA-Zа-яА-ЯіїєІЇЄ,\s-]*?)'
            r'(?=\s*\d+\s*шт|\s*$)',
            re.IGNORECASE
        )

        # Патерн 5: "Айкос-40шт" - назва й кількість зліплені через дефіс БЕЗ
        # жодних пробілів (на відміну від патерну 1, де дефіс завжди оточений
        # пробілами). Дефіс тут - явний роздільник, тому (на відміну від
        # патерну 2/3) у класі символів назви його НЕМАЄ - інакше для назв із
        # внутрішнім дефісом ("Стік фо-0225-40шт") було б неоднозначно, який
        # саме дефіс відділяє кількість.
        p_glued_dash = re.compile(
            r'([a-zA-Zа-яА-ЯіїєІЇЄ][a-zA-Zа-яА-ЯіїєІЇЄ0-9"\'\.\\\/\s,]*?)'
            r'-(\d+)\s*шт\.?',
            re.IGNORECASE
        )

        found = {}  # використовуємо dict щоб уникнути дублів

        for n, q in p_with_dash.findall(spend_text):
            n = re.sub(r'\s+', ' ', n).strip()
            if re.search(r'[a-zA-Zа-яА-ЯіїєІЇЄ]', n):
                found[n] = q

        # прибираємо вже розпізнані патерном 1 фрагменти з тексту, перш ніж
        # шукати наступними патернами - інакше їхній ширший клас символів у
        # назві (пробіл, кома, дефіс) повторно "сканує" той самий уривок і
        # хапає хвіст на кшталт "... 1,3 -" як окрему сурогатну позицію
        remainder = p_with_dash.sub(' ', spend_text)

        for q, n in p_qty_before_name.findall(remainder):
            n = re.sub(r'\s+', ' ', n).strip()
            if re.search(r'[a-zA-Zа-яА-ЯіїєІЇЄ]', n) and n not in found:
                found[n] = q

        remainder = p_qty_before_name.sub(' ', remainder)

        for n, q in p_without_dash.findall(remainder):
            n = re.sub(r'\s+', ' ', n).strip()
            if re.search(r'[a-zA-Zа-яА-ЯіїєІЇЄ]', n) and n not in found:
                found[n] = q

        remainder = p_without_dash.sub(' ', remainder)

        for n, q in p_parenthesized.findall(remainder):
            n = re.sub(r'\s+', ' ', n).strip()
            if re.search(r'[a-zA-Zа-яА-ЯіїєІЇЄ]', n) and n not in found:
                found[n] = q

        remainder = p_parenthesized.sub(' ', remainder)

        for n, q in p_glued_dash.findall(remainder):
            n = re.sub(r'\s+', ' ', n).strip()
            if re.search(r'[a-zA-Zа-яА-ЯіїєІЇЄ]', n) and n not in found:
                found[n] = q

        # неофіційні написання відомих виробів (напр. "КО Пузатий змій") ->
        # офіційна назва - ДО побудови spend/Витрата, щоб обидва поля й
        # підсумок ВБпАК показували однакову офіційну назву
        if found:
            renamed_found = {}
            for n, q in found.items():
                display_name = _normalize_munition_display_name(n)
                renamed_found[display_name] = renamed_found.get(display_name, 0) + int(q)
            found = renamed_found

        # 'spend' (для агрегації у зведенні ВБпАК) виключає компоненти-носії
        # (Лупиніс/Вирій/ЕД-8/Плата ініціації) - вони йдуть одним актом з
        # основним боєприпасом. Але 'Витрата' (повний текст ЦЬОГО конкретного
        # повідомлення) має показувати ВСЕ знайдене без фільтрації - так само,
        # як написано в сирому повідомленні; раніше обидва поля будувались із
        # ВЖЕ відфільтрованого списку, і компоненти губились навіть із
        # повного тексту окремого вильоту, хоча мали лишатись там видимими.
        if found:
            spend_for_totals = {n: q for n, q in found.items() if not _is_component_item(clean_text(n))}
            record['spend'] = {clean_text(n): int(q) for n, q in spend_for_totals.items()}
            record['Витрата'] = clean_text(", ".join([f"{n} - {q} шт" for n, q in found.items()]))

    # Стрим - "-" чи повна відсутність значення (порожньо після мітки) -
    # показуємо як "відсутній"; раніше пусте значення (без жодного символу
    # після "Стрім:") взагалі не проходило регулярку (вимагала хоча б 1
    # символ), і поле "Стрім:" в результаті мовчки зникало з тексту повністю.
    # [ \t]* (не \s*) після двокрапки - \s* захопив би і сам перевід рядка,
    # тож при порожньому значенні ("Стрім:  \nПілот: ...") капturing-група
    # почала б зі СЛІДУЮЧОГО рядка замість порожнього рядка.
    result = re.search(r'Стр[іи]м\s*:?[ \t]*([^\n]*)', text, re.IGNORECASE)
    if result:
        val = result.group(1).strip()
        record['Стрим'] = "відсутній" if val == "" else val.replace("-", "Відсутній")

    return record
    
def get_text(json_data, rows_with_data):
    data = []

    for d in json_data:
        text = ""
        text += f"{d.get('час', '')} " if d.get('час') else ""
        text += f"{d.get('Дата', '')} " if d.get('Дата') else ""
        text += f"Підрозділ: {d.get('Підрозділ', '')}. " if d.get('Підрозділ') else ""
        text += f"Засіб: {d.get('Засіб', '')}. " if d.get('Засіб') else ""
        text += f"Екіпаж: {d.get('Екіпаж', '')}. " if d.get('Екіпаж') else ""
        text += f"Ціль: {d.get('Ціль', '')}. " if d.get('Ціль') else ""
        text += f"Координати: ({d.get('Координати', '')}). " if d.get('Координати') else ""
        text += f"Результат: {d.get('Результат', '')}. " if d.get('Результат') else ""
        text += f"Витрата: {d.get('Витрата', '')}. " if d.get('Витрата') else ""
        if d.get('Пілот'): text += f"Пілот: {fmt(d['Пілот'], rows_with_data)}. "
        if d.get('Штурман'): text += f"Штурман: {fmt(d['Штурман'], rows_with_data)}. "
        if d.get('Споряджаючий'): text += f"Споряджаючий: {fmt(d['Споряджаючий'], rows_with_data)}. "
        text += f"Стрим: {d.get('Стрим', '')}. " if d.get('Стрим') else ""

        data.append({"text": clean_text(text), 
            "Дата": d.get('Дата', ''), 
            "час": d.get('час', ''),
            "date": f"{d.get('час', '')} {d.get('Дата', '')}",
            "spend": d.get('spend', {}), 
            "Витрата": d.get('Витрата', '')
        })

    return data

def fmt(val, rows_with_data, crew_label=None):
    # 1. Створюємо чисті словники для пошуку {позивний_lower: [усі ПІБ]} та
    # {позивний_lower: позивний як записаний у СПИСКУ}. Список (а не одне
    # значення) - бо один і той самий позивний у СПИСКУ інколи належить
    # ОДРАЗУ кільком людям (напр. кілька людей з однаковим позивним) -
    # без списку останній рядок мовчки перезаписував би попередні.
    names = {}
    display_names = {}
    for r in rows_with_data:
        p_val = r.get('Позивний')
        pib_val = r.get('П.І.Б.')
        if p_val and str(pib_val) != 'nan':
            clean_key = str(p_val).strip().lower().replace('.', '').replace(',', '')
            names.setdefault(clean_key, []).append(str(pib_val).strip())
            display_names[clean_key] = str(p_val).strip()

    crew_key = (crew_label or "").strip().lower()

    res = []
    # 2. Обробка вхідного тексту
    # Розбиваємо за комами, але також враховуємо можливі крапки як роздільники
    raw_parts = re.split(r'[,|.]', str(val))

    for s in raw_parts:
        # Очищаємо частину від зайвих слів (Пілот/Штурман/тощо), крапок та пробілів
        callsign = s.strip()
        # Видаляємо слова-мітки, якщо вони є в тексті
        callsign = clean_text(re.sub(r'(Пілот:|Штурман:|Споряджаючий:)', '', callsign, flags=re.IGNORECASE))

        if not callsign: continue

        # Ключ для пошуку без крапок у нижньому регістрі
        search_key = callsign.lower().replace('.', '').replace(',', '')
        alias_key = CALLSIGN_ALIASES.get(search_key)

        # AMBIGUOUS_CALLSIGN_OVERRIDES перевіряємо ПЕРШИМ, ще до прямого
        # пошуку - бо один із випадків саме такий: позивний з повідомлення
        # ЗБІГАЄТЬСЯ з чужим позивним у СПИСКУ ("Дед" - реально інша людина,
        # ніж потрібний "Дєд"), тож прямий пошук знайшов би, але не того.
        pib = AMBIGUOUS_CALLSIGN_OVERRIDES.get((crew_key, search_key))

        if pib is None:
            candidates = names.get(search_key)
            if not candidates and alias_key and alias_key in names:
                # знайдено лише через скорочення/типо ("Пух" -> "вінні-пух",
                # "Мпсяня" -> "масяня") - показуємо повний/виправлений позивний
                # зі СПИСКУ, а не буквальне написання з сирого повідомлення
                candidates = names[alias_key]
                callsign = display_names[alias_key]
                search_key = alias_key

            if not candidates:
                pib = None
            elif len(candidates) == 1:
                pib = candidates[0]
            else:
                # позивний належить кільком людям, і override для цього
                # екіпажу не заданий - явно повідомляємо, а не мовчки беремо
                # довільного з кількох (щоб такі випадки не губились)
                print(f"⚠️ УВАГА: Позивний '{callsign}' належить кільком людям у СПИСКУ ({', '.join(candidates)}), і немає AMBIGUOUS_CALLSIGN_OVERRIDES для екіпажу '{crew_label}' - беру першого зі списку!")
                pib = candidates[0]

        if pib:
            res.append(f"{callsign} ({pib})")
        else:
            print(f"⚠️ УВАГА: Для позивного '{callsign}' ПІБ досі не знайдено!")
            res.append(callsign)

    return ", ".join(res)

def get_uav_spend_totals(data_uav):
    """Сирі підсумки витрати БК по назві (лише clean_text, без пониження регістру чи
    злиття схожих написань) - використовується для відображення (calculate_uav_data)."""
    totals = defaultdict(int)
    for item in data_uav:
        for name, qty in item.get("spend", {}).items():
            totals[clean_text(name)] += qty
    return totals

def calculate_uav_data(data_uav) -> str:
    summary = defaultdict(int)

    for name, qty in get_uav_spend_totals(data_uav).items():
        # Очищення для відображення: нижній регістр + злиття відомих варіантів написання
        clean_name = name.lower()

        if "моа композит 61" in clean_name:
            clean_name = "моа композит 61"

        # "ОЗМ-72" - той самий боєприпас, що й просто "ОЗМ" (варіант з номером)
        if re.match(r'^озм(-\d+)?$', clean_name):
            clean_name = "озм"

        if clean_name in ["ед-8", "ед 8", "ед8", "ед-8 ", "ед 8 ", "ед8 "]:
            clean_name = "ед 8"

        summary[clean_name] += qty

    # 3. Сортування: за кількістю (desc), потім за назвою (asc)
    sorted_items = sorted(summary.items(), key=lambda x: (-x[1], x[0]))

    # 4. Форматування для виводу
    formatted_results = []
    for name, qty in sorted_items:
        # Якщо назва (без урахування регістру) - вже офіційна назва з
        # UAV_MUNITION_DISPLAY_NAMES, лишаємо ЇЇ написання як є: generic
        # .capitalize() нижче ламає багатослівні офіційні назви з великою
        # літерою НЕ на першому слові (напр. 'ко 1,3 "пузатий змій"' ставало
        # б 'Ко 1,3 "пузатий змій"' замість 'КО 1,3 "Пузатий змій"').
        # Робимо назви охайними (перша літера велика)
        # Абревіатури (ЕД, СВП, HFB) краще виводити капсом
        if name in _MUNITION_DISPLAY_NAMES_BY_LOWER:
            display_name = _MUNITION_DISPLAY_NAMES_BY_LOWER[name]
        elif name.startswith(("ед", "свп", "ярм", "hfb", "hgo", "зб")):
            display_name = name.upper()
        else:
            display_name = name.capitalize()

        formatted_results.append(f"{display_name} - {qty} шт.")

    return clean_text(", ".join(formatted_results))