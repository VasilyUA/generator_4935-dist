from datetime import date as date_row

from constants import SECTION_LISTS, SHORT_UNIT_BATTALION, SHORT_UNIT_BRIGADE, MONEY_REPORT_CATEGORIES
from utils.date_utils import to_date
from content.money_report_helpers import normalize_name, resolve_status_category


def _enemy_territory_categories():
    """Назви категорій "дії на території противника" - такі особи не отримують БР
    взагалі. Вираховується ДИНАМІЧНО з фактичного constants.MONEY_REPORT_CATEGORIES
    (категорії з "generate_br": False, напр. "РТГр"/"70_РТГр"/"170_РТГр" - зараз), а
    не захардкоджений список назв: видалення чи перейменування такої категорії (або
    цілого пункту) автоматично прибирає й цей виняток - людина просто отримає БР як
    усі інші, за замовчуванням ("generate_br" за відсутності ключа - True)."""
    return {
        name
        for point_categories in MONEY_REPORT_CATEGORIES.values()
        for name, config in point_categories.items()
        if name != "general" and isinstance(config, dict) and not config.get("generate_br", True)
    }


def generate_content_br_general(rows_with_data, rows_with_dowries_data, col_name, higher_commander_data, status_lookup=None):
    for key in SECTION_LISTS:
        SECTION_LISTS[key].clear()

    # Значення комірки ОБЛІК.xlsx, що рахуються як "бойовий" день для щоденного БР -
    # усі точки MONEY_REPORT_CATEGORIES, крім 30 і 10 (не бойові дні - БР на них не
    # генерується; 10 - "виконання обов'язків військової служби", теж не бойові дії).
    combat_cell_values = set(MONEY_REPORT_CATEGORIES) - {30, 10}
    enemy_territory_categories = _enemy_territory_categories()
    current_date = to_date(col_name)

    # Відбір/нормалізація не залежать від конкретної секції SECTION_LISTS - рахуються
    # РІВНО ОДИН раз (а не по колу для кожної секції), лише зіставлення s з k
    # відбувається нижче, окремо для кожної секції.
    eligible_rows = []
    for row in rows_with_data:
        if higher_commander_data['ПОСАДА'] == row['ПОСАДА']:
            continue
        value = row.get(col_name, None)
        current = value.upper() if isinstance(value, str) else value
        if current not in combat_cell_values:
            continue

        # Ті, хто виконує дії на території противника (категорії з
        # "generate_br": False в MONEY_REPORT_CATEGORIES), не отримують
        # БР - для них БР взагалі не генерується.
        if status_lookup is not None:
            status_raw = status_lookup.get(normalize_name(row.get('ПІБ', '')))
            if resolve_status_category(status_raw, row.get('ПІДРОЗДІЛ')) in enemy_territory_categories:
                continue

        unit = row.get('ПІДРОЗДІЛ', '')
        s = unit.replace(' ', '').lower() if isinstance(unit, str) else unit
        eligible_rows.append((s, row))

    dowry_entries = []
    for row in rows_with_dowries_data:
        unit = row.get('ПРИДАНИЙ ДО', '')
        s = unit.replace(' ', '').lower() if isinstance(unit, str) else unit
        dowry_info = {
            'ЗВАННЯ': row.get('військове звання'),
            'ПІБ': row.get('Прізвище та ініціали'),
            'З': row.get('ДАТА ПРИБУВ (З)'),
            'ПО': row.get('ДАТА ВІДБУТТЯ (ПО)')
        }
        dowry_entries.append((s, dowry_info))

    for k, val in SECTION_LISTS.items():
        for s, row in eligible_rows:
            if (s in k if isinstance(k, (tuple, list)) else s == k):
                val.append(row)

        for s, dowry_info in dowry_entries:
            if s in k or s == k:
                row_from = to_date(dowry_info['З'])
                row_to = to_date(dowry_info['ПО']) if dowry_info['ПО'] is not None else date_row.max

                if row_from and row_from < current_date < row_to:
                    val.append(dowry_info)

    return get_unique_sections_list(SECTION_LISTS)


def get_unique_sections_list(section_list):
    unique_section_list = {}
    for section, rows in section_list.items():
        unique_rows = []
        seen_records = set()

        for row in rows:
            record_id = (row.get('ПІБ'), row.get('ЗВАННЯ'))

            if record_id not in seen_records:
                seen_records.add(record_id)
                unique_rows.append(row)

        unique_section_list[section] = unique_rows
    return unique_section_list


def get_content_log_general_extract_log_war(section_lists, paragraph_map, city, coordinates):
    catalog = []

    for section, data in section_lists.items():
        if not data:
            continue

        section_key = ", ".join(section) if isinstance(section, (list, tuple)) else section
        template = paragraph_map.get(section_key, "")

        if not template:
            continue

        params = {
            "city": city,
            "coordinates": coordinates,
            "battalion": SHORT_UNIT_BATTALION,
            "unit": f"{SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}"
        }
        text = template.format(**params)

        personnel = ", ".join([f"{row.get('ЗВАННЯ', '')} {row.get('ПІБ', '')}".strip() for row in data])

        catalog.append(f"{text} {personnel};")

    return "\n".join(catalog)


def get_catalog_log_general_extract_log_war(section_lists, store, city, coordinates):
    for section, data in section_lists.items():
        key = ", ".join(section) if isinstance(section, (list, tuple)) else section
        parts = [f"{row.get('ЗВАННЯ', '')} {row.get('ПІБ', '')}".strip() for row in data if row.get('ПІБ')]
        if not parts: store.pop(key, None); continue
        content = "; ".join(parts) + ";"
        params = {
            "city": city,
            "coordinates": coordinates,
            "battalion": SHORT_UNIT_BATTALION,
            "unit": f'{SHORT_UNIT_BATTALION} {SHORT_UNIT_BRIGADE}'
        }

        header = store[key].format(**params)
        store[key] = f"{header} {content}"

    return "\n".join(map(str, store.values())).replace("  ", " ").strip()
