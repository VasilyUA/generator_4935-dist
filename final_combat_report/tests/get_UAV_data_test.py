import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from get_normalized_data import get_UAV_data as m
from constants import UNIT_BATTALION, UNIT_BRIGADE, UNIT_BRIGADE_MIXED_CASE


# -------------------------
# get_parsed (cost/spend messages)
# -------------------------
def test_get_parsed_simple_fields():
    text = "Екіпаж: Джоні\nЦіль: укриття\nРезультат: уражено\nЗасіб: Vampire\nПілот: Джоні\nШтурман: Лондон\nСпоряджаючий: Макро"
    record = m.get_parsed(text)
    assert record["Екіпаж"] == "Джоні"
    assert record["Ціль"] == "укриття"
    assert record["Результат"] == "уражено"
    assert record["Засіб"] == "Vampire"
    assert record["Пілот"] == "Джоні"
    assert record["Штурман"] == "Лондон"
    assert record["Споряджаючий"] == "Макро"


def test_get_parsed_extracts_unit():
    record = m.get_parsed(f"Підрозділ: {UNIT_BATTALION} {UNIT_BRIGADE}")
    assert record["Підрозділ"] == f"{UNIT_BATTALION} {UNIT_BRIGADE}"


def test_get_parsed_tolerates_space_before_colon_in_labels():
    # регресія: реальні повідомлення частини екіпажів пишуть "Екіпаж :" (з
    # пробілом перед двокрапкою) - без допуску поле взагалі не розпізнавалось,
    # і тип БпЛА (FPV) зникав із тексту звіту, ламаючи резолв синонімів БК
    text = f"Підрозділ : {UNIT_BATTALION} {UNIT_BRIGADE}\nЕкіпаж : FPV Дзиґа\nЦіль : укриття\nЗасіб : Vampire\nКоординати : 37U CP 30887 27518"
    record = m.get_parsed(text)
    assert record["Підрозділ"] == f"{UNIT_BATTALION} {UNIT_BRIGADE}"
    assert record["Екіпаж"] == "FPV Дзиґа"
    assert record["Ціль"] == "укриття"
    assert record["Засіб"] == "Vampire"
    assert record["Координати"] == "37U CP 30887 27518"


def test_get_parsed_date_and_time():
    record = m.get_parsed("20.07.2026 09:15\nЦіль: щось")
    assert record["Дата"] == "20.07.2026"
    assert record["час"] == "09:15"


def test_get_parsed_coordinates_formatted_when_compact():
    record = m.get_parsed("Координати: 37UCP3088727518")
    assert record["Координати"] == "37U CP 30887 27518"


def test_get_parsed_coordinates_left_as_is_when_already_spaced():
    record = m.get_parsed("Координати: 37U CP 30887 27518")
    assert record["Координати"] == "37U CP 30887 27518"


def test_get_parsed_time_accepts_single_digit_hour_and_zero_pads():
    # регресія: "26.07.2026 6:20" (година без нуля попереду) з \d{2}:\d{2}
    # взагалі не проходило - запис лишався зовсім без часу
    record = m.get_parsed("26.07.2026 6:20\nЦіль: щось")
    assert record["Дата"] == "26.07.2026"
    assert record["час"] == "06:20"


def test_get_parsed_coordinates_multiline_place_name_then_grid():
    # регресія: "Координати: околиці нп. Філія\n37U CP 31973 27452" - назва
    # населеного пункту і сама MGRS-сітка на РІЗНИХ рядках; без DOTALL+lookahead
    # захоплювався лише перший рядок (назва), а сітка координат губилась повністю
    text = "Координати: околиці нп. Філія\n37U CP 31973 27452\nРезультат: обстріляно"
    record = m.get_parsed(text)
    assert record["Координати"] == "околиці нп. Філія 37U CP 31973 27452"


def test_get_parsed_coordinates_partially_glued_zone_spaced_grid_joined():
    # регресія: "37U CP 3200127426" - зона й квадрат уже з пробілами, а схід+
    # північ злиплись в одне 10-значне число без пробілу
    record = m.get_parsed("Координати: 37U CP 3200127426")
    assert record["Координати"] == "37U CP 32001 27426"


def test_get_parsed_spend_quantity_before_name_multiple_items():
    # регресія: "1шт осколок\n1шт фугас" - кількість ПЕРЕД назвою, злита з "шт"
    # без пробілу; без окремого патерну p_without_dash хибно хапав лишок "шт"
    # від "1шт" як частину назви ("шт осколок"), а другу позицію ("фугас") губив
    text = "Витрата : 1шт осколок\n1шт фугас"
    record = m.get_parsed(text)
    assert record["spend"] == {"осколок": 1, "фугас": 1}


def test_get_parsed_spend_with_dash_pattern():
    record = m.get_parsed("Витрата: Айкоси - 40 шт, ГХО-1 - 3 шт")
    assert record["spend"] == {"Айкоси": 40, "ГХО-1": 3}
    assert record["Витрата"] == "Айкоси - 40 шт, ГХО-1 - 3 шт"


def test_get_parsed_spend_name_with_internal_hyphen_not_truncated():
    # регресія: дефіс усередині назви ("МОА-400", "ОГ-Б1") раніше сприймався як
    # роздільник назва/кількість і обрізав назву - тепер роздільник вимагає
    # пробіл з обох боків, а дефіс дозволений у самій назві
    record = m.get_parsed("Витрата: МОА-400 - 2 шт. ОГ-Б1 - 4 шт.")
    assert record["spend"] == {"МОА-400": 2, "ОГ-Б1": 4}


def test_get_parsed_spend_name_with_comma_not_truncated():
    # регресія: кома як десятковий роздільник у назві ("КО 1,3") раніше рвала
    # назву навпіл, лишаючи в spend лише хвіст після останньої коми
    record = m.get_parsed('Витрата: Боєприпас КО 1,3 "Пузатий змій" - 1 шт.')
    assert record["spend"] == {'Боєприпас КО 1,3 "Пузатий змій"': 1}


def test_get_parsed_spend_em_dash_separator_recognized():
    # регресія: em-dash "—" (не лише "-"/"–") теж трапляється як роздільник
    # назва/кількість у реальних повідомленнях
    record = m.get_parsed("Витрата: ОЗМ-72 — 2 шт.")
    assert record["spend"] == {"ОЗМ-72": 2}


def test_get_parsed_spend_parenthesized_quantity():
    # "мінування": кількість часто пишуть у дужках, без тире
    record = m.get_parsed("Витрата: стік уф-0225 (24шт)")
    assert record["spend"] == {"стік уф-0225": 24}


def test_get_parsed_spend_multiword_name_without_dash():
    # регресія: без коми/пробілу в класі назва обривалась на слові прямо
    # перед кількістю ("Лом 9,5кг 1шт" розпізнавалось як просто "кг")
    record = m.get_parsed("Витрата: Лом 9,5кг 1шт")
    assert record["spend"] == {"Лом 9,5кг": 1}


def test_get_parsed_spend_excludes_rozryv_result_word():
    # "розрив" - слово-опис результату детонації, що стоїть прямо перед/після
    # кількості і хибно розпізнавалось як назва виробу
    record = m.get_parsed("Витрата: Моа-400 (1 шт) - розрив 1 шт. Термобар (1 шт) - розрив 1 шт.")
    assert record["spend"] == {"Моа-400": 1, "Термобар": 1}


def test_get_parsed_excludes_component_items_from_spend_but_not_from_display():
    # "Лупиніс"/"Вирій" (носій), "Плата ініціації" й "ЕД 8" (детонатор) завжди
    # йдуть в одній "Витрата:" з основним боєприпасом і списуються як один акт -
    # окремим рядком не показуються в spend (агрегація ВБпАК), АЛЕ мають
    # лишатись видимими в тексті "Витрата:" ЦЬОГО конкретного повідомлення -
    # так само, як написано в сирому повідомленні.
    text = (
        'Витрата: Лупиніс 10" день 2.3/5.8 - 1 шт, '
        'Боєприпас КО 1,3 "Пузатий змій" - 1 шт. '
        'Плата ініціації - 1 шт, ЕД 8 - 1 шт.'
    )
    record = m.get_parsed(text)
    assert record["spend"] == {'Боєприпас КО 1,3 "Пузатий змій"': 1}
    assert "Лупиніс" in record["Витрата"]
    assert "Плата ініціації" in record["Витрата"]
    assert "ЕД 8" in record["Витрата"]
    assert 'Боєприпас КО 1,3 "Пузатий змій"' in record["Витрата"]


def test_get_parsed_spend_without_dash_pattern():
    record = m.get_parsed("Витрата: МОА400 4шт")
    assert record["spend"] == {"МОА400": 4}


def test_get_parsed_spend_glued_dash_pattern():
    # реальний формат мінування "Бомбери Посіпаки": назва й кількість
    # зліплені через дефіс БЕЗ пробілів ("Айкос-40шт") - раніше не збігалось
    # із жодним із 4 попередніх патернів (усі вимагають пробіл або дужки між
    # назвою й кількістю), і "Витрата:" зникала з тексту повністю.
    record = m.get_parsed("Витрата: Айкос-40шт")
    assert record["spend"] == {"Айкоси": 40}
    assert record["Витрата"] == "Айкоси - 40 шт"


def test_get_parsed_spend_glued_dash_pattern_does_not_break_internal_hyphen_names():
    # дефіс у класі символів назви НЕМАЄ (на відміну від патерну 1) - інакше
    # для назви з внутрішнім дефісом було б неоднозначно, який дефіс власне
    # відділяє кількість. "Стік уф-0225" уже розпізнається патерном 3
    # (дужки), тож glued-патерн тут узагалі не мав би спрацювати.
    record = m.get_parsed("Витрата: Стік уф-0225 (24шт)")
    assert record["spend"] == {"Стік уф-0225": 24}


def test_get_parsed_normalizes_ko_puzatyy_zmiy_informal_name_to_official():
    # неофіційне написання з сирих повідомлень FPV Дзиґи ("КО Пузатий змій",
    # без "1,3" і лапок) -> офіційна назва з таблиці БК ВБАК, і в spend, і у
    # відображуваній "Витрата:" - раніше лишалось неофіційним написанням.
    record = m.get_parsed("Витрата: Пегас 10\" 2.3/5.8 - 1 шт, КО Пузатий змій - 1 шт")
    assert record["spend"] == {'Пегас 10" 2.3/5.8': 1, 'КО 1,3 "Пузатий змій"': 1}
    assert 'КО 1,3 "Пузатий змій" - 1 шт' in record["Витрата"]


def test_get_parsed_spend_stops_before_next_labeled_field():
    text = "Витрата: Айкоси - 40 шт\nПілот: Джоні"
    record = m.get_parsed(text)
    assert record["spend"] == {"Айкоси": 40}
    assert record["Пілот"] == "Джоні"


def test_get_parsed_stream_dash_becomes_vidsutniy():
    record = m.get_parsed("Стрім: -")
    assert record["Стрим"] == "Відсутній"


def test_get_parsed_stream_empty_becomes_lowercase_vidsutniy():
    # регресія: "Стрім:" без жодного символу після (лише пробіли до кінця
    # рядка) взагалі не проходило регулярку (вимагала хоча б 1 символ), і
    # поле "Стрім:" мовчки зникало з тексту повністю
    record = m.get_parsed("Стрім:  \nПілот: Джоні")
    assert record["Стрим"] == "відсутній"


def test_get_parsed_soryadzhauchyy_alternate_ending():
    # деякі екіпажі пишуть "Споряджаючій:" (з "і" на кінці) замість
    # "Споряджаючий:" - обидва варіанти мають розпізнаватись
    record = m.get_parsed("Споряджаючій:Пупупу")
    assert record["Споряджаючий"] == "Пупупу"


def test_get_parsed_no_matches_returns_empty_record():
    assert m.get_parsed("випадковий текст без міток") == {}


def test_get_parsed_fixes_aikos_spelling_to_aikosy():
    # граматична правка "Айкос" (без закінчення) -> "Айкоси" (REPLACEMENTS у
    # constants.py) - не має ламати вже правильне написання (перевіряється
    # межею слова: після "с" у "Айкоси" йде "и", межі слова там немає)
    record = m.get_parsed("Витрата: Айкос - 40 шт.")
    assert record["spend"] == {"Айкоси": 40}

    record_already_correct = m.get_parsed("Витрата: Айкоси - 40 шт.")
    assert record_already_correct["spend"] == {"Айкоси": 40}


# -------------------------
# get_text / fmt
# -------------------------
def test_fmt_resolves_callsign_to_full_name():
    rows = [{"Позивний": "Джоні", "П.І.Б.": "Давидов Євген Валерійович"}]
    assert m.fmt("Джоні", rows) == "Джоні (Давидов Євген Валерійович)"


def test_fmt_unknown_callsign_falls_back_to_bare_callsign(capsys):
    result = m.fmt("Невідомий", [])
    assert result == "Невідомий"
    assert "УВАГА" in capsys.readouterr().out


def test_fmt_splits_multiple_callsigns_by_comma():
    rows = [{"Позивний": "Джоні", "П.І.Б.": "Давидов Євген"}, {"Позивний": "Лондон", "П.І.Б.": "Андріянов Євген"}]
    result = m.fmt("Джоні, Лондон", rows)
    assert result == "Джоні (Давидов Євген), Лондон (Андріянов Євген)"


def test_get_text_combines_available_fields():
    result = m.get_text([{"Дата": "24.07.2026", "час": "09:15", "Підрозділ": UNIT_BATTALION, "Результат": "уражено"}], [])
    text = result[0]["text"]
    assert "09:15" in text
    assert "24.07.2026" in text
    assert f"Підрозділ: {UNIT_BATTALION}." in text
    assert "Результат: уражено." in text
    assert result[0]["date"] == "09:15 24.07.2026"


def test_get_text_resolves_pilot_through_fmt():
    rows = [{"Позивний": "Джоні", "П.І.Б.": "Давидов Євген"}]
    result = m.get_text([{"Пілот": "Джоні"}], rows)
    assert "Пілот: Джоні (Давидов Євген)." in result[0]["text"]


# -------------------------
# get_uav_spend_totals / calculate_uav_data
# -------------------------
def test_get_uav_spend_totals_aggregates_across_items():
    data = [{"spend": {"Айкоси": 40}}, {"spend": {"Айкоси": 5, "ЕД 8": 3}}]
    totals = m.get_uav_spend_totals(data)
    assert totals == {"Айкоси": 45, "ЕД 8": 3}


def _spend_items(*names_and_qty):
    return [{"spend": {name: qty}} for name, qty in names_and_qty]


def test_calculate_uav_data_merges_moa_composite_variants():
    # злиття по substring-перевірці "моа композит 61" - працює лише для написань
    # без лапок навколо "композит" (лапки розривають підрядок і зливання не станеться)
    data = _spend_items(("МОА композит 61 камікадзе", 40), ("моа композит 61 щось інше", 5))
    result = m.calculate_uav_data(data)
    assert "Моа композит 61 - 45 шт." in result


def test_calculate_uav_data_merges_ozm_variants():
    # "ОЗМ-72" - той самий боєприпас, що й просто "ОЗМ" (варіант з номером)
    data = _spend_items(("ОЗМ", 2), ("ОЗМ-72", 2))
    result = m.calculate_uav_data(data)
    assert "Озм - 4 шт." in result


def test_calculate_uav_data_merges_ed8_variants():
    data = _spend_items(("ЕД-8", 2), ("ед 8", 1), ("ЕД8", 1))
    result = m.calculate_uav_data(data)
    assert result == "ЕД 8 - 4 шт."


def test_calculate_uav_data_uppercases_known_abbreviation_prefixes():
    data = _spend_items(("свп уламок", 1))
    result = m.calculate_uav_data(data)
    assert result == "СВП УЛАМОК - 1 шт."


def test_calculate_uav_data_uppercases_zb_prefix():
    # "ЗБ-2500" (осколок->ЗБ-2500 через BK_SYNONYM_CREW_OVERRIDES) - без цього
    # префікса generic .capitalize() ламав би абревіатуру на "Зб-2500"
    data = _spend_items(("зб-2500", 1))
    result = m.calculate_uav_data(data)
    assert result == "ЗБ-2500 - 1 шт."


def test_calculate_uav_data_sorts_by_qty_desc_then_name_asc():
    data = _spend_items(("кг", 1), ("плата ініціації", 1), ("айкоси", 40))
    result = m.calculate_uav_data(data)
    assert result.split(", ")[0] == "Айкоси - 40 шт."
    assert result.index("Кг") < result.index("Плата")


def test_calculate_uav_data_uses_official_casing_for_known_munition_display_names():
    # 'ко пузатий змій' (з UAV_MUNITION_DISPLAY_NAMES) - generic .capitalize()
    # зламав би регістр багатослівної офіційної назви ('Ко 1,3 "пузатий
    # змій"'); має лишитись офіційне написання 'КО 1,3 "Пузатий змій"'.
    data = _spend_items(('КО 1,3 "Пузатий змій"', 2))
    result = m.calculate_uav_data(data)
    assert result == 'КО 1,3 "Пузатий змій" - 2 шт.'


def test_calculate_uav_data_empty_input():
    assert m.calculate_uav_data([]) == ""


# -------------------------
# get_logistics_data
# -------------------------
def test_get_logistics_data_extracts_key_fields_from_real_zmiy_format():
    # реальний формат звіту ВБНК (екіпаж "Змій", НРК "Рись Про") - раніше
    # шаблон очікував вигадані окремі мітки "Підрозділ:"/"Екіпаж:", яких
    # немає в жодному справжньому повідомленні цього типу
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК:  Рись Про\n'
        'Маршрут: Околиці Веселого - околиці Тарасівки - околиці Новопавлівки.\n'
        'План роботи НРК: 27.07.2026-28.07.2026\n'
        'Задачі: логістика та перегонка борта.\n'
        'Час виїзду: 17:20-01:02\n'
        'Загальна вага: 250 кг.\n'
        'Відстань: 23 км.\n'
        'Логістика успішно.'
    )
    result = m.get_logistics_data([{"body": text}])
    assert len(result) == 1
    line = result[0]["text"]
    assert result[0]["date"] == "17:20-01:02 27.07.2026-28.07.2026"
    assert f'Звіт по роботі екіпажу Змій вбнк {UNIT_BATTALION} {UNIT_BRIGADE}.' in line
    assert "Тип НРК: Рись Про." in line
    assert "Маршрут: Околиці Веселого - околиці Тарасівки - околиці Новопавлівки." in line
    assert "Задачі: логістика та перегонка борта." in line
    assert "Загальна вага: 250 кг." in line
    assert "Відстань: 23 км." in line
    assert "Логістика успішно." in line


def test_get_logistics_data_uses_vaha_label_when_no_zagalna_vaha():
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК: Рись Про\n'
        'Маршрут: точка А - точка Б\n'
        'Логістика виконана.\n'
        'Дата: 5.05.2026\n'
        'Вага: 100 кг\n'
        'Відстань: 10 км\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Вага: 100 кг." in result[0]["text"]
    assert "Загальна вага" not in result[0]["text"]


def test_get_logistics_data_omits_zadachi_when_absent():
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК: Рись Про\n'
        'Маршрут: точка А - точка Б\n'
        'Логістика виконана.\n'
        'Дата: 5.05.2026\n'
        'Загальна вага: 100 кг\n'
        'Відстань: 10 км\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Задачі:" not in result[0]["text"]


def test_get_logistics_data_omits_success_line_when_not_reported():
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК: Рись Про\n'
        'Маршрут: точка А - точка Б\n'
        'Логістика виконана.\n'
        'Дата: 5.05.2026\n'
        'Загальна вага: 100 кг\n'
        'Відстань: 10 км\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Логістика успішно" not in result[0]["text"]


def test_get_logistics_data_falls_back_to_zasib_when_no_type_nrk():
    text = (
        'Звіт по роботі екіпажу "Кобра"\n'
        'Засіб: Vampire\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Тип НРК: Vampire." in result[0]["text"]


def test_get_logistics_data_falls_back_to_single_date_when_no_range():
    text = (
        'Звіт по роботі екіпажу "Кобра"\n'
        'Дата: 5.05.2026\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert result[0]["date"].strip() == "05.05.2026"


def test_get_logistics_data_extracts_departure_time():
    text = (
        'Звіт по роботі екіпажу "Кобра"\n'
        'Час виїзду: 09:00-10:00\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert result[0]["date"].startswith("09:00-10:00")


def test_get_logistics_data_falls_back_to_vylit_when_no_vyizd():
    text = (
        'Звіт по роботі екіпажу "Кобра"\n'
        'Виліт 4\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Виліт 4" in result[0]["text"]


def test_get_logistics_data_ignores_messages_without_report_keywords():
    # get_filtered_body_from_logistics_uav_messages вимагає одразу "звіт" і
    # "логістика" в тексті - без обох слів повідомлення відсіюється раніше,
    # ніж дійде до парсингу полів.
    result = m.get_logistics_data([{"body": "Звіт без потрібного слова"}])
    assert result == []


def test_get_logistics_data_extracts_viyizd_when_present():
    text = (
        'Звіт по роботі екіпажу "Кобра"\n'
        'Виїзд 3 (доставки 2)\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "Виїзд 3 (доставки 2)" in result[0]["text"]


def test_get_logistics_data_falls_back_to_za_single_date_short_year():
    # реальний формат "ЗМІЙ": дата рядком "за 2.08.26" - без 4-значного року
    # і без нуля попереду дня; попередні гілки ("План роботи НРК: X-Y",
    # "Дата: X") жодна з них не збігається, тож дата лишалась порожньою.
    text = (
        'Звіт по роботі екіпажу «ЗМІЙ»,\n'
        'Тип НРК:Рись Про \n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Маршрут: околиці Новопавлівки\n'
        'за 2.08.26\n'
        'Виїзд 1\n'
        'Логістика\n'
        '1. Боєприпаси\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "02.08.2026" in result[0]["text"]


def test_get_logistics_data_falls_back_to_bare_date_line():
    # реальний формат "Звіт по роботі екіпажу Змій": дата - окремий рядок без
    # жодної мітки ("Дата:"/"за"/"План роботи НРК:") узагалі.
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК: Рись ПРО\n'
        'Маршрут: хаб Хабіб - Альбатроси\n'
        '01.08.2026\n'
        'Логістика успішно.\n'
        'Загальна вага: 250 кг.\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert "01.08.2026" in result[0]["text"]


def test_get_logistics_data_departure_time_pads_both_sides():
    # "Час виїзду: 6:20-7:10" (обидві межі без нуля попереду) -> обидві
    # доповнюються до 2 цифр: "06:20-07:10"
    text = (
        'Звіт по роботі екіпажу Змій\n'
        f'ВБНК {UNIT_BATTALION.upper().replace(' ', '')} {UNIT_BRIGADE_MIXED_CASE}\n'
        'Тип НРК: Рись Про\n'
        'Маршрут: точка А - точка Б\n'
        'Час виїзду: 6:20-7:10\n'
        'Логістика виконана.\n'
        'Дата: 5.05.2026\n'
    )
    result = m.get_logistics_data([{"body": text}])
    assert result[0]["date"].startswith("06:20-07:10")


# -------------------------
# get_loss_data
# -------------------------
def test_get_loss_data_builds_expected_line():
    text = (
        'Звіт по роботі екіпажу "Посіпаки"\n'
        f"{UNIT_BRIGADE_MIXED_CASE} {UNIT_BATTALION.upper()}\n"
        "Дата: 24.07.2026 19:45\n"
        "Позивний: Каспер\n"
        "Засіб: Vampire\n"
        "Виліт 5\n"
        "Логістичне забезпечення\n"
        "Позиція: Донька\n"
        "(Околиці с. Філія)\n"
        "Результат: доставлено\n"
        "Події: без подій"
    )
    rows = [{"Позивний": "Каспер", "П.І.Б.": "Іванов Іван Іванович"}]
    result = m.get_loss_data([{"body": text}], rows)
    assert len(result) == 1
    assert result[0]["date"] == "19:45 24.07.2026"
    assert result[0]["text"] == (
        f'19:45 24.07.2026 Звіт по роботі екіпажу "Посіпаки" {UNIT_BATTALION} {UNIT_BRIGADE}. '
        'Позивний: Каспер (Іванов Іван Іванович). '
        'Засіб: Vampire. '
        'Виліт 5. Логістичне забезпечення. '
        'Позиція: Донька (Околиці с. Філія). '
        'Результат — доставлено. '
        'Події: без подій.'
    )


def test_get_loss_data_empty_flight_when_no_vyliet_pattern():
    # фільтр вимагає підрядки "виліт" і "логістичне" в тілі, але БЕЗ цифри
    # одразу після "Виліт" flight_num лишається порожнім
    text = (
        'Звіт по роботі екіпажу "Посіпаки"\n'
        "Позивний: Каспер\nЗасіб: Vampire\nДата: 24.07.2026 19:45\n"
        "Логістичне забезпечення виконано, виліт не зафіксовано\n"
        "(Околиці с. Філія)\nРезультат: доставлено\nПодії: без подій"
    )
    result = m.get_loss_data([{"body": text}], [])
    assert result[0]["text"] == (
        f'19:45 24.07.2026 Звіт по роботі екіпажу "Посіпаки" {UNIT_BATTALION} {UNIT_BRIGADE}. '
        'Позивний: Каспер. '
        'Засіб: Vampire. '
        'Виліт. Логістичне забезпечення. '
        'Позиція: (Околиці с. Філія). '
        'Результат — доставлено. '
        'Події: без подій.'
    )


def test_get_loss_data_position_name_without_parenthetical_description():
    # "Позиція: НАЗВА" без окремого опису в дужках - показуємо саму назву
    text = (
        'Звіт по роботі екіпажу "Посіпаки"\n'
        "Позивний: Каспер\nЗасіб: Vampire\nДата: 24.07.2026 19:45\n"
        "Виліт 5\nЛогістичне забезпечення\nПозиція: Донька\n"
        "Результат: доставлено\nПодії: без подій"
    )
    result = m.get_loss_data([{"body": text}], [])
    assert "Позиція: Донька." in result[0]["text"]


def test_get_loss_data_skips_messages_without_date():
    # проходить фільтр (є "виліт" і "логістичне"), але без "Дата: dd.mm.yyyy hh:mm"
    text = "Виліт 5 логістичне забезпечення, Позивний: Каспер, без дати"
    result = m.get_loss_data([{"body": text}], [])
    assert result == []


def test_get_loss_data_includes_driver_when_present():
    text = (
        "Позивний: Каспер\nВодій: Петров\nДата: 24.07.2026 10:00\n"
        "Виліт 1\nЛогістичне забезпечення\nтест"
    )
    result = m.get_loss_data([{"body": text}], [])
    assert "Водій: Петров" in result[0]["text"]


# -------------------------
# get_cost_data (filters out already-delivered logistics results)
# -------------------------
def test_get_cost_data_excludes_delivered_messages():
    messages = [
        {"body": "24.07.2026\nРезультат: доставлено\nВитрата: Айкоси - 1 шт"},
        {"body": "24.07.2026\nРезультат: уражено\nВитрата: Айкоси - 1 шт"},
    ]
    result = m.get_cost_data(messages, [], "24.07.2026")
    assert len(result) == 1
    assert "уражено" in result[0]["text"]


# -------------------------
# get_UAV_data
# -------------------------
def test_get_UAV_data_reads_filters_and_combines_all_three_sections(monkeypatch):
    raw_messages = [{"body": "24.07.2026\nРезультат: уражено\nВитрата: Айкоси - 1 шт"}]
    monkeypatch.setattr(m, "get_signal_data", lambda path: raw_messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_UAV_data(18, "24.07.2026", [])

    assert set(result.keys()) == {"cost_data", "loss_data", "logistics_data"}
    assert len(result["cost_data"]) == 1
    assert result["loss_data"] == []
    assert result["logistics_data"] == []


def test_get_UAV_data_logistics_data_includes_kasper_flights(monkeypatch):
    # регресія: get_UAV_data мав об'єднувати get_logistics_data з get_kasper_
    # logistics_data - без цього вильоти екіпажу "Каспер" повністю зникали
    # з підрахунку "Всього N вильотів" у 3.3, хоча логістика "Посіпаки" рахувалась
    raw_messages = [{"body": (
        f'Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк)\n'
        'Екіпаж "Каспер" працював на логістику.\n'
        '24.07.26\n'
        'Виліт 7   21:54/22:12   ( Кувалда )\n'
        'Успішно доставлена   -   їжа+вода'
    )}]
    monkeypatch.setattr(m, "get_signal_data", lambda path: raw_messages)
    monkeypatch.setattr(m, "get_filtered_messages_for_date", lambda msgs, hour, date: msgs)

    result = m.get_UAV_data(18, "24.07.2026", [])

    assert len(result["logistics_data"]) == 1
    assert "Виліт 7" in result["logistics_data"][0]["text"]


# -------------------------
# get_loss_data - однозначна година в "Дата:" без нуля попереду
# -------------------------
def test_get_loss_data_accepts_single_digit_hour_and_zero_pads():
    # регресія: реальне повідомлення писало "Дата: 25.07.2026 7:21" (без нуля
    # попереду) - стара регулярка вимагала рівно 2 цифри години і мовчки
    # відкидала все повідомлення (виліт зникав з підрахунку "Всього N вильотів")
    text = (
        "Позивний: Каспер\nЗасіб: Vampire\nДата: 25.07.2026 7:21\n"
        "Виліт 2\nЛогістичне забезпечення\n(Околиці с.Філія)\n"
        "Результат: доставлено\nПодії: без подій"
    )
    result = m.get_loss_data([{"body": text}], [])
    assert len(result) == 1
    assert result[0]["date"] == "07:21 25.07.2026"
    assert result[0]["text"].startswith("07:21 25.07.2026")


# -------------------------
# fmt - скорочення позивних (CALLSIGN_ALIASES)
# -------------------------
def test_fmt_resolves_callsign_via_alias_when_operator_uses_shorthand():
    # регресія: оператори іноді підписують "Позивний: Пух" замість повного
    # "Вінні-Пух", зареєстрованого в СПИСКУ - без аліасу ПІБ не резолвився
    # і в тексті лишався голий скорочений позивний без ПІБ.
    rows = [{"Позивний": "Вінні-Пух", "П.І.Б.": "Колесніков Олександр Олександрович"}]
    result = m.fmt("Пух", rows)
    assert result == "Вінні-Пух (Колесніков Олександр Олександрович)"


def test_fmt_direct_match_takes_priority_over_alias():
    rows = [{"Позивний": "Пух", "П.І.Б.": "Інший Пух"}, {"Позивний": "Вінні-Пух", "П.І.Б.": "Колесніков Олександр"}]
    result = m.fmt("Пух", rows)
    assert result == "Пух (Інший Пух)"


def test_fmt_ambiguous_callsign_without_override_warns_and_takes_first(capsys):
    # позивний "Малий" належить кільком людям у СПИСКУ; для екіпажу, не
    # заданого в AMBIGUOUS_CALLSIGN_OVERRIDES, немає способу обрати
    # однозначно - явно попереджаємо і беремо першого зі списку
    rows = [
        {"Позивний": "Малий", "П.І.Б.": "Перший Перший"},
        {"Позивний": "Малий", "П.І.Б.": "Другий Другий"},
    ]
    result = m.fmt("Малий", rows, crew_label="Невідомий екіпаж")
    assert result == "Малий (Перший Перший)"
    assert "належить кільком людям" in capsys.readouterr().out


def test_fmt_ambiguous_callsign_resolved_via_crew_override(monkeypatch):
    monkeypatch.setattr(m, "AMBIGUOUS_CALLSIGN_OVERRIDES", {("карлсон", "малий"): "Четвертий Четвертий"})
    rows = [
        {"Позивний": "Малий", "П.І.Б.": "Третій Третій"},
        {"Позивний": "Малий", "П.І.Б.": "Четвертий Четвертий"},
    ]
    result = m.fmt("Малий", rows, crew_label="Карлсон")
    assert result == "Малий (Четвертий Четвертий)"


# -------------------------
# get_kasper_logistics_data
# -------------------------
def test_get_kasper_logistics_data_parses_single_flight_block():
    body = (
        f'Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк)\n'
        'Екіпаж "Каспер" працював на логістику.\n'
        '24.07.26\n'
        'Виліт 7   21:54/22:12   ( Кувалда )\n'
        'Успішно доставлена   -   їжа+вода'
    )
    result = m.get_kasper_logistics_data([{"body": body}])
    assert len(result) == 1
    assert result[0]["date"] == "21:54 24.07.2026"
    assert result[0]["text"] == (
        f'21:54-22:12 24.07.2026 Звіт {UNIT_BATTALION} {UNIT_BRIGADE}: бомбер-вампір (Старлінк). '
        'Екіпаж "Каспер" працював на логістику. Виліт 7. Позиція: Кувалда. '
        'Успішно доставлена - їжа+вода'
    )


def test_get_kasper_logistics_data_splits_multiple_blocks_in_one_message():
    # реальний кейс: оператор копіює попередній звіт і дописує новий - в
    # одному Signal-повідомленні опиняється ДВА повних блоки підряд
    body = (
        f'Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк)\n'
        'Екіпаж "Каспер" працював на логістику.\n'
        '25.07.26\n'
        'Виліт 4   12:00/12:22   ( Синок )\n'
        'Успішно доставлена   -   вода\n'
        '\n\n'
        f'Звіт {UNIT_BATTALION} {UNIT_BRIGADE_MIXED_CASE}: бомбер-вампір(Старлінк)\n'
        'Екіпаж "Каспер" працював на логістику.\n'
        '25.07.26\n'
        'Виліт 5   12:27/12:48   ( Донька )\n'
        'Успішно доставлена   -   вода'
    )
    result = m.get_kasper_logistics_data([{"body": body}])
    assert len(result) == 2
    assert "Виліт 4" in result[0]["text"] and "Синок" in result[0]["text"]
    assert "Виліт 5" in result[1]["text"] and "Донька" in result[1]["text"]


def test_get_kasper_logistics_data_ignores_unrelated_messages():
    result = m.get_kasper_logistics_data([{"body": "Звіт по роботі екіпажу \"Посіпаки\" Логістичне забезпечення"}])
    assert result == []
