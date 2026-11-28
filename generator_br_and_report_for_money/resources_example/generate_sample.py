"""Генерує МІНІМАЛЬНИЙ приклад resources/ОБЛІК.xlsx (кілька вигаданих рядків,
щоб було видно точну структуру) і resources/НАЛАШТУВАННЯ.xlsx (вигадані
назви, номери документів, підстави й тексти - з тією самою структурою
пунктів і категорій рапорту, з якою працює програма). Наявні файли не чіпає.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

З обома файлами програма відпрацьовує повністю (БР за перші три дні місяця
й рапорти) - так видно, що куди потрапляє. Далі замініть вигадане своїм."""
import os
import sys
from datetime import datetime, timedelta

from openpyxl import Workbook

# Скрипт запускається напряму (python resources_example/generate_sample.py),
# тож тека проєкту - не в sys.path сама собою, а книга налаштувань пишеться
# тим самим кодом, що її й читає (utils/project_settings.py).
_PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

from utils import project_settings as ps  # noqa: E402
from utils.settings_book import PREVIOUS_MONTH_DAY, SETTINGS_FILE_NAME  # noqa: E402

_MONTH_NAMES = ["СІЧЕНЬ", "ЛЮТИЙ", "БЕРЕЗЕНЬ", "КВІТЕНЬ", "ТРАВЕНЬ", "ЧЕРВЕНЬ",
                "ЛИПЕНЬ", "СЕРПЕНЬ", "ВЕРЕСЕНЬ", "ЖОВТЕНЬ", "ЛИСТОПАД", "ГРУДЕНЬ"]

RESOURCES_DIR = os.path.join(_PROJECT_DIR, "resources")

HIGHER_COMMANDER_TITLE = "Командир батальйону"
COMMANDER_TITLE = "Начальник штабу-заступник командира батальйону"

# Вигадані люди для прикладу - НІКОЛИ не вставляйте сюди реальні ПІБ.
# (підрозділ, посада, звання, ПІБ, значення трьох перших днів місяця)
PEOPLE = [
    ("упр", HIGHER_COMMANDER_TITLE, "підполковник", "ПЕРШИЙ Перший Першович", (100, 100, 100)),
    ("штаб", COMMANDER_TITLE, "майор", "ДРУГИЙ Другий Другович", (100, 100, 100)),
    ("1 рмп", "Командир взводу", "лейтенант", "ТРЕТІЙ Третій Третьович", (100, 100, "ВП")),
    ("2 рмп", "Стрілець", "солдат", "ЧЕТВЕРТИЙ Четвертий Четвертович", (30, 30, 30)),
    ("мп", "Санітар", "старший солдат", "П'ЯТИЙ П'ятий П'ятович", (10, 10, 10)),
]


def _day(date):
    return date.strftime("%d.%m.%Y")


def settings_example(today=None):
    """Вигадані налаштування - словник тієї самої форми, що й
    project_settings.read_settings. Дати відносно today (поточного місяця),
    щоб підстави діяли в місяці, обраному при першому запуску."""
    today = today or datetime.now()
    first = today.replace(day=1)
    previous_first = (first - timedelta(days=1)).replace(day=1)
    previous_last = first - timedelta(days=1)
    mid = first.replace(day=15)
    unit, brigade = "0 бат", "00 обрбр"

    def reference(group, start, lines, end=None, higher_hq=False, ref_id=""):
        return {"group": group, "start": _day(start), "end": _day(end) if end else None, "lines": lines,
                "higher_hq": higher_hq, "id": ref_id}

    def category(point, name, grounds=(), exclude_general="", use_brs=False, default=False, generate_br=False,
                 include_to_report=True, required_basis=True, column=()):
        return {"point": str(point), "category": name, "grounds": list(grounds), "exclude_general": exclude_general,
                "use_brs": use_brs, "use_brs_from_selected_folder": False, "default": default,
                "generate_br": generate_br, "include_to_report": include_to_report,
                "required_basis": required_basis, "column": list(column)}

    def point(name, general=(), column=()):
        return {"point": str(name), "general": list(general), "column": list(column)}

    period = "за період з {start_date_day} {manth} по {end_date_day} {manth} {year} року"
    return {
        "general": {
            "SHORT_UNIT_BATTALION": unit,
            "SHORT_UNIT_BRIGADE": brigade,
            "FULL_UNIT_BUT": "0 батальйону морської піхоти",
            "FULL_MILITARY_UNIT": "військової частини А0000",
            "HIGHER_COMMANDER_TITLE": HIGHER_COMMANDER_TITLE,
            "COMMANDER_TITLE": COMMANDER_TITLE,
            "SEQUENCE_HIGHER_COMMANDER": ["ЗКБ", "НШ-ЗКБ", "ГС"],
        },
        "numbers": {
            "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY": [
                {"day": PREVIOUS_MONTH_DAY, "бат": "100", "брг": "1000", "general_br_bat": f"90 від {_day(previous_first)}"},
                {"day": 1, "бат": "101", "брг": "1001", "general_br_bat": f"100 від {_day(first)}"},
                {"day": 2, "бат": "102", "брг": "1002", "general_br_bat": f"100 від {_day(first)}"},
                {"day": 3, "бат": "103", "брг": "1003", "general_br_bat": f"100 від {_day(first)}"},
                *({"day": day, "бат": "", "брг": "", "general_br_bat": ""} for day in range(4, 32)),
            ],
            "NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK": [
                {"day": 1, "бат": "100", "посилання_брг": f"900 від {_day(previous_first)}", "посилання_бат": f"90 від {_day(previous_first)}"},
                *({"day": day, "бат": "", "посилання_брг": "", "посилання_бат": ""} for day in range(2, 32)),
            ],
            "NUMBER_OF_DOCUMENTS_BRS_SAVE": [
                {"day": 1, "бз_бат": "104", "бз_брг": "2001"},
                {"day": 2, "бз_бат": "105", "бз_брг": "2001"},
                {"day": 3, "бз_бат": "106", "бз_брг": "2002"},
                *({"day": day, "бз_бат": "", "бз_брг": ""} for day in range(4, 32)),
            ],
        },
        "references": [
            reference("ЖБД", previous_first, [f"ЖБД {unit} {brigade} №000дск/1 від {_day(previous_first)}"], end=mid - timedelta(days=1), ref_id="JBD_1"),
            reference("ЖБД", mid, [f"ЖБД {unit} {brigade} №000дск/2 від {_day(mid)}"], ref_id="JBD_2"),
            reference("БН", previous_first, [f"БН {unit} {brigade} №1 від {_day(previous_first)}"], ref_id="BN_1"),
            reference("Бригада загальні", previous_first, ["ЖБД (Справа №0т)"], higher_hq=True),
            reference("ЖИТТЄДІЯЛЬНІСТЬ", previous_first, [f"БР 0 АК від {_day(previous_first)} №000/0/0000т/0"], higher_hq=True),
            reference("ЖИТТЄДІЯЛЬНІСТЬ", previous_first, [f"БР {brigade} № 0000 від {_day(previous_first)}"], higher_hq=True),
            reference("РТГр", previous_first, [f"БН ТГр «Назва» № 00/ВРСО/00т від {_day(previous_first)}"], higher_hq=True),
            reference("РТГр_100", previous_first, [f"БР ком РТГр {brigade} №1 від {_day(previous_first)}; ЖБД РТГр {brigade} №000дск від {_day(previous_first)};"]),
            reference("ЗРАДН", previous_first, [f"№000дск, том 0 від {_day(previous_first)}", f"БР {brigade} з ППО №0000 від {_day(previous_first)} року"]),
            reference("РТГр_170", previous_first, [f"БР ком РТГр №1 від {_day(previous_first)}; ЖБД РТГр {brigade} №000дск від {_day(previous_first)};"]),
            reference("70_РТГр", previous_first, [f"БР {brigade} №0000 від {_day(previous_first)}", "ЖБД (Справа №0т)"]),
        ],
        "changes_orders": [
            {"start": _day(previous_first), "end": _day(previous_last), "lines": [f"№0000 від {_day(first)}"], "id": "C_1"},
        ],
        "points": [
            point(30, ["Бригада загальні"]), point(50), point(100, ["ЖБД", "БН"]), point("РТГр_100"), point("100_ЗРАДН"),
            point(170, ["ЖБД", "БН"]), point(70, ["ЖБД", "БН"]), point(10), point("10_МЕД"),
            point("NOT_PAID", column=["ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"]), point("100_ВПБП"), point("100_БПШП"), point("100_ДТЛ"),
            point("100_СПЕЦКОНТИНГЕНТ", column=["ПІДСТАВИ СПЕЦКОНТИНГЕНТУ"]),
        ],
        "categories": [
            category(30, "ЖИТТЄДІЯЛЬНІСТЬ", ["ЖИТТЄДІЯЛЬНІСТЬ"], default=True),
            category(30, "РТГр_30", ["РТГр"]),
            category(50, "ПУ", include_to_report=False, required_basis=False),
            category(100, "БД(СЗ)", use_brs=True, default=True, generate_br=True),
            category(100, "МЕДИК", use_brs=True, generate_br=True),
            category("РТГр_100", "РТГр", ["РТГр_100"], exclude_general="всі", default=True),
            category("100_ЗРАДН", "ЗРАДН", ["ЗРАДН"], default=True),
            category(170, "РТГр_170", ["РТГр_170"], exclude_general="всі"),
            category(170, "БД(СЗ)", use_brs=True, default=True, generate_br=True),
            category(70, "70_РТГр", ["70_РТГр"], exclude_general="всі"),
            category(70, "БД(СЗ)", use_brs=True, default=True, generate_br=True),
            category(10, "ЖИТТЄДІЯЛЬНІСТЬ", default=True, required_basis=False, column=["ПІДСТАВИ 10К"]),
            category("10_МЕД", "МЕД", default=True, column=["ПІДСТАВИ 10_МЕД"]),
            category("NOT_PAID", "СЗЧ"),
            category("NOT_PAID", "Задув.|задув|задут"),
            category("NOT_PAID", "Арешт|арешт|арештований"),
            category("100_ВПБП", "ВПБП", default=True, column=["ПІДСТАВИ ВПБП"]),
            category("100_БПШП", "БПШП|100_БПШП|100_ШПБП", default=True, column=["ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП"]),
            category("100_ДТЛ", "ДТЛ", default=True, column=["ПІДСТАВИ ДТЛ"]),
            category("100_СПЕЦКОНТИНГЕНТ", "полон"),
            category("100_СПЕЦКОНТИНГЕНТ", "безвісті|безв|зниклі|безв.знк.|Безвісти|БЗ|100_БЗ"),
            category("100_СПЕЦКОНТИНГЕНТ", "інд"),
        ],
        "commander_points": [
            point(10), point(30), point(50), point(100, ["Бригада загальні"]), point(170), point(70), point("100_БПШП"),
            point("NOT_PAID", column=["ПІДСТАВИ ДЛЯ НЕ ВИПЛАТИ"]),
        ],
        "commander_categories": [
            category(10, "ЖИТТЄДІЯЛЬНІСТЬ", default=True, generate_br=True, required_basis=False, column=["ПІДСТАВИ 10К"]),
            category(30, "ЖИТТЄДІЯЛЬНІСТЬ", default=True, generate_br=True),
            category(50, "ПУ", default=True, include_to_report=False, required_basis=False),
            category(100, "ОБОРОНА", ["ЖИТТЄДІЯЛЬНІСТЬ"], use_brs=True, default=True, generate_br=True),
            category(170, "ОБОРОНА", default=True, generate_br=True),
            category(70, "ОБОРОНА", default=True, generate_br=True),
            category("100_БПШП", "ЛІКУВАННЯ_ПО_ПОРАНЕННЮ", default=True, column=["ПІДСТАВИ ШПБП", "ПІДСТАВИ БПШП"]),
            category("NOT_PAID", "СЗЧ"),
            category("NOT_PAID", "Задув.|задув|задут"),
        ],
        "section_texts": [
            {"point": "10", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. за час виконання обов'язків військової служби {period}:",
             "commander": f"{{amount}} гривень – за час виконання обов'язків військової служби {period}, а саме:"},
            {"point": "30", "report": f"Виплатити додаткову винагороду у розмірі {{amount}} грн. 00 коп. військовослужбовцям, які виконують бойові (спеціальні) завдання, {period}:",
             "commander": f"{{amount}} гривень – за участь у бойових (спеціальних) завданнях {period}, а саме:"},
            {"point": "50", "report": f"Виплатити додаткову винагороду у розмірі {{amount}} грн. 00 коп. військовослужбовцям на пунктах управління {period}:",
             "commander": f"{{amount}} гривень – за участь у виконанні завдань на пунктах управління {period}, а саме:"},
            {"point": "70", "report": f"Виплатити одноразову винагороду у розмірі {{amount}} грн. 00 коп. {period}:",
             "commander": f"{{amount}} гривень – одноразова винагорода {period}, а саме:"},
            {"point": "100", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям, які беруть безпосередню участь у бойових діях, {period}:",
             "commander": f"{{amount}} гривень – за безпосередню участь у бойових діях {period}, а саме:"},
            {"point": "170", "report": f"Виплатити одноразову винагороду у розмірі {{amount}} грн. 00 коп. {period}:",
             "commander": f"{{amount}} гривень – одноразова винагорода {period}, а саме:"},
            {"point": "РТГр_100", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям у складі РТГр {period}:", "commander": ""},
            {"point": "100_ЗРАДН", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям, залученим до протиповітряної оборони, {period}:", "commander": ""},
            {"point": "10_МЕД", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям, які перебувають на лікуванні, {period}:", "commander": ""},
            {"point": "NOT_PAID", "report": "Не виплачувати додаткову винагороду нижчепойменованим військовослужбовцям:",
             "commander": "Не виплачувати додаткову винагороду:"},
            {"point": "100_ВПБП", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям, які перебувають у відпустці для лікування після поранення, {period}:", "commander": ""},
            {"point": "100_БПШП", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям, які перебувають на лікуванні після бойового поранення, {period}:",
             "commander": f"{{amount}} гривень – за лікування після бойового поранення {period}, а саме:"},
            {"point": "100_ДТЛ", "report": f"Виплатити додаткову винагороду в розмірі {{amount}} грн. 00 коп. військовослужбовцям на довготривалому лікуванні {period}:", "commander": ""},
            {"point": "100_СПЕЦКОНТИНГЕНТ", "report": f"Виплатити додаткову винагороду в розмірі 100 000 грн. 00 коп. військовослужбовцям, які перебувають у полоні або зникли безвісти, {period}:", "commander": ""},
        ],
        "report_texts": {
            "REPORT_TITLE": "РАПОРТ",
            "HIGH_COMMANDER": "Командиру військової частини А0000",
            "START_SECTION": "Дійсним доповідаю, що особовий склад 0 батальйону морської піхоти військової частини А0000 виконував бойові завдання згідно з бойовими наказами (розпорядженнями).",
            "SECTION_INFORMATION": "Прошу Вашого дозволу на включення до наказу про виплату додаткової винагороди за {month} {year} року нижчепойменованих військовослужбовців та:",
            "COMMANDER_INTRO": "Прошу виплатити мені {rank} {pib}, {posada} додаткову винагороду за {month} {year} року у таких розмірах:",
            "CHANGES_INTRO": "Прошу внести зміни в наказ командира військової частини А0000 (з адміністративно-господарської діяльності) {order_reference}, а саме:",
            "CHANGES_EXCLUDE_LABEL": "Виключити пункти в Додатку {appendix_number}:",
            "CHANGES_ADD_LABEL": "Додаток {appendix_number} доповнити наступними пунктами:",
            "CHANGES_RESTATE_LABEL": "Викласти в новій редакції в Додатку {appendix_number}:",
            "MISSING_COVERAGE_HEADING": "Відсутні підстави (перевірте перед підписанням):",
        },
        "not_paid_exceptions": [{"status": "СЗЧ", "confirmed_by": ["РОЗП"]}],
        "subdivisions": [{"name": "1 рмп"}, {"name": "2 рмп"}, {"name": "мп"}],
        "subdivision_texts": [
            {"name": "1 рмп",
             "br_text": "1 рмп утримувати займані рубежі в смузі відповідальності {battalion}, не допустити просування противника, залучити нижчепойменований особовий склад:",
             "extract_text": "1 рмп утримувала займані рубежі в смузі відповідальності {battalion}, не допускала просування противника, залучала нижчепойменований особовий склад:"},
            {"name": "2 рмп",
             "br_text": "2 рмп забезпечити готовність резервів батальйону в смузі відповідальності {battalion}, залучити нижчепойменований особовий склад:",
             "extract_text": "2 рмп забезпечувала готовність резервів батальйону в смузі відповідальності {battalion}, залучала нижчепойменований особовий склад:"},
            {"name": "мп",
             "br_text": "Медичному пункту забезпечити надання медичної допомоги особовому складу підрозділів, залучити нижчепойменований особовий склад:",
             "extract_text": "Медичний пункт забезпечував надання медичної допомоги особовому складу підрозділів, залучав нижчепойменований особовий склад:"},
            {"name": "упр, штаб",
             "br_text": "Управління силами батальйону здійснювати з КСП {unit} в районі н. п. {city} (координати: {coordinates}), залучити нижчепойменований особовий склад:",
             "extract_text": "Управління силами батальйону здійснювали з КСП {unit} в районі н. п. {city} (координати: {coordinates}), залучали нижчепойменований особовий склад:"},
        ],
        "task_texts": {
            "BR_ENEMY_ASSESSMENT_TEXT": "Противник (вигаданий опис для прикладу) продовжує активні дії в смузі оборони {brigade}.",
            "BR_FIRE_SUPPORT_INTRO_TEMPLATE": "За викликом засобами старшого начальника уражаються цілі відповідно до таблиці вогню артилерії згідно плану.",
            "BR_TASK_SECTION_TEMPLATE": "Відповідно до бойового розпорядження командира {brigade} №{brg_ref} та бойового розпорядження командира {unit} №{bat_ref} {unit} продовжує утримувати район оборони.",
            "BR_SITUATION_UPDATE_TEMPLATE": "Основні зусилля зосередити на напрямку Н.П.1 – Н.П.2.",
            "BR_DEFENSE_TASKS_TEXT": "Продовжити виконувати завдання з оборони позицій, знищення противника вогнем наявних засобів.",
            "BR_DEFENSE_MEASURES_TEXT": "Продовжити утримувати займані рубежі, накопичувати запаси матеріальних засобів.",
        },
        "task_variants": [
            {"variant": 1, "title": "1 рмп (оборона РОП «Приклад-1»)", "text": "Утримувати РОП «Приклад-1», не допустити прориву противника."},
            {"variant": 1, "title": "2 рмп (у складі резерву)", "text": "Перебувати в готовності до посилення оборони на напрямку Н.П.1."},
            {"variant": 2, "title": "1 рмп (у складі резерву)", "text": "Перебувати в готовності до посилення оборони на напрямку Н.П.2."},
            {"variant": 2, "title": "2 рмп (оборона РОП «Приклад-2»)", "text": "Утримувати РОП «Приклад-2», не допустити прориву противника."},
        ],
    }


def write_oblik(path, today=None):
    """ОБЛІК.xlsx: аркуш поточного місяця (3 перші дні) + порожній "ТВО"."""
    today = today or datetime.now()
    sheet_name = _MONTH_NAMES[today.month - 1]
    first_of_month = today.replace(day=1)
    dates = [first_of_month + timedelta(days=i) for i in range(3)]  # лише 3 дні для прикладу

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(["ПІДРОЗДІЛ", "ПОСАДА", "ЗВАННЯ", "ПІБ", *dates, "ПІДСТАВИ 10К"])
    for pidrozdil, posada, zvannya, pib, values in PEOPLE:
        ws.append([pidrozdil, posada, zvannya, pib, *values, "наказ №0 від 01.01.2026" if 10 in values else None])
    for col in range(5, 5 + len(dates)):
        ws.cell(row=1, column=col).number_format = "dd.MM.yyyy"  # див. generators/generate_timetable.py

    ws_tvo = wb.create_sheet("ТВО")
    ws_tvo.append(["Підрозділ", "ПОСАДА", "Start", "End", "ЗВАННЯ", "ПІБ", "ТВО"])
    # Порожній (без рядків людей) - так само коректно, як і заповнений.

    wb.save(path)
    return sheet_name


def main(resources_dir=RESOURCES_DIR):
    os.makedirs(resources_dir, exist_ok=True)

    oblik_path = os.path.join(resources_dir, "ОБЛІК.xlsx")
    if os.path.exists(oblik_path):
        print(f"{oblik_path} вже існує - нічого не роблю (заберіть/перейменуйте файл, якщо хочете новий приклад).")
    else:
        sheet_name = write_oblik(oblik_path)
        print(f"Створено {oblik_path} (аркуш '{sheet_name}' + порожній 'ТВО').")
        print("Замініть вигаданих людей на реальних і заповніть усі дні місяця.")

    settings_path = os.path.join(resources_dir, SETTINGS_FILE_NAME)
    if os.path.exists(settings_path):
        print(f"{settings_path} вже існує - не чіпаю.")
    else:
        ps.write_settings(settings_path, settings_example())
        print(f"Створено {settings_path} (вигадані назви, номери й підстави) - впишіть свої дані.")


if __name__ == "__main__":
    main()
