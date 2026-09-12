import difflib
import os
import re
from collections import Counter
from datetime import date, datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from constants import (
    DEFAULT_STATUS,
    LABEL_COLUMN_NAMES,
    PAYMENT_STATUS_CONVERSIONS,
    PAYMENT_STATUS_OVERRIDES,
    PAYMENT_STATUS_PRIORITY,
    PIB_COLUMN_NAME,
    STATUS_COLORS,
    STATUS_CONFLICT_OVERRIDES,
)


def _fill_for_status(status):
    """PatternFill для STATUS_COLORS[status], чи None, якщо для цього статусу
    кольору не задано (лишає клітинку без заливки - не помилка, див.
    constants.STATUS_COLORS)."""
    color = STATUS_COLORS.get(status)
    return PatternFill(start_color=color, end_color=color, fill_type="solid") if color else None


def normalize_name(name):
    """Нормалізує ПІБ для зіставлення рапорту з ОБЛІК.xlsx (пробіли, апостроф,
    регістр) - та сама логіка, що й у generator_br_and_report_for_money
    (content/money_report_helpers.py, checker_accounting/checker.py), своя
    самодостатня копія - той самий підхід, що й checker.py: ці два проєкти не
    імпортують одне одного."""
    if not isinstance(name, str):
        return ""
    cleaned = re.sub(r"(В|в)'[\s]+([а-яА-ЯІіЇїЄєҐґ])", r"\1'\2", name.strip())
    return " ".join(cleaned.split()).upper()


def _strip_external_links(wb):
    """Реальний випадок: resources/ОБЛІК.xlsx має "осиротіле" зовнішнє
    посилання (Excel External Reference - залишок від колись видаленого
    Data Validation/формули, що вказувала на КОНКРЕТНИЙ, застарілий файл
    resources/information_unit/*.xlsx) - перевірено: жодна формула чи
    правило перевірки даних у самій книзі ним фактично НЕ користується,
    саме лише кешоване посилання лишилось "висіти". openpyxl НЕ вміє
    коректно ПЕРЕЗБЕРЕГТИ таку книгу (при повторному save() зв'язок r:id
    всередині externalLinks/externalLink1.xml лишається "rId1", а
    відповідний запис у externalLink1.xml.rels губиться) - результат Excel
    сприймає як пошкоджений вміст ("Знайдено проблему з вмістом... Відновлені
    записи") при кожному відкритті згенерованого файлу. Оскільки ніщо цим
    посиланням не користується - просто прибираємо його перед збереженням,
    а не намагаємось "полагодити" те, що openpyxl однаково не вміє записати
    коректно."""
    wb._external_links = []


class Timesheet:
    """Обгортка над робочою книгою ОБЛІК.xlsx (аркуш "Табель"): зчитує лише
    структуру заголовка (колонки-дати й "іменовані" колонки - за назвою, а не
    фіксованою позицією), сама книга (wb/ws) лишається відкритою для
    редагування комірок НАПРЯМУ (apply_events) і збереження (save). На
    відміну від checker_accounting/checker.py (будує НОВУ книгу й тому копіює
    стилі вручну), тут результат - ТА САМА книга, тож стилі/ширини колонок/
    об'єднані клітинки зберігаються самі собою."""

    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.wb = load_workbook(file_path)
        self.ws = self.wb[sheet_name]

        header = [self.ws.cell(row=1, column=c).value for c in range(1, self.ws.max_column + 1)]

        self.label_columns = {
            value.strip().upper(): idx
            for idx, value in enumerate(header, start=1)
            if isinstance(value, str) and value.strip().upper() in LABEL_COLUMN_NAMES
        }
        if PIB_COLUMN_NAME not in self.label_columns:
            raise ValueError(f"У файлі {file_path} не знайдено колонку '{PIB_COLUMN_NAME}' у заголовку.")

        self.date_columns = sorted(
            (
                (idx, value.date() if isinstance(value, datetime) else value)
                for idx, value in enumerate(header, start=1)
                if isinstance(value, date)
            ),
            key=lambda pair: pair[1],
        )
        if not self.date_columns:
            raise ValueError(f"У файлі {file_path} не знайдено жодної колонки-дати в заголовку.")

    def person_rows(self):
        """{normalize_name(ПІБ): row_idx} - рядок з порожнім ПІБ пропускається."""
        pib_col = self.label_columns[PIB_COLUMN_NAME]
        rows = {}
        for row_idx in range(2, self.ws.max_row + 1):
            pib_raw = self.ws.cell(row=row_idx, column=pib_col).value
            if pib_raw:
                rows[normalize_name(pib_raw)] = row_idx
        return rows

    def save(self, output_path):
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        _strip_external_links(self.wb)
        try:
            self.wb.save(output_path)
        except PermissionError as e:
            raise PermissionError(
                f"Не вдалось зберегти '{output_path}': файл зараз відкритий в іншій програмі "
                "(напр. Excel). Закрийте його та запустіть генерацію знову."
            ) from e


def export_for_money_project(source_path, target_path, target_sheet_name):
    """За прямою вказівкою користувача - готує ВЖЕ згенерований файл "ОБЛІК
    для виплат" (source_path) для сусіднього проєкту
    generator_br_and_report_for_money (index.py, лише коли прогін ПОВНІСТЮ
    чистий): завантажує книгу ОКРЕМО від source_path (джерело лишається
    незмінним) і перейменовує її єдиний аркуш на target_sheet_name -
    constants.MONEY_PROJECT_SHEET_NAME (назва ПОТОЧНОГО місяця, напр. "СЕРПЕНЬ" -
    обчислюється динамічно, див. constants.py) - точну назву, яку шукає той
    проєкт (інакше він падає на родовий fallback-аркуш і читає геть не ті
    дані). Колонки (ПІДРОЗДІЛ/ПОСАДА/ЗВАННЯ/ПІБ + по колонці на
    кожну дату + ПІДСТАВИ/ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ) вже структурно ті
    самі, що чекає той проєкт - перетворення даних не потрібне."""
    wb = load_workbook(source_path)
    wb.active.title = target_sheet_name
    _strip_external_links(wb)
    os.makedirs(os.path.dirname(target_path) or ".", exist_ok=True)
    try:
        wb.save(target_path)
    except PermissionError as e:
        raise PermissionError(
            f"Не вдалось зберегти '{target_path}': файл зараз відкритий в іншій програмі "
            "(напр. Excel). Закрийте його та спробуйте ще раз."
        ) from e


# Поріг схожості (difflib.SequenceMatcher.ratio) для _typo_hint нижче -
# емпірично перевірено на реальних даних (resources/ОБЛІК.xlsx, ~300+ ПІБ):
# на цьому порозі реальна друкарська помилка (переставлені літери в
# прізвищі) дає ЄДИНИЙ збіг (ratio ~0.96), а вигадане, справді відсутнє ПІБ -
# жодного - досить високий, щоб не підказувати ІНШУ, НЕ ТУ людину.
_TYPO_HINT_SIMILARITY_CUTOFF = 0.8


def _typo_match(timesheet, person_rows, pib_col, normalized):
    """(candidate_normalized, candidate_pib_raw) - якщо СЕРЕД роcтера є
    РІВНО ОДНЕ нормалізоване ПІБ, дуже СХОЖЕ (difflib, поріг
    _TYPO_HINT_SIMILARITY_CUTOFF) на normalized, але НЕ ІДЕНТИЧНЕ (інакше
    normalized уже був би в person_rows) - типова причина: друкарська
    помилка в рапорті (реальний випадок - переставлені літери в прізвищі
    порівняно з роcтеровим написанням). За прямою вказівкою користувача -
    подія ЗАСТОСОВУЄТЬСЯ до цього кандидата автоматично (apply_events
    нижче), а НЕ лишається лише підказкою (той самий принцип "показати
    ЙМОВІРНОГО кандидата", що й _candidate_roster_person у content/
    payment_mismatch_checker.py, лише ширший критерій схожості - тут, на
    відміну від ТОГО випадку, помилка може бути й у самому ПРІЗВИЩІ, не
    лише в по-батькові, тож "перші два слова" тут не спрацювали б, і, НА
    ВІДМІНУ від payment_mismatch_checker.py, результат тут ЗАСТОСОВУЄТЬСЯ,
    а не лише відображається). (None, None), якщо збігів немає чи їх
    кілька (неоднозначно - краще нічого не застосувати, ніж застосувати до
    НЕ ТІЄЇ людини).

    Реальний випадок (баг, підтверджений на реальних даних): ДВІ РІЗНІ
    людини з ТИМ САМИМ ім'ям+по-батькові (поширене поєднання), але РІЗНИМИ
    прізвищами (двоскладові прізвища-антоніми на кшталт "ПЕРШИЙ"/"ДРУГИЙ" -
    див. тест test_apply_events_does_not_auto_apply_when_similar_pib_has_a_different_surname
    у tests/test_oblik_timesheet.py) - ratio ЦІЛОГО ПІБ хибно перевищував
    поріг (спільна велика частина імені+по-батькові "переважує" різницю в
    прізвищі - 0.89, вище порогу), хоча самі ПРІЗВИЩА - геть різні слова
    (ratio ~0.5, а НЕ типова друкарська помилка ~0.9+, підтверджено на
    встановленому реальному прикладі "ЧОТИРНАЦДЯТИЙ"/"ЧОТИРНАДЦЯТИЙ" - 0.92).
    Тож прізвище (перше слово нормалізованого ПІБ) МАЄ ОКРЕМО пройти той
    самий поріг - інакше це НЕ друкарська помилка в ОДНІЄЇ людини, а просто
    ІНША людина з випадково схожим рештою імені."""
    matches = difflib.get_close_matches(normalized, person_rows.keys(), n=2, cutoff=_TYPO_HINT_SIMILARITY_CUTOFF)
    if len(matches) != 1:
        return None, None
    candidate_normalized = matches[0]
    target_surname = normalized.split(" ", 1)[0]
    candidate_surname = candidate_normalized.split(" ", 1)[0]
    surname_ratio = difflib.SequenceMatcher(None, target_surname, candidate_surname).ratio()
    if surname_ratio < _TYPO_HINT_SIMILARITY_CUTOFF:
        return None, None
    candidate_pib_raw = timesheet.ws.cell(row=person_rows[candidate_normalized], column=pib_col).value
    return candidate_normalized, candidate_pib_raw


def apply_events(timesheet, events, report_dates):
    """Заповнює колонки-дати аркуша: для кожної дати, за яку є рапорт (у
    порядку зростання) - застосовує події ЦІЄЇ дати (зміна статусу), а для
    решти людей КОПІЮЄ статус із попередньої колонки-дати ("продовжує
    статус") - те саме правило діє і для дат МІЖ рапортами, за які рапорту
    немає (просто немає подій на ту дату - і всі так само переносять
    попередній статус). За прямою вказівкою користувача - ЯКЩО для дати
    НЕМАЄ події з рапорту, АЛЕ в самому resources/ОБЛІК.xlsx ця колонка ВЖЕ
    має вручну вписаний статус, ВІДМІННИЙ від того, що переноситься з
    попередньої колонки (реальний випадок - "БЗ" вручну вписано на дату
    В МЕЖАХ діапазону наявних рапортів) - цей статус
    так само стає новим "якорем" (як і подія з рапорту) і ПРОДОВЖУЄТЬСЯ у
    ВСІ наступні колонки, аж доки не трапиться подія з рапорту чи ще один
    вручну вписаний статус. Найперша колонка-дата (базовий стан на "вчора"
    першого рапорту) НЕ змінюється - вона лишається початковою точкою
    відліку (лише її ЗАЛИВКА кольором оновлюється відповідно до вже
    наявного там тексту, для єдиної кольорової гами по всьому рядку - САМ
    текст лишається як є). Дати ПІСЛЯ останнього наявного рапорту САМІ ПО
    СОБІ не займаються - жоден рапорт ще не підтвердив для них "без змін" -
    ПОРОЖНІ колонки в цьому діапазоні лишаються порожніми. Так само, як і в
    межах діапазону рапортів, вручну вписаний статус тут теж стає "якорем" і
    ПРОДОВЖУЄТЬСЯ у ВСІ наступні порожні колонки, аж доки не трапиться ще
    один вручну вписаний статус чи не скінчаться колонки табеля.

    Повертає список {"reason", ...} - записи ПРО ЯКІ ПОТРІБНО ЗНАТИ, не
    завжди лише "не вдалось застосувати": людину з events не знайдено в
    аркуші за нормалізованим ПІБ буквально, АЛЕ є РІВНО ОДНЕ дуже схоже
    написання в роcтері (_typo_match вище, типово - друкарська помилка в
    рапорті) - за прямою вказівкою користувача, подія ВСЕ ОДНО
    ЗАСТОСОВУЄТЬСЯ до цього кандидата (як типовий case нижче), а запис у
    unresolved лише ІНФОРМУЄ про це (не потребує дії, для прозорості - щоб
    у файлі "Потребує ручної перевірки" було видно, що сталось і кого це
    стосується); СПРАВДІ не застосовується (і лишається на ручну
    перевірку), лише якщо схожих написань немає взагалі чи їх кілька
    (неоднозначно). effective_date не збігається з жодною колонкою-датою
    аркуша (типова причина - друкарська помилка дати в рапорті), або на
    одну людину й одну дату в рапорті надійшло кілька РІЗНИХ статусів -
    якщо пара статусів є в STATUS_CONFLICT_OVERRIDES, застосовується
    визначений там "переможець" БЕЗ запису на ручну перевірку (це не
    справжня суперечність), інакше застосовується останній і запис іде на
    ручну перевірку. effective_date РАНІШИЙ за найпершу (базову) колонку-дату
    І статус події ЗБІГАЄТЬСЯ з тим, що ВЖЕ стоїть у базовій колонці - реальний
    випадок: подія зі щоденного рапорту вказує дату, що передує ВЖЕ наявному в
    resources/ОБЛІК.xlsx базовому стану (напр. "з 21.07 у відпустці", а базова
    колонка табеля починається щойно з 31.07, і ТАМ ВЖЕ стоїть той самий
    статус - ВПСЗ) - подія лише ПІДТВЕРДЖУЄ те, що вже записано, базова
    колонка НІКОЛИ не змінюється (докстрінг вище), тож мовчки пропускається,
    БЕЗ запису на ручну перевірку (нема нічого, що людина могла б виправити чи
    підтвердити). ЯКЩО Ж статус НЕ збігається з базовим (типова причина -
    друкарська помилка дати в самому рапорті, напр. "06.07" замість "06.08") -
    це, як і раніше, лишається на ручну перевірку: раніша за діапазон дата
    сама по собі ще НЕ доказ, що це не помилка, лише ЗБІГ зі статусом, що вже
    підтверджений іншим джерелом (базовою колонкою), знімає сумнів.

    effective_date - РІВНО НАСТУПНИЙ день ПІСЛЯ останньої наявної
    колонки-дати табеля (типово - 1 число НАСТУПНОГО місяця, якого табель
    ще не охоплює) - за прямою вказівкою користувача, теж мовчки
    пропускається, БЕЗ запису на ручну перевірку: реальний випадок -
    рапорт ОСТАННЬОГО дня місяця описує складену подію ("виписаний ...
    та проходить ВЛК ..., потребує відпустки ..." - content/
    daily_report_reader._discharge_vlk_leave_events), де відпустка
    (друга з двох подій) обчислюється як "наступний день" - людина ЩЕ НЕ
    підтвердила фактичний початок відпустки НІЧИМ, окрім цього
    обчислення (на відміну від "потребує" з явною датою кимось
    "з ДАТА направляється..." - там ДАТА - з самого тексту, а не
    обчислена), і сам НАСТУПНИЙ місяць ЩЕ НЕ існує як колонки цього
    табеля - перевірити це ЗАРАЗ неможливо, а сама подія НЕ помилка (typo
    дати дав би дату, що НЕ дорівнює РІВНО "останній день+1" - той
    випадок і далі лишається на ручну перевірку нижче). Підтвердження
    прийде ОКРЕМО, коли з'явиться рапорт наступного місяця (content/
    daily_report_reader._confirmed_health_leave_event чи звичайний
    перенос статусу)."""
    person_rows = timesheet.person_rows()
    date_col_by_date = {d: idx for idx, d in timesheet.date_columns}
    ordered_dates = [d for _, d in timesheet.date_columns]
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]

    unresolved = []
    events_by_person_date = {}
    for event in events:
        normalized = normalize_name(event["pib_raw"])
        if normalized not in person_rows:
            candidate_normalized, candidate_pib_raw = _typo_match(timesheet, person_rows, pib_col, normalized)
            if candidate_normalized is None:
                unresolved.append({
                    "reason": f"ПІБ \"{event['pib_raw']}\" з рапорту не знайдено в {timesheet.file_path}.",
                    **event,
                })
                continue
            unresolved.append({
                "reason": (
                    f"ПІБ \"{event['pib_raw']}\" з рапорту не збігається буквально з роcтеровим написанням "
                    f"\"{candidate_pib_raw}\" - схоже на друкарську помилку в рапорті, статус застосовано "
                    "автоматично за схожістю написання (перевірте, чи це справді та сама людина)."
                ),
                **event,
            })
            normalized = candidate_normalized
        if event["effective_date"] not in date_col_by_date:
            if event["effective_date"] < ordered_dates[0]:
                baseline_col = timesheet.date_columns[0][0]
                baseline_value = timesheet.ws.cell(row=person_rows[normalized], column=baseline_col).value
                if baseline_value == event["status"]:
                    continue
            elif event["effective_date"] == ordered_dates[-1] + timedelta(days=1):
                continue
            unresolved.append({
                "reason": (
                    f"Дата \"{event['effective_date'].strftime('%d.%m.%Y')}\" для \"{event['pib_raw']}\" "
                    f"відсутня серед колонок-дат {timesheet.file_path}."
                ),
                **event,
            })
            continue

        key = (normalized, event["effective_date"])
        previous = events_by_person_date.get(key)
        if previous is not None and previous["status"] != event["status"]:
            override_status = STATUS_CONFLICT_OVERRIDES.get(frozenset({previous["status"], event["status"]}))
            if override_status is not None:
                events_by_person_date[key] = event if event["status"] == override_status else previous
                continue
            unresolved.append({
                "reason": (
                    f"\"{event['pib_raw']}\" {event['effective_date'].strftime('%d.%m.%Y')}: суперечливі статуси "
                    f"в рапорті - \"{previous['status']}\" і \"{event['status']}\" (застосовано останній)."
                ),
                **event,
            })
        events_by_person_date[key] = event

    if not report_dates:
        return unresolved

    last_report_date = max(report_dates)
    fill_dates = [d for d in ordered_dates if d > ordered_dates[0] and d <= last_report_date]
    # Дати ПІСЛЯ останнього наявного рапорту - жоден рапорт ще не підтвердив
    # для них "без змін" (докстрінг вище), тож САМІ ПО СОБІ вони НЕ
    # заповнюються. АЛЕ якщо для котроїсь із них у resources/ОБЛІК.xlsx ВЖЕ
    # стоїть вручну вписаний статус (реальний випадок: "БЗ", вписаний поза
    # діапазоном наявних рапортів) - цей статус ПРОДОВЖУЄТЬСЯ (forward-fill) у ВСІ наступні
    # порожні колонки, аж доки не трапиться ЩЕ ОДИН вручну вписаний статус
    # (новий "якір") чи не скінчаться колонки. Порожні колонки ДО першого
    # такого запису лишаються порожніми (як і раніше) - немає підтвердженого
    # "якоря", з якого починати продовження.
    tail_dates = [d for d in ordered_dates if d > last_report_date]

    for normalized, row_idx in person_rows.items():
        baseline_cell = timesheet.ws.cell(row=row_idx, column=timesheet.date_columns[0][0])
        current_status = baseline_cell.value
        # Значення базової колонки НЕ змінюється (див. докстрінг вище) - лише
        # заливка, щоб вона теж відповідала кольоровій гамі решти клітинок.
        baseline_fill = _fill_for_status(current_status)
        if baseline_fill:
            baseline_cell.fill = baseline_fill
        for date_value in fill_dates:
            event = events_by_person_date.get((normalized, date_value))
            if event is not None:
                current_status = event["status"]
            else:
                # За прямою вказівкою користувача - жоден рапорт не подає
                # подію САМЕ на цю дату, АЛЕ якщо в resources/ОБЛІК.xlsx тут
                # ВЖЕ вручну вписаний ІНШИЙ статус (не порожньо, не той, що
                # переноситься з попередньої колонки) - цей статус стає новим
                # "якорем" (як і подія з рапорту) і ПРОДОВЖУЄТЬСЯ далі -
                # реальний випадок: "БЗ", вписаний вручну на дату В МЕЖАХ
                # діапазону наявних рапортів (не лише ПІСЛЯ нього).
                existing_value = timesheet.ws.cell(row=row_idx, column=date_col_by_date[date_value]).value
                if existing_value and existing_value != current_status:
                    current_status = existing_value
            cell = timesheet.ws.cell(row=row_idx, column=date_col_by_date[date_value], value=current_status)
            fill = _fill_for_status(current_status)
            if fill:
                cell.fill = fill

        tail_status = None
        for date_value in tail_dates:
            cell = timesheet.ws.cell(row=row_idx, column=date_col_by_date[date_value])
            if cell.value:
                tail_status = cell.value
            elif tail_status is not None:
                cell.value = tail_status
            if tail_status is not None:
                fill = _fill_for_status(tail_status)
                if fill:
                    cell.fill = fill

    return unresolved


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _normalize_status_text(value):
    """Верхній регістр + обрізані/згорнуті пробіли - для ПОРІВНЯННЯ статусів
    із PAYMENT_STATUS_CONVERSIONS/PAYMENT_STATUS_OVERRIDES (constants.py)
    БЕЗ урахування регістру чи зайвих пробілів. Реальний випадок: значення з
    information_unit - вручну введений текст (напр. "БПШП " із зайвим
    пробілом) - точний (не нормалізований) пошук у словнику МОВЧКИ "губив"
    застосування конвертації/заміни, без жодного повідомлення про помилку
    (values[0] чи status просто НЕ збігався з ключем словника, і клітинка
    лишалась незміненою). Саме ЗНАЧЕННЯ, що записується в клітинку, і далі
    береться з оригінального (НЕ нормалізованого) словника - ця функція лише
    для ПОШУКУ ключа, не для відображення."""
    return " ".join(value.split()).upper() if isinstance(value, str) else value


def _normalized_lookup(mapping):
    """{нормалізований_ключ: (оригінальний_ключ, значення), ...}."""
    return {_normalize_status_text(key): (key, value) for key, value in mapping.items()}


_NORMALIZED_PAYMENT_STATUS_CONVERSIONS = _normalized_lookup(PAYMENT_STATUS_CONVERSIONS)
_NORMALIZED_PAYMENT_STATUS_OVERRIDES = {
    _normalize_status_text(base_status): _normalized_lookup(override_map)
    for base_status, override_map in PAYMENT_STATUS_OVERRIDES.items()
}
_NORMALIZED_PAYMENT_STATUS_PRIORITY = {
    frozenset(_normalize_status_text(status) for status in pair): winner
    for pair, winner in PAYMENT_STATUS_PRIORITY.items()
}

# Скільки РІЗНИХ ключів PAYMENT_STATUS_CONVERSIONS дають ОДНЕ й ТЕ САМЕ
# значення (напр. "БЗ"/"полон"/"Інт" - усі троє в "100_БЗ") - для цих
# конвертація фарбує клітинку ЗА НОВИМ значенням (STATUS_COLORS["100_БЗ"]),
# а НЕ за колишнім статусом ("БЗ"/"полон"/"Інт" мають РІЗНІ власні кольори) -
# реальний випадок: клітинка "100_БЗ" виглядала б ТРЬОМА РІЗНИМИ кольорами
# залежно від того, з ЯКОГО статусу її сконвертовано, хоча текст той самий.
# Конвертації з ЄДИНИМ джерелом (напр. "ППД" -> 10) і далі фарбуються ЗА
# КОЛИШНІМ статусом (підтверджено користувачем раніше) - тут неоднозначності
# немає, бо жоден ІНШИЙ статус не веде до того самого значення.
_CONVERSION_VALUE_COUNTS = Counter(PAYMENT_STATUS_CONVERSIONS.values())


def _resolve_value_priority_conflict(values):
    """values, АЛЕ ЯКЩО серед них (нормалізовано) РІВНО ДВА РІЗНІ значення, і
    ЦЯ ПАРА - ключ у PAYMENT_STATUS_PRIORITY (constants.py, той самий
    об'єкт, керований користувачем САМОСТІЙНО) - повертає [переможець]
    (один елемент) ЗАМІСТЬ оригінального списку: конфлікт МІЖ РІЗНИМИ
    файлами information_unit для ОДНІЄЇ людини й дати РОЗВ'ЯЗУЄТЬСЯ
    автоматично на користь переможця (переможець далі проходить ЗВИЧАЙНИМ
    шляхом - PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS] нижче, як і будь-яке
    ІНШЕ однозначне значення), замість "не єдине значення" (ручна
    перевірка). Реальний випадок, підтверджений користувачем: "БЗ" (зниклий
    безвісті) МАЄ переважати "Розп" (розпорядження).

    values лишається БЕЗ ЗМІН, якщо унікальних (нормалізованих) значень НЕ
    рівно два (0/1 - нема конфлікту взагалі; 3+ - неоднозначно, жодна проста
    пара цього не покриває) чи для ЦІЄЇ пари немає запису в
    PAYMENT_STATUS_PRIORITY."""
    normalized_values = {_normalize_status_text(value) for value in values}
    if len(normalized_values) != 2:
        return values
    winner = _NORMALIZED_PAYMENT_STATUS_PRIORITY.get(frozenset(normalized_values))
    return [winner] if winner is not None else values


def apply_payment_values(timesheet, payment_values):
    """Для КОЖНОЇ клітинки з якимось зі статусів PAYMENT_STATUS_CONVERSIONS
    (constants.py - напр. "ППД"/"БПШП") - ОДРАЗУ, без жодних умов, підставляє
    ФІКСОВАНЕ значення з цього словника (той самий статус завжди дає той
    самий запис, незалежно від файлів information_unit).

    Для КОЖНОЇ клітинки з якимось зі статусів PAYMENT_STATUS_OVERRIDES
    (constants.py - напр. "ВД", DEFAULT_STATUS) - якщо payment_values для
    цієї людини й дати дає РІВНО ОДНЕ значення, і воно - ключ у мапі
    прийнятних замінників для цього статусу (PAYMENT_STATUS_OVERRIDES
    [status]) - ЗАМІНЮЄ статус на ВІДПОВІДНЕ значення з мапи (НЕ обов'язково
    той самий текст, що прийшов від information_unit - реальні випадки:
    людина зі статусом "ВД" в ОБЛІК, файл information_unit каже
    "Навч" - "ВД" замінюється на буквальний текст "Навч"; людина зі
    статусом "РВЗ" (ще не визначено), файл information_unit каже "БЗВП" чи
    "Адап" - це, по суті, ТА САМА причина, що й "ППД" (категорія виплати
    10), тож замінюється на ЧИСЛО 10, а не на текст "БЗВП"/"Адап"; той самий
    "РВЗ", файл каже "ЗВФН" чи "ЛМР" - замінюється на буквальний текст, той
    самий принцип, що й "ВД" -> "Навч"). Заливка - за НОВИМ значенням
    (STATUS_COLORS, за прямою вказівкою користувача - КОЖНА категорія
    виплати (10/30/70/100/170) має ВЛАСНИЙ колір там само, не лише текстові
    статуси). Для "ВД"-подібних статусів (не DEFAULT_STATUS) - будь-яка
    суперечність (кілька РІЗНИХ значень) чи значення, якого немає серед
    ключів мапи, - статус лишається як є, без жодного запису на ручну
    перевірку ("ВД" сам собою не потребує уваги, це не "ще не визначений"
    стан). Для DEFAULT_STATUS - якщо ОДНОЗНАЧНОГО замінника немає, обробка
    ПРОДОВЖУЄТЬСЯ (див. нижче) - на відміну від "ВД", "РВЗ" ДІЙСНО потребує
    уваги, якщо його не вдалось замінити НІ звідси, НІ числом нижче.

    Для КОЖНОЇ клітинки зі статусом DEFAULT_STATUS ("РВЗ") АБО зовсім
    ПОРОЖНЬОЇ (None - реальний випадок, підтверджений користувачем:
    новоприбула людина ЩОЙНО додана до роcтера, ще БЕЗ жодного статусу - ні
    базового, ні з рапорту, бо в самих рапортах вона взагалі не фігурує,
    лише в information_unit; порожньо тут означає ТЕ САМЕ "ще не визначено",
    що й "РВЗ"), для якої НЕ спрацював жоден замінник вище, - підставляє
    категорію виплати з payment_values ({(normalize_name(ПІБ), date):
    [значення, ...]} - content/information_unit_reader.read_payment_values).
    Перед тим - _resolve_value_priority_conflict: якщо РІЗНІ файли
    information_unit подають РІВНО ДВА суперечливих значення, і ЦЯ пара -
    ключ у PAYMENT_STATUS_PRIORITY (constants.py) - конфлікт РОЗВ'ЯЗУЄТЬСЯ
    автоматично на користь переможця (напр. "БЗ" переважає "Розп"), і далі
    оброблюється як звичайне ОДНОЗНАЧНЕ значення. Підставляється, лише якщо
    (після цього) лишилось РІВНО ОДНЕ ЧИСЛОВЕ значення - будь-яка ІНША
    суперечність (кілька РІЗНИХ значень з різних файлів підрозділу, не
    покритих PAYMENT_STATUS_PRIORITY) чи нечислове значення (текстовий
    статус, якого немає серед ключів PAYMENT_STATUS_OVERRIDES
    [DEFAULT_STATUS]) НЕ застосовується мовчки для "РВЗ", а лишається на
    ручну перевірку.

    За прямою вказівкою користувача - ДЛЯ ПОРОЖНЬОЇ клітинки (не "РВЗ") це
    останнє правило М'ЯКШЕ: ОДНЕ узгоджене (не суперечливе) значення
    information_unit, навіть нечислове й БЕЗ окремого запису серед ключів
    PAYMENT_STATUS_OVERRIDES[DEFAULT_STATUS] (напр. "ВП"/"СЗЧ" - самі по
    собі ВЖЕ змістовні, розпізнавані статуси), застосовується НАПРЯМУ як
    буквальний текст, МОВЧКИ (без запису на ручну перевірку) - новоприбула
    людина ще НЕ має ЖОДНОГО іншого джерела статусу, тож немає з чим
    "сперечатись". Лише СПРАВЖНЯ суперечність (2+ РІЗНИХ значення для
    ОДНІЄЇ дати) лишається непідставленою - але ТЕЖ МОВЧКИ (не записується
    на КОЖНУ таку дату окремо). "Підрозділ ще не подав дані" (людина/дата,
    якої немає в payment_values узагалі) записується на КОЖНУ таку дату
    лише для "РВЗ" (людина ВЖЕ відстежується, категорії просто ще нема) -
    для ПОРОЖНЬОЇ клітинки цей запис НЕ повторюється на кожну з ~30
    колонок-дат окремо (надто галасливо для новоприбулої людини без
    жодного джерела взагалі); замість цього - ОДИН підсумковий запис
    наприкінці (нижче), якщо ЖОДНА дата для цієї людини так і не отримала
    жодного значення.

    Увесь пошук статусу/значення в PAYMENT_STATUS_CONVERSIONS/PAYMENT_
    STATUS_OVERRIDES - БЕЗ урахування регістру чи зайвих пробілів
    (_normalize_status_text) - значення information_unit - вручну введений
    текст, реальний ризик зайвого пробілу чи іншого регістру, що інакше
    МОВЧКИ "губив" би застосування (без жодного повідомлення про помилку).

    Наприкінці - ЯКЩО для людини ЖОДНА колонка-дата так і не отримала
    ЖОДНОГО значення (ні базового статусу в ОБЛІК.xlsx, ні події з рапорту,
    ні жодного значення information_unit узагалі) - ОДИН явний запис на
    ручну перевірку: людині потрібен базовий статус, вписаний вручну. Якщо ж
    information_unit ДАВ хоч ОДНЕ значення для ХОЧА Б ОДНІЄЇ дати - цей
    запис НЕ додається (за прямою вказівкою користувача - саме ця зміна
    прибирає раніші хибні "усі колонки-дати порожні" повідомлення для
    новоприбулих людей, чиї дані information_unit ВЖЕ фактично підставив).

    Повертає список {"reason", ...} - записи, що потребують ручної перевірки."""
    pib_col = timesheet.label_columns[PIB_COLUMN_NAME]
    unresolved = []

    for normalized, row_idx in timesheet.person_rows().items():
        pib_raw = timesheet.ws.cell(row=row_idx, column=pib_col).value
        has_any_value = False
        for col_idx, date_value in timesheet.date_columns:
            cell = timesheet.ws.cell(row=row_idx, column=col_idx)
            status = cell.value
            if status is not None:
                has_any_value = True
            normalized_status = _normalize_status_text(status)

            conversion_entry = _NORMALIZED_PAYMENT_STATUS_CONVERSIONS.get(normalized_status)
            if conversion_entry is not None:
                original_status, new_value = conversion_entry
                # Кілька РІЗНИХ статусів, що ведуть до ОДНОГО й ТОГО САМОГО
                # значення (напр. "БЗ"/"полон"/"Інт" - усі троє в "100_БЗ") -
                # фарбуються ЗА НОВИМ значенням (щоб та сама клітинка "100_БЗ"
                # завжди виглядала ОДНАКОВО, незалежно від походження); статуси
                # з ЄДИНИМ джерелом (напр. "ППД" -> 10) - і далі за КОЛИШНІМ
                # статусом (_CONVERSION_VALUE_COUNTS вище).
                color_key = new_value if _CONVERSION_VALUE_COUNTS[new_value] > 1 else original_status
                cell.value = new_value
                fill = _fill_for_status(color_key)
                if fill:
                    cell.fill = fill
                has_any_value = True
                continue

            values = _resolve_value_priority_conflict(payment_values.get((normalized, date_value)) or [])
            unique_values = set(values)

            # Порожня клітинка + ОДНЕ узгоджене значення information_unit,
            # яке САМЕ ПО СОБІ - ключ PAYMENT_STATUS_CONVERSIONS (напр.
            # "ППД"/"БПШП"/"ВПБП"/"БЗ"/"полон"/"Інт" - ТОЙ САМИЙ словник, що
            # застосовується ОДРАЗУ вище, коли ЦЕ статус САМОЇ клітинки) -
            # конвертується ТИМ САМИМ шляхом, навіть коли це лише значення
            # information_unit для новоприбулої людини: реальний випадок,
            # підтверджений користувачем - "ВПБП" МАЄ стати "100_ВПБП", той
            # самий факт, що й коли роcтер сам написав би "ВПБП". Лише для
            # ПОРОЖНЬОЇ клітинки (не "РВЗ") - "РВЗ" сама по собі НІКОЛИ не є
            # ключем PAYMENT_STATUS_CONVERSIONS, тож для неї це нічого не
            # змінює.
            if normalized_status is None and len(unique_values) == 1:
                conversion_entry = _NORMALIZED_PAYMENT_STATUS_CONVERSIONS.get(_normalize_status_text(values[0]))
                if conversion_entry is not None:
                    original_status, new_value = conversion_entry
                    color_key = new_value if _CONVERSION_VALUE_COUNTS[new_value] > 1 else original_status
                    cell.value = new_value
                    fill = _fill_for_status(color_key)
                    if fill:
                        cell.fill = fill
                    has_any_value = True
                    continue

            # Порожня клітинка (normalized_status is None) шукає override_map
            # ЗА DEFAULT_STATUS ("РВЗ") - _NORMALIZED_PAYMENT_STATUS_OVERRIDES
            # не має власного запису для "None" (лише для реальних текстових
            # статусів), а порожньо тут ЗНАЧИТЬ ТЕ САМЕ "ще не визначено", що
            # й "РВЗ" (див. докстрінг вище) - без цього текстові замінники
            # ("НОВ" -> "НОВ" тощо) мовчки НЕ спрацьовували б для новоприбулих
            # людей, і кожна така дата хибно "падала" б у "не єдине числове
            # значення" нижче.
            override_map = _NORMALIZED_PAYMENT_STATUS_OVERRIDES.get(normalized_status or DEFAULT_STATUS)
            if override_map:
                # Порівняння - за НОРМАЛІЗОВАНИМИ формами (не сирими values)
                # - реальний випадок: ДВА файли кажуть "БПШП"/"бпшп" (різний
                # регістр) - це ТА САМА відповідь, а не суперечність.
                normalized_values = {_normalize_status_text(value) for value in values}
                if len(normalized_values) == 1:
                    override_entry = override_map.get(next(iter(normalized_values)))
                    if override_entry is not None:
                        _original_value, new_status = override_entry
                        cell.value = new_status
                        fill = _fill_for_status(new_status)
                        if fill:
                            cell.fill = fill
                        has_any_value = True
                        continue

            # ПОРОЖНЯ клітинка (status is None) - той самий шлях підстановки
            # категорії виплати, що й DEFAULT_STATUS ("РВЗ") нижче: реальний
            # випадок, підтверджений користувачем - новоприбулу людину додано
            # до роcтера, але ЩЕ БЕЗ жодного статусу (ні базового, ні з
            # рапорту - вона в рапортах узагалі не фігурує, лише в
            # information_unit), тож "порожньо" тут означає ТЕ САМЕ "ще не
            # визначено", що й "РВЗ" - information_unit має шанс заповнити її
            # так само.
            if normalized_status not in (DEFAULT_STATUS, None):
                continue

            date_text = date_value.strftime("%d.%m.%Y")
            if not values:
                # За прямою вказівкою користувача - "підрозділ ще не подав
                # дані" МАЄ сенс лише для "РВЗ" (людина ВЖЕ відстежується,
                # категорії просто ще нема на ЦЮ дату). Для ПОРОЖНЬОЇ
                # клітинки (новоприбула людина без жодних інших даних) - НЕ
                # записується окремо на КОЖНУ таку дату (мовчки пропускається
                # тут - інакше на людину без жодного інформ_unit джерела
                # припадав би запис на КОЖНУ з ~30 колонок-дат); замість
                # цього - ОДИН підсумковий запис "усі колонки порожні"
                # наприкінці, якщо ЖОДНА дата взагалі нічого не отримала.
                if normalized_status == DEFAULT_STATUS:
                    unresolved.append({
                        "reason": f"Немає категорії виплати для \"{pib_raw}\" на {date_text} - підрозділ ще не подав дані.",
                        "pib_raw": pib_raw, "report_date": date_value,
                    })
                continue

            if len(unique_values) > 1 or not _is_number(values[0]):
                if normalized_status is None:
                    # Порожня клітинка + ОДНЕ узгоджене (не суперечливе)
                    # значення information_unit, яке НЕ число і не підійшло
                    # під жодну з мап вище - реальний випадок, підтверджений
                    # користувачем: "ВП"/"СЗЧ" тощо - самі по собі ВЖЕ
                    # змістовні, розпізнавані статуси (STATUS_COLORS), просто
                    # без окремого запису для КОНВЕРТАЦІЇ (не кожен статус її
                    # потребує) - для новоприбулої людини БЕЗ жодного іншого
                    # джерела застосовується НАПРЯМУ, буквальним текстом,
                    # МОВЧКИ (за прямою вказівкою користувача - жодного
                    # запису на ручну перевірку на КОЖНУ таку дату). Справжня
                    # суперечність (2+ РІЗНИХ значення) і далі лишається
                    # непідставленою - немає підстави вибрати ОДНЕ з них, але
                    # так само МОВЧКИ (той самий принцип, що й "немає даних"
                    # вище - для порожньої клітинки жоден варіант тут не
                    # записується на КОЖНУ дату окремо).
                    if len(unique_values) == 1:
                        cell.value = values[0]
                        has_any_value = True
                        fill = _fill_for_status(values[0])
                        if fill:
                            cell.fill = fill
                    continue

                unresolved.append({
                    "reason": (
                        f"\"{pib_raw}\" {date_text}: підрозділ подає {sorted(map(str, unique_values))} - "
                        f"не єдине числове значення категорії виплати, статус \"{normalized_status}\" лишено як є."
                    ),
                    "pib_raw": pib_raw, "report_date": date_value,
                })
                continue

            cell.value = values[0]
            has_any_value = True
            fill = _fill_for_status(values[0])
            if fill:
                cell.fill = fill

        # Реальний випадок (баг, підтверджений користувачем): людину ЩОЙНО
        # додано до роcтера БЕЗ базового статусу, і ЖОДЕН рапорт її не
        # згадує - ЯКЩО information_unit ТЕЖ не дав ЖОДНОГО значення для
        # ЖОДНОЇ дати (has_any_value лишився False - жодна клітинка вище так
        # і не отримала значення) - ось ТУТ дійсно нема на чому будувати
        # подальшу обробку, тож ОДИН явний запис на ручну перевірку. Якщо ж
        # information_unit ДАВ хоч щось - has_any_value True, цей запис НЕ
        # додається (саме та зміна поведінки, про яку прямо попросив
        # користувач - "не має бути цих помилок в txt", коли підставилось
        # значення з information_unit).
        if not has_any_value:
            unresolved.append({
                "reason": (
                    f"\"{pib_raw}\": усі колонки-дати порожні - немає базового статусу в {timesheet.file_path} "
                    "і жодних даних від information_unit. Впишіть базовий статус уручну."
                ),
                "pib_raw": pib_raw, "report_date": None,
            })

    return unresolved
