import math
import os
import re
from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

from constants import (
    OUTPUT_DIR,
    MONEY_REPORT_STATIC,
    SHORT_UNIT_BATTALION,
    HIGHER_COMMANDER_TITLE,
    MONEY_REPORT_CATEGORIES,
    BASIS_REQUIRED_POINT_COLUMN_NAMES,
    MONTH,
    YEAR,
)
from utils.logging_utils import print_green, print_purple, print_red
from formatting.docx_utils import (
    set_margins,
    add_page_number_header,
    add_paragraph_with_style,
    add_blank_paragraphs,
    add_signature_block,
    create_table,
    save_docx_safely,
)
from content.br_helpers import find_higher_commander
from content.money_report_helpers import (
    month_nominative_upper,
    month_genitive_lower,
    normalize_name,
    get_date_columns,
    build_category_rows,
    build_day_value_periods,
    build_pidstavy_extra_grounds_by_person,
    resolve_category_config,
    resolve_reportable_categories,
    rtgr_cross_tier_rows,
    exclude_rows_without_disappearance_date,
    _merge_preserving_file_order,
    _exclude_by_excluded_day_value,
    _blank_not_paid_days_confirmed_elsewhere,
    _blank_non_szch_days_for_szch_started_this_month,
    _exclude_commander_and_tvo,
    _ordered_points,
    _COMBINED_TARGET_CELL_VALUES,
    _RTGR_CROSS_TIER_TARGET_POINT,
    BASIS_REQUIRED_POINTS,
)
from content.report_changes import build_changes_entries, RESTATE_POINTS
from content.report_document_reader import SPETSKONTYNGENT_VALUE

STATIK = MONEY_REPORT_STATIC

TABLE_HEADERS = ["№", "Посада", "Військове звання", "Прізвище, ім'я, по батькові", "Період участі", "Кількість днів", "Підстава для виплати"]
NOT_PAID_TABLE_HEADERS = TABLE_HEADERS[:-1] + ["Підстава для не виплати"]
TABLE_WIDTHS_CM = [0.75, 2.36, 1.67, 2.85, 2.0, 0.5, 7.00]

# "100_СПЕЦКОНТИНГЕНТ" (полон/зниклі безвісти/заручники) - ЄДИНИЙ пункт, чия
# таблиця має ІНШИЙ набір колонок замість "Період участі"/"Кількість днів":
# "Дата зникнення" (з ОБЛІК.xlsx, content.money_report_helpers.exclude_rows_
# without_disappearance_date) і "Період виплати" (сам ПЕРІОД, як завжди) -
# підтверджено користувачем; той самий формат, що вже читає (за назвою
# заголовка колонки, не фіксованою позицією) content.report_document_reader.
# Ширини 5-ї/6-ї колонок (на відміну від решти) НЕ ті самі, що в
# TABLE_WIDTHS_CM: "Дата зникнення" - коротка (ДД.ММ.РРРР), а "Період
# виплати" - повний діапазон (ДД.ММ.РРРР-ДД.ММ.РРРР, як звичайний "Період
# участі") - тож саме вона отримує ширшу колонку, а не навпаки, як було б
# при простому перейменуванні заголовків без зміни ширин.
SPETSKONTYNGENT_TABLE_HEADERS = TABLE_HEADERS[:4] + ["Дата зникнення", "Період виплати", "Примітка (підстави)"]
SPETSKONTYNGENT_TABLE_WIDTHS_CM = TABLE_WIDTHS_CM[:4] + [0.85, 1.65] + TABLE_WIDTHS_CM[6:]
BASIS_COLUMN_INDEX = 6
_TABLE_FONT_SIZE_PT = 10  # кегль тексту в табличках рапорту - ЗАВЖДИ, без стискання (підтверджено користувачем)


# Оцінка висоти рядка таблиці - python-docx не має рушія розмітки (реальне перенесення
# слів визначає лише Word під час відкриття файлу), тож нижче - наближені емпіричні
# коефіцієнти для Times New Roman, а не точний розрахунок.
_CM_PER_PT = 2.54 / 72
_CELL_MARGIN_H_CM = 0  # внутрішній лівий+правий відступ комірки (create_table викликається з cell_margin_cm=0)
_CELL_MARGIN_V_CM = 0  # внутрішній верхній+нижній відступ комірки (create_table викликається з cell_margin_cm=0)
_AVG_CHAR_WIDTH_FACTOR = 0.52  # середня ширина символу відносно кегля (Times New Roman, кирилиця)
_LINE_HEIGHT_FACTOR = 1.15  # висота рядка відносно кегля при одинарному міжрядковому інтервалі


def _estimate_row_height_cm(row_values, widths_cm, font_size_pt):
    """Наближено оцінює висоту (см) рядка таблиці з переданими значеннями колонок -
    бере найвищу з комірок (те саме робить і сам Word, бо висота рядка = висота
    найвищої комірки в ньому)."""
    char_width_cm = font_size_pt * _CM_PER_PT * _AVG_CHAR_WIDTH_FACTOR
    line_height_cm = font_size_pt * _CM_PER_PT * _LINE_HEIGHT_FACTOR

    max_lines = 1
    for value, width_cm in zip(row_values, widths_cm):
        usable_width_cm = max(width_cm - 2 * _CELL_MARGIN_H_CM, char_width_cm)
        cell_lines = sum(
            max(1, math.ceil(len(segment) * char_width_cm / usable_width_cm)) if segment else 1
            for segment in str(value).split("\n")
        )
        max_lines = max(max_lines, cell_lines)

    return max_lines * line_height_cm + 2 * _CELL_MARGIN_V_CM


def _usable_page_height_cm(doc):
    # section.page_height - section.top_margin - ... повертає ЗВИЧАЙНИЙ int (в EMU) - Length
    # не перевизначає арифметичні оператори, тож віднімати треба вже готові .cm-значення.
    section = doc.sections[0]
    return section.page_height.cm - section.top_margin.cm - section.bottom_margin.cm


def _usable_text_width_cm(doc):
    """Ширина (см) тексту НЕ в таблиці (звичайного абзацу на всю сторінку) -
    ширина сторінки мінус лівий+правий відступи."""
    section = doc.sections[0]
    return section.page_width.cm - section.left_margin.cm - section.right_margin.cm


def _estimate_paragraph_height_cm(text, font_size_pt, usable_width_cm):
    """Оцінка висоти (см) звичайного (не табличного) абзацу на всю ширину
    сторінки - та сама логіка переносу рядків, що й для комірки таблиці
    (_estimate_row_height_cm), лише з однією "колонкою" на всю ширину тексту."""
    return _estimate_row_height_cm([text], [usable_width_cm], font_size_pt)


HEADER_ROW_HEIGHT_CM = 2.35  # той самий header_row_height_cm, що передається в create_table
_HEADING_FONT_SIZE_PT = 13.5  # кегль заголовків пунктів (секцій) - передається в add_paragraph_with_style

# Запас на неточність оцінки висоти (python-docx не має рушія розмітки, тож реальна
# висота в Word може трохи відрізнятись від оціненої) - для рішення про заповнення
# сторінки рахуємо, що на ній є трохи менше місця, ніж є насправді.
_PAGE_FIT_SAFETY_MARGIN = 0.90


def _row_height_cm(row, font_size_pt, use_disappearance_date=False):
    """Висота (см) одного рядка таблиці - висота найвищої з його комірок (так
    само, як рахує сам Word). use_disappearance_date=True (таблиця
    "100_СПЕЦКОНТИНГЕНТ") - "Дата зникнення"/"Період" замість "Період"/"Дні",
    із ШИРИНАМИ SPETSKONTYNGENT_TABLE_WIDTHS_CM (інакше оцінка висоти рахувала
    б перенесення тексту за ЧУЖИМИ, звичайними ширинами колонок)."""
    if use_disappearance_date:
        values = [row["ПОСАДА"], row["ЗВАННЯ"], row["ПІБ"], row["ДАТА_ЗНИКНЕННЯ"], row["ПЕРІОД"], row["ПІДСТАВА"]]
        widths_cm = SPETSKONTYNGENT_TABLE_WIDTHS_CM[1:]
    else:
        values = [row["ПОСАДА"], row["ЗВАННЯ"], row["ПІБ"], row["ПЕРІОД"], str(row["ДНІ"]), row["ПІДСТАВА"]]
        widths_cm = TABLE_WIDTHS_CM[1:]
    return _estimate_row_height_cm(values, widths_cm, font_size_pt)


def _place_rows_for_page_fill(rows, usable_height_cm, initial_used_cm, font_size_pt, use_disappearance_date=False):
    """Розставляє рядки по (симульованих) сторінках так, щоб порожнього місця не
    лишалось без потреби: якщо наступний за порядком рядок не влазить у залишок
    сторінки, а десь далі є менший, що влазить - саме він підставляється замість
    нього (найщільніше заповнює залишок) - природний порядок (як у файлі
    ОБЛІК.xlsx) зберігається завжди, коли й так усе влазить, переставляються лише
    рядки, потрібні саме для того, щоб забрати порожнечу. Кегль ЗАВЖДИ лишається
    font_size_pt - жодного стискання шрифту (підтверджено користувачем: у
    табличках рапорту й "змін в наказі" шрифт має бути рівно 10, без винятків);
    рядок, що не влазить навіть на щойно почату чисту сторінку, просто йде на неї
    як є - Word сам розіб'є таку таблицю на кілька сторінок природним чином.

    Повертає (список рядків у фінальному порядку таблиці, скільки см ОСТАННЬОЇ
    (симульованої) сторінки зайнято після розміщення всіх рядків) - друге
    значення передається як initial_used_cm ДЛЯ НАСТУПНОГО пункту (70к/100к/170к
    і т.д. в реальному документі часто продовжуються на тій самій фізичній
    сторінці, де закінчився попередній пункт, а не завжди починаються з чистої)."""
    remaining = list(rows)
    heights = [_row_height_cm(row, font_size_pt, use_disappearance_date) for row in remaining]
    placements = []
    page_used_cm = initial_used_cm
    just_reset = False

    while remaining:
        remaining_space_cm = usable_height_cm - page_used_cm
        candidates = [idx for idx, h in enumerate(heights) if h <= remaining_space_cm]

        if heights[0] <= remaining_space_cm:
            chosen = 0
        elif candidates:
            chosen = max(candidates, key=lambda idx: heights[idx])
        elif not just_reset:
            # Жоден рядок не влазить у залишок цієї сторінки - переходимо на
            # нову (чисту) і пробуємо ще раз саме з неї.
            page_used_cm = 0.0
            just_reset = True
            continue
        else:
            # Навіть на щойно почату чисту сторінку перший за порядком рядок не
            # влазить (величезна підстава) - розміщуємо як є, він природно
            # розтягнеться на кілька сторінок.
            chosen = 0

        just_reset = False
        row = remaining.pop(chosen)
        height_cm = heights.pop(chosen)
        placements.append(row)
        page_used_cm += height_cm

    return placements, page_used_cm


def _center_basis_column(table):
    """Текст у колонці "Підстава для виплати" (BASIS_COLUMN_INDEX) - по центру
    комірки, і горизонтально, і вертикально (вертикальне центрування вже є в
    усіх комірках з create_table(..., cell_alignment=WD_ALIGN_PARAGRAPH.CENTER) -
    тут лише горизонтальне, явно, для цієї колонки)."""
    for row in table.rows:
        for paragraph in row.cells[BASIS_COLUMN_INDEX].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _build_combined_section(rows_with_data, date_columns, status_lookup, point, target_cell_values, categories=None, extra_grounds_by_person=None):
    """Будує рядки для одного пункту: категорії categories[point] (за замовчуванням -
    constants.MONEY_REPORT_CATEGORIES; звіт на командира/ТВО передає
    constants.COMMANDER_MONEY_REPORT_CATEGORIES) і порядок їх обходу визначає
    resolve_reportable_categories (content/money_report_helpers.py) - див. її docstring
    щодо named-категорій/catch-all/"include_to_report". Рядки з усіх категорій
    об'єднуються через _merge_preserving_file_order - порядок точно як в ОБЛІК.xlsx,
    а не згруповано по категоріях.

    target_cell_values - які значення комірки ОБЛІК.xlsx рахуються як "день участі"
    для ЦЬОГО виклику - зазвичай {point}, але для самого пункту 100 свідомо ширше
    ({70, 100, 170}), щоб дні 70/170 категорій (catch-all чи "70_РТГр"/"170_РТГр")
    також потрапляли додатковим рядком у 100_SECTION (з підставою САМЕ пункту
    100, а не 70/170) - так само, як для пункту 100 і зараз.

    extra_grounds_by_person - див. build_category_rows; тут - build_pidstavy_extra_grounds_by_person
    (колонка "ПІДСТАВИ"), передається В КОЖНУ категорію КОЖНОГО пункту однаково."""
    if categories is None:
        categories = MONEY_REPORT_CATEGORIES
    specs = resolve_reportable_categories(point, categories)
    rows_lists = [
        build_category_rows(
            rows_with_data, date_columns, target_cell_values, status_lookup,
            status_filter=spec["status_filter"], category_config=resolve_category_config(point, spec["category_name"], categories),
            exclude_status=spec["exclude_status"], category_name=spec["category_name"],
            extra_grounds_by_person=extra_grounds_by_person, categories=categories,
        )
        for spec in specs
    ]
    return _merge_preserving_file_order(rows_with_data, *rows_lists)


# Точка, що завжди присутня в результаті першою, навіть порожня - історична конвенція
# документа (перший пункт рапорту показаний завжди, незалежно від того, чи є рядки).
_ALWAYS_PRESENT_POINT = 30

# Псевдо-пункт "Не виплачувати додаткову винагороду..." (наказ МО України №260 від
# 07.06.2023, п.14 розд. ХХХІV) - звичайний пункт MONEY_REPORT_CATEGORIES/
# COMMANDER_MONEY_REPORT_CATEGORIES (constants.py), з категоріями "СЗЧ"/"Задув."
# (той самий формат, що й у решти пунктів) - будується ТІЄЮ САМОЮ логікою
# (_build_combined_section/build_category_rows), нічим не відрізняючись від,
# напр., "100_ШП". Використовується тут лише для РЕНДЕРИНГУ - цей пункт має
# з'являтись ОДРАЗУ ПІСЛЯ звичайних пунктів ОБРАНОГО місяця, ПЕРЕД "Прошу
# внести зміни..." (обраний місяць - у пріоритеті, підтверджено користувачем;
# раніше рендерився В САМОМУ КІНЦІ документа, ПІСЛЯ ВСІХ місяців "Прошу внести
# зміни...", через що виглядав як частина ЦИХ попередніх місяців), тож
# витягується з categories_list окремо (process_generate_report_for_get_money/
# generate_report_for_commander_money.py), а не рендериться в загальному циклі.
_NOT_PAID_POINT = "NOT_PAID"

# BASIS_REQUIRED_POINTS (content.money_report_helpers) - "100_СПЕЦКОНТИНГЕНТ"/
# "100_БПШП", де порожня колонка "ПІДСТАВИ" ОБЛІК.xlsx (build_pidstavy_extra_grounds_by_person)
# виключає людину з рапорту ЦІЛКОМ (а не лишає рядок із порожньою "Підстава для
# виплати", як для решти пунктів) - підтверджено користувачем: без задокументованої
# підстави (наказу про визнання зниклим безвісти/полоненим тощо) людина в цю секцію
# не потрапляє взагалі, але сам факт виключення (і за кого саме) друкується в
# термінал - див. _build_categories. Живе в content.money_report_helpers (не тут),
# бо content.report_changes теж потребує цей самий набір пунктів (ретроактивне
# додавання, коли підстава з'явилась лише в actual-файлі), а не може імпортувати
# звідси (циклічний імпорт - цей модуль сам імпортує build_changes_entries звідти).


def _build_categories(rows_with_data, date_columns, status_lookup, categories=None):
    """Повертає впорядкований список (шаблон, рядки) лише для непорожніх пунктів, у
    ФІКСОВАНОМУ пріоритеті (_ordered_points, content/money_report_helpers.py) - 30,
    100, 70, 170, 10 (а не за зростанням номера точки) - відсутній пункт просто
    пропускається, наступний за пріоритетом підіймається вище.

    categories - джерело підстав (за замовчуванням - constants.MONEY_REPORT_CATEGORIES
    для головного рапорту; звіт на командира/ТВО передає окремий, незалежно редагований
    constants.COMMANDER_MONEY_REPORT_CATEGORIES з тією ж структурою).

    ПОВНІСТЮ ДИНАМІЧНО, за фактичними ключами categories - які саме точки існують, набір
    названих категорій усередині кожної (і яка з них catch-all, за "default"), і сам
    номер точки (повертається як РЯДОК - шукається в STATIK["SECTIONS"]/["COMMANDER_SECTIONS"]
    у _add_category_section) - усе вираховується із самого словника, а НЕ захардкоджено.
    Щоб додати ЗОВСІМ НОВУ точку (напр. "50") - досить дописати MONEY_REPORT_CATEGORIES[50]
    (за тим самим зразком, що й інші точки) і ключ "50" у
    data_statik_report_for_get_additional_money.SECTIONS у resources/data.json (за потреби -
    і в .COMMANDER_SECTIONS для рапорту на командира) - жодних змін коду; вона просто
    з'явиться в результаті ПІСЛЯ п'яти пріоритетних точок (за зростанням номера).

    Кожен день рахується НЕЗАЛЕЖНО (як і завжди) - людина може мати 10-дні одного
    дня й 100-дні іншого в тому самому місяці, і потрапить рядком в ОБИДВІ секції
    (10 і 100) одночасно, без жодного взаємовиключення між пунктами. "NOT_PAID"/
    "СЗЧ" - так само: день СЗЧ рахується ЛИШЕ в "NOT_PAID", решта днів людини й
    далі рахуються за своїми звичайними категоріями як завжди.

    Колонка "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ" (build_pidstavy_extra_grounds_by_person)
    рахується РІВНО один раз тут і передається в КОЖЕН виклик _build_combined_section
    для пунктів ПОЗА BASIS_REQUIRED_POINTS - тож вона завжди додається до підстави
    людини незалежно від пункту/категорії. КОЖЕН пункт BASIS_REQUIRED_POINTS -
    ОКРЕМИЙ виклик build_pidstavy_extra_grounds_by_person зі СВОЄЮ, ВЛАСНОЮ
    колонкою (BASIS_REQUIRED_POINT_COLUMN_NAMES, constants.py) - підтверджено
    користувачем: "100_СПЕЦКОНТИНГЕНТ"/"100_ВПБП"/"100_БПШП" раніше помилково
    ділили ОДНУ спільну "ПІДСТАВИ", тепер - кожен свою, незалежно від інших.

    Винятки, що потребують коду (бо це бізнес-правила, а не структура даних) -
    _COMBINED_TARGET_CELL_VALUES (див. docstring-коментар вище); _ALWAYS_PRESENT_POINT
    (30) завжди в результаті першим, навіть порожній; BASIS_REQUIRED_POINTS
    ("100_СПЕЦКОНТИНГЕНТ", "100_БПШП", "100_ВПБП") - людина без непорожньої
    ВЛАСНОЇ підстави виключається з цього пункту цілком (для решти пунктів
    "ПІДСТАВИ ДЛЯ ІНШИХ СИТУАЦІЙ" лише доповнює підставу, нічого не виключаючи),
    з попередженням у термінал про кожного виключеного; rtgr_cross_tier_rows -
    "зайвий" рядок пункту "100_РТГр" для в/сл, чий день форсовано позначений
    "70_РТГр"/"170_РТГр" (пункти 70/170), а не буквально "100_РТГр" - лишається
    РАЗОМ з рештою РТГр, а не в загальній "100" (content/money_report_helpers.py,
    _RTGR_CROSS_TIER_SOURCE_CATEGORIES)."""
    if categories is None:
        categories = MONEY_REPORT_CATEGORIES

    general_extra_grounds_by_person = build_pidstavy_extra_grounds_by_person(rows_with_data)

    categories_list = []
    for point in _ordered_points(categories):
        target_cell_values = _COMBINED_TARGET_CELL_VALUES.get(point, {point})
        if point in BASIS_REQUIRED_POINTS:
            extra_grounds_by_person = build_pidstavy_extra_grounds_by_person(
                rows_with_data, BASIS_REQUIRED_POINT_COLUMN_NAMES[point],
            )
        else:
            extra_grounds_by_person = general_extra_grounds_by_person
        rows = _build_combined_section(
            rows_with_data, date_columns, status_lookup, point, target_cell_values,
            categories=categories, extra_grounds_by_person=extra_grounds_by_person,
        )
        if point == _RTGR_CROSS_TIER_TARGET_POINT:
            cross_tier_rows = rtgr_cross_tier_rows(
                rows_with_data, date_columns, status_lookup, categories, extra_grounds_by_person,
            )
            rows = _merge_preserving_file_order(rows_with_data, rows, cross_tier_rows)
        if point in BASIS_REQUIRED_POINTS:
            kept_rows, excluded_rows = [], []
            for row in rows:
                (kept_rows if normalize_name(row["ПІБ"]) in extra_grounds_by_person else excluded_rows).append(row)
            rows = kept_rows
            for row in excluded_rows:
                print_red(f"⚠ Немає підстави для {point}: {row['ПІБ']} - виключено з рапорту.")
        # "100_СПЕЦКОНТИНГЕНТ" - ДОДАТКОВО (незалежно від наявності підстави
        # вище) вимагає "ДАТА В СТАТУС СПЕЦКОНТИНГЕНТУ"/"ДАТА ЗНИКНЕННЯ" з
        # ОБЛІК.xlsx для КОЖНОГО рядка - exclude_rows_without_disappearance_date
        # (content.money_report_helpers) сама пропускає решту пунктів без змін.
        rows = exclude_rows_without_disappearance_date(rows, point)

        if rows or point == _ALWAYS_PRESENT_POINT:
            categories_list.append((str(point), rows))

    return categories_list


def _add_intro(doc, month_int):
    add_paragraph_with_style(doc, STATIK["HIGH_COMMANDER"], alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_blank_paragraphs(doc, 3)
    add_paragraph_with_style(doc, STATIK["REPORT_TITLE"], alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_blank_paragraphs(doc, 2)
    add_paragraph_with_style(doc, STATIK["START_SECTION"], first_line_indent=34)
    add_paragraph_with_style(
        doc,
        STATIK["SECTION_INFORMATION"].format(month=month_nominative_upper(month_int), year=YEAR),
        first_line_indent=34,
    )
    add_blank_paragraphs(doc, 1)


def _add_changes_only_intro(doc):
    """Вступ окремого документа "Поправки в наказ..." (RUN_MODE_CHANGES_ONLY) - той
    самий адресат+заголовок, що й у звичайному рапорті (_add_intro), але без
    START_SECTION/SECTION_INFORMATION: той текст описує звичайний облік ПОТОЧНОГО
    місяця, що тут не застосовне - документ одразу переходить до пунктів
    "Прошу внести зміни..." (_add_changes_sections)."""
    add_paragraph_with_style(doc, STATIK["HIGH_COMMANDER"], alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_blank_paragraphs(doc, 3)
    add_paragraph_with_style(doc, STATIK["REPORT_TITLE"], alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_blank_paragraphs(doc, 2)


_MONEY_AMOUNT_RE = re.compile(r"\d[\d\s]*\d\s*грн\.\s*\d+\s*коп\.")


def _bold_money_amounts(text):
    """Сума виплати (напр. "50 000 грн. 00 коп.") у заголовку пункту - жирним,
    для БУДЬ-ЯКОГО пункту (регулярний вираз за форматом суми, а не хардкод по
    номеру точки) - щоб нова точка (напр. майбутня 50) отримувала це форматування
    автоматично, без змін коду, як і решта "нульових змін коду" гарантій цього
    модуля. Текст без суми (напр. підпункти розділу змін) - порожній список,
    add_paragraph_with_style тоді просто не виділяє нічого."""
    return _MONEY_AMOUNT_RE.findall(text)


def _add_heading_and_table_section(
    doc, heading_text, rows, carried_over_cm, headers=TABLE_HEADERS, widths_cm=TABLE_WIDTHS_CM, use_disappearance_date=False,
):
    """Заголовок пункту (вже готовим текстом - "{номер}. {текст}" для звичайних
    пунктів рапорту, чи "{номер}.{підномер} {текст}" для підпунктів "Прошу внести
    зміни..." - див. _add_category_section/_add_changes_subsection) + таблиця з
    rows, з тим самим заповненням сторінок, що й для звичайних пунктів - щоб
    новий розділ "зміни за попередні місяці" виглядав так само, як решта
    документа.

    headers/widths_cm - заголовки/ширини колонок таблиці (за замовчуванням -
    TABLE_HEADERS/TABLE_WIDTHS_CM); _add_category_section передає
    NOT_PAID_TABLE_HEADERS для _NOT_PAID_POINT (та сама структура, лише інша
    остання колонка - "Підстава для не виплати", а не "для виплати"), чи
    SPETSKONTYNGENT_TABLE_HEADERS/_WIDTHS_CM для "100_СПЕЦКОНТИНГЕНТ".

    use_disappearance_date - таблиця "100_СПЕЦКОНТИНГЕНТ": "Дата зникнення"/
    "Період" у 5-й/6-й колонці замість "Період"/"Дні" (той самий зсув, що вже
    очікує content.report_document_reader._detect_columns при ЧИТАННІ готового
    рапорту) - підтверджено користувачем.

    carried_over_cm - скільки см ПОТОЧНОЇ (симульованої) сторінки вже зайнято
    ДО цього пункту (з попереднього пункту чи вступу) - у реальному документі
    пункт 2/3/4 здебільшого продовжується на тій самій фізичній сторінці, де
    закінчився попередній, а не завжди починається з чистої. Без цього
    заповнення "з нуля" для об'єднання/переупорядкування вважало б, що на
    сторінці лишається значно більше вільного місця, ніж є насправді - і
    залишок реальної сторінки (під довгим заголовком пункту) лишався б порожнім.

    Повертає, скільки см ОСТАННЬОЇ сторінки зайнято після цього пункту (разом з
    порожнім абзацом після таблиці) - для передачі як carried_over_cm наступному."""
    add_paragraph_with_style(
        doc, heading_text, first_line_indent=34, heading_level=1, font_size=_HEADING_FONT_SIZE_PT,
        bold=_bold_money_amounts(heading_text),
    )

    usable_height_cm = _usable_page_height_cm(doc) * _PAGE_FIT_SAFETY_MARGIN
    usable_width_cm = _usable_text_width_cm(doc)
    heading_height_cm = _estimate_paragraph_height_cm(heading_text, _HEADING_FONT_SIZE_PT, usable_width_cm)
    initial_used_cm = carried_over_cm + heading_height_cm + HEADER_ROW_HEIGHT_CM

    rows, final_page_used_cm = _place_rows_for_page_fill(
        rows, usable_height_cm, initial_used_cm, font_size_pt=_TABLE_FONT_SIZE_PT, use_disappearance_date=use_disappearance_date,
    )

    if use_disappearance_date:
        data_rows = [
            [str(i + 1), row["ПОСАДА"], row["ЗВАННЯ"], row["ПІБ"], row["ДАТА_ЗНИКНЕННЯ"], row["ПЕРІОД"], row["ПІДСТАВА"]]
            for i, row in enumerate(rows)
        ]
    else:
        data_rows = [
            [str(i + 1), row["ПОСАДА"], row["ЗВАННЯ"], row["ПІБ"], row["ПЕРІОД"], str(row["ДНІ"]), row["ПІДСТАВА"]]
            for i, row in enumerate(rows)
        ]
    # Ширини колонок масштабуються так, щоб таблиця РІВНО заповнювала всю
    # ширину сторінки між полями (ні більше, ні менше) - widths_cm задає
    # лише СПІВВІДНОШЕННЯ між колонками, а не абсолютні см.
    width_scale = usable_width_cm / sum(widths_cm)
    scaled_widths_cm = [w * width_scale for w in widths_cm]
    create_table(
        doc, headers=headers, data_rows=data_rows, widths_cm=scaled_widths_cm, font_size=_TABLE_FONT_SIZE_PT,
        cell_alignment=WD_ALIGN_PARAGRAPH.CENTER, rotated_header_indices={5}, header_row_height_cm=HEADER_ROW_HEIGHT_CM,
        header_bold=False, body_first_line_bold=False, body_italic_first_cell=False, cell_margin_cm=0,
    )
    _center_basis_column(doc.tables[-1])
    add_blank_paragraphs(doc, 1)

    blank_paragraph_height_cm = _estimate_paragraph_height_cm("", _HEADING_FONT_SIZE_PT, usable_width_cm)
    return final_page_used_cm + blank_paragraph_height_cm


def _table_kwargs_for_point(template_key):
    """headers/widths_cm/use_disappearance_date для _add_category_section -
    SPETSKONTYNGENT_TABLE_* + use_disappearance_date=True для
    template_key == SPETSKONTYNGENT_VALUE (усі місця, де може з'явитись
    "100_СПЕЦКОНТИНГЕНТ" - звичайний пункт поточного місяця в головному циклі
    process_generate_report_for_get_money І ретроактивні додавання за
    попередні місяці, _add_extra_point_sections - обидва просто викликають
    цю функцію, замість дублювати ту саму перевірку двічі), {} (тобто
    типові TABLE_HEADERS/TABLE_WIDTHS_CM/use_disappearance_date=False,
    _add_category_section сама їх підставить) для решти пунктів."""
    if str(template_key) != SPETSKONTYNGENT_VALUE:
        return {}
    return {
        "headers": SPETSKONTYNGENT_TABLE_HEADERS,
        "widths_cm": SPETSKONTYNGENT_TABLE_WIDTHS_CM,
        "use_disappearance_date": True,
    }


def _add_category_section(
    doc, point_number, template_key, rows, month_int, start_day, end_day, carried_over_cm, section_map_key="SECTIONS",
    headers=TABLE_HEADERS, widths_cm=TABLE_WIDTHS_CM, use_disappearance_date=False,
):
    """Пункт звичайного рапорту - заголовок будується з шаблону
    STATIK[section_map_key][template_key] (текст із
    {start_date_day}/{end_date_day}/{manth}/{year} - для _NOT_PAID_POINT у тексті
    просто немає цих плейсхолдерів, .format() їх і не підставляє), решта - див.
    _add_heading_and_table_section.

    section_map_key - "SECTIONS" (звичайний рапорт, за замовчуванням) чи
    "COMMANDER_SECTIONS" (рапорт на командира/ТВО, generate_report_for_commander_money.py) -
    обидва в data_statik_report_for_get_additional_money, ключ - номер пункту як
    рядок (той самий template_key, що й у категоріях MONEY_REPORT_CATEGORIES/
    COMMANDER_MONEY_REPORT_CATEGORIES).

    headers/widths_cm - заголовки/ширини колонок таблиці (за замовчуванням -
    TABLE_HEADERS/TABLE_WIDTHS_CM); викликач передає NOT_PAID_TABLE_HEADERS для
    template_key == _NOT_PAID_POINT (інша остання колонка - "Підстава для не
    виплати"), чи SPETSKONTYNGENT_TABLE_HEADERS/_WIDTHS_CM (+
    use_disappearance_date=True) для template_key == SPETSKONTYNGENT_VALUE."""
    text = STATIK[section_map_key][template_key].format(
        start_date_day=f"{start_day:02d}", end_date_day=f"{end_day:02d}",
        manth=month_genitive_lower(month_int), year=YEAR,
    )
    heading_text = f"{point_number}. {text}"
    return _add_heading_and_table_section(
        doc, heading_text, rows, carried_over_cm, headers=headers, widths_cm=widths_cm, use_disappearance_date=use_disappearance_date,
    )


# Заповнюване вручну "№___ від___" у STATIK["CHANGES_INTRO"], коли для амендованого
# місяця немає відповідного запису в constants.MONEY_REPORT_CHANGES_ORDER_REFERENCES -
# та сама довжина порожнього місця, що й у тексті до появи цього механізму.
_BLANK_ORDER_REFERENCE = "№      від                "


def _add_changes_intro(doc, point_number, order_reference, carried_over_cm):
    """Заголовок пункту "Прошу внести зміни в наказ ... а саме:" - БЕЗ таблиці
    (сама таблиця йде під підпунктами N.1/N.2/... - _add_changes_subsection).
    order_reference - готовий текст (напр. "№1657 від 08.07.2026") з
    content.report_changes.build_changes_entries (перший "lines" запис
    MONEY_REPORT_CHANGES_ORDER_REFERENCES, чий період включає амендований
    місяць) - якщо для амендованого місяця немає відповідного запису (None),
    "№___ від___" лишається порожнім місцем для ручного заповнення, як і раніше."""
    heading_text = f"{point_number}. " + STATIK["CHANGES_INTRO"].format(
        order_reference=order_reference or _BLANK_ORDER_REFERENCE,
    )
    add_paragraph_with_style(doc, heading_text, first_line_indent=34, heading_level=1, font_size=_HEADING_FONT_SIZE_PT)

    usable_width_cm = _usable_text_width_cm(doc)
    heading_height_cm = _estimate_paragraph_height_cm(heading_text, _HEADING_FONT_SIZE_PT, usable_width_cm)
    return carried_over_cm + heading_height_cm


def _add_changes_subsection(doc, point_number, sub_index, label_key, appendix_number, rows, carried_over_cm):
    """Один підпункт "N.M Виключити пункти в Додатку .../Додаток ... доповнити
    наступними пунктами:" + таблиця - той самий заголовок+таблиця+пагінація, що
    й для звичайних пунктів (_add_heading_and_table_section), лише з дворівневим
    номером і appendix_number, підставленим у STATIK[label_key]."""
    heading_text = f"{point_number}.{sub_index} {STATIK[label_key].format(appendix_number=appendix_number)}"
    return _add_heading_and_table_section(doc, heading_text, rows, carried_over_cm)


def _add_changes_sections(doc, point_number, changes_entries, carried_over_cm):
    """Рендерить УСІ записи "Прошу внести зміни..." (build_changes_entries,
    content/report_changes.py) - один пункт (з наступним ЦІЛИМ номером) на КОЖЕН
    попередній місяць, що має файли в resources/changes/, найближчий до обраного
    місяця першим; підпункт N.1/N.2/... - по "Додати" на КОЖЕН пункт (30/100/50)
    цього місяця, що дійсно мав нові/змінені дні.

    "Виключити пункти в Додатку N" тут НЕ рендериться ВЗАГАЛІ, для ЖОДНОГО
    пункту - підтверджено користувачем ("я не маю взагалі бачити виключити"):
    report_changes.build_appendix_pairs_for_month/_delta_rows_for_point завжди
    повертає порожній exclude_rows (перший елемент кожної пари appendix_pairs) -
    ігнорується тут навмисно, а не тому що завжди порожній.

    RESTATE_POINTS (report_changes.py, зараз - лише 30) - рядки "add"-сторони
    пари РОЗДІЛЯЮТЬСЯ за "_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ" (report_changes.
    _delta_rows_for_point): рядки З прапорцем (в prev УЖЕ був запис цієї
    категорії - напр. частина 30-днів "переїхала" на 100, а решта лишилась
    30) - "CHANGES_RESTATE_LABEL" ("Викласти в новій редакції в Додатку N"),
    рядки БЕЗ прапорця (людина взагалі не фігурувала в prev - порожні
    клітинки, щойно долучилась) - звичайний "CHANGES_ADD_LABEL" ("Додаток N
    доповнити") - підтверджено користувачем: "перевидавати" ще неіснуючий
    запис не має сенсу. ОБИДВА варіанти можуть з'явитись для ОДНОГО й того
    самого пункту/місяця одночасно (для різних людей) - тоді рендеряться
    ДВОМА окремими підпунктами під ТИМ САМИМ номером Додатка, "Викласти в
    новій редакції" ПЕРШИМ, "доповнити" другим - підтверджено користувачем
    (реальний підсумок: пункт 100 - "доповнити Додаток 1"; пункт 30 - 2
    підпункти, 1-й "Викласти в новій редакції в Додатку 3" (дні перейшли на
    100, решта лишилась), 2-й "Додаток 3 доповнити" (людина не фігурувала в
    prev узагалі)). Пункти ПОЗА RESTATE_POINTS (напр. 100) - завжди ОДНИМ
    підпунктом, "CHANGES_ADD_LABEL".

    Запис, чий "appendix_pairs" порожній (лише "extra_points" - 70/170,
    _add_extra_point_sections нижче, рендеряться там, поза цим розділом), тут
    просто пропускається - інакше "N. Прошу внести зміни..." надрукувався б
    без жодного підпункту під собою.

    Повертає (наступний вільний номер пункту, carried_over_cm) - для продовження
    нумерації/заповнення сторінок після цього розділу (у звичайному рапорті після
    нього більше нічого немає, окрім підпису, але сигнатура симетрична решті
    рендер-функцій)."""
    for entry in changes_entries:
        if not entry["appendix_pairs"]:
            continue
        carried_over_cm = _add_changes_intro(doc, point_number, entry.get("order_reference"), carried_over_cm)
        sub_index = 1
        for point, appendix_number, _exclude_rows, add_rows in entry["appendix_pairs"]:
            if not add_rows:  # pragma: no cover - build_appendix_pairs_for_month
                # (report_changes.py) уже кладе в appendix_pairs ЛИШЕ пункти, де є
                # ХОЧ ОДНА зміна (свій докстрінг: "лише для пунктів, де є ХОЧ ОДНА
                # зміна") - add_rows тут порожнім бути не може; лишено як захисна
                # гілка на випадок майбутньої зміни джерела appendix_pairs.
                continue
            if point in RESTATE_POINTS:
                new_rows = [row for row in add_rows if not row.get("_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ")]
                restate_rows = [row for row in add_rows if row.get("_ВИКЛАСТИ_В_НОВІЙ_РЕДАКЦІЇ")]
            else:
                new_rows, restate_rows = add_rows, []
            if restate_rows:
                carried_over_cm = _add_changes_subsection(
                    doc, point_number, sub_index, "CHANGES_RESTATE_LABEL", appendix_number, restate_rows, carried_over_cm,
                )
                sub_index += 1
            if new_rows:
                carried_over_cm = _add_changes_subsection(
                    doc, point_number, sub_index, "CHANGES_ADD_LABEL", appendix_number, new_rows, carried_over_cm,
                )
                sub_index += 1
        point_number += 1
    return point_number, carried_over_cm


def _add_extra_point_sections(doc, point_number, changes_entries, carried_over_cm):
    """Рендерить "extra_points" (report_changes.build_changes_entries - об'єднання
    70/170 і BASIS_REQUIRED_POINTS, "100_СПЕЦКОНТИНГЕНТ"/"100_БПШП"/"100_ВПБП") -
    ВСІ "як у звичайному рапорті" (SECTIONS у resources/data.json, той самий
    _add_category_section, що й для звичайних пунктів - str(section["point"])
    працює однаково і для чисел (70/170), і для готових рядків-ключів
    ("100_СПЕЦКОНТИНГЕНТ" тощо)) - ПОЗА межами "Прошу внести зміни..."
    (_add_changes_sections вище), окремими, звичайними пунктами - підтверджено
    користувачем: "коли є цей статус тоді ми просто відображаємо секції які в
    resources/data.json SECTIONS 100_БПШП 100_ВПБП, 100_СПЕЦКОНТИНГЕНТ" (а не
    "Додаток N доповнити", як раніше). Один пункт (зі своїм цілим номером) на
    КОЖНУ (місяць, точка) пару, де є хоч один рядок - у порядку появи
    changes_entries (найближчий місяць першим), у межах місяця - спершу 70/170
    (report_changes._extra_point_rows_for_month), потім BASIS_REQUIRED_POINTS
    (report_changes._basis_appeared_extra_points_for_month) - порядок їхнього
    об'єднання в "extra_points" (content/report_changes.py).

    Повертає (наступний вільний номер пункту, carried_over_cm), як і решта
    рендер-функцій цього розділу."""
    for entry in changes_entries:
        for section in entry.get("extra_points", []):
            carried_over_cm = _add_category_section(
                doc, point_number, str(section["point"]), section["rows"], entry["month"],
                section["period_start_day"], section["period_end_day"], carried_over_cm,
                **_table_kwargs_for_point(section["point"]),
            )
            point_number += 1
    return point_number, carried_over_cm


def _add_missing_coverage_section(doc, changes_entries):
    """Рендерить "missing_coverage_warnings" (report_changes.build_changes_entries,
    content.document_content_search.collect_missing_document_coverage_warnings) -
    окремим списком у КІНЦІ документа (перед підписом) - підтверджено
    користувачем: "не потрібно щоб в термінал, а у файл word виводились
    відсутні" - раніше ці попередження друкувались у термінал (print_red), тепер
    - лише сюди, у сам документ. Збирає попередження з УСІХ записів
    changes_entries (усіх місяців) в ОДИН список - без свого номера пункту (не
    звичайний пункт рапорту, а примітка) - нічого не рендерить, якщо попереджень
    немає взагалі."""
    warnings = [w for entry in changes_entries for w in entry.get("missing_coverage_warnings", [])]
    if not warnings:
        return
    add_paragraph_with_style(doc, STATIK["MISSING_COVERAGE_HEADING"], bold=True)
    for warning in warnings:
        add_paragraph_with_style(doc, warning)
    add_blank_paragraphs(doc, 1)


def _add_footer(doc, narrator_row):
    add_paragraph_with_style(doc, "Рапорт подається на ____ аркушах.", first_line_indent=34)
    add_blank_paragraphs(doc, 1)

    # narrator_row - той самий, хто підписує рапорт (find_higher_commander на дату
    # підписання - сьогодні, рахує викликач ОДИН раз і використовує і для підпису
    # тут, і для _exclude_commander_and_tvo/build_changes_entries вище).
    role = "Тимчасово виконуючий обов'язки командира" if narrator_row.get("ТВО", False) else "Командир"

    add_signature_block(doc, role, narrator_row)


def _log_personnel_day_value_periods(rows_with_data, date_columns, log_file_path):
    """Діагностичне логування (фіолетовим у консоль + текстовий файл): для кожної
    людини - ПІБ і період(и) для КОЖНОГО значення, що трапилось у її комірках дат
    цього місяця (30, 100, коди відпустки/відрядження ВП/ВД/ВЛК/ПРВД, текст-назва
    категорії тощо) - щоб одразу бачити по днях, чи все розпізналось так, як
    очікується, не заглиблюючись у сам рапорт."""
    lines = []
    for row in rows_with_data:
        breakdown = build_day_value_periods(row, date_columns)
        if not breakdown:
            continue

        parts = [f"{value}: {period_text} ({days_count} дн.)" for value, (period_text, days_count) in breakdown.items()]
        lines.append(f"{row.get('ПІБ', '')} - " + "; ".join(parts))

    for line in lines:
        print_purple(line)

    with open(log_file_path, "w", encoding="utf-8") as log_file:
        log_file.write("\n".join(lines))


def _flatten_changes_rows_for_logging(changes_entries):
    """Перетворює changes_entries (content.report_changes.build_changes_entries) у
    той самий формат (section_key, rows), що й _build_categories повертає для
    звичайних пунктів - щоб _log_missing_legal_basis попереджала про порожню
    "Підстава для виплати" так само й для рядків "зміни за попередні місяці"
    (підстава минулого місяця могла спиратись на конфігурацію БР/grounds, якої
    вже немає в поточному constants.py - на відміну від звичайних пунктів, це
    цілком реальний сценарій для рядка, що описує МИНУЛИЙ місяць)."""
    flattened = []
    for entry in changes_entries:
        for point, _appendix_number, exclude_rows, add_rows in entry["appendix_pairs"]:
            if exclude_rows:
                flattened.append((f"{point}_CHANGES_EXCLUDE", exclude_rows))
            if add_rows:
                flattened.append((f"{point}_CHANGES_ADD", add_rows))
        for section in entry.get("extra_points", []):
            flattened.append((f"{section['point']}_CHANGES_EXTRA", section["rows"]))
    return flattened


def _log_missing_legal_basis(categories):
    """Попереджає (червоним у консоль) про кожну людину, чия "Підстава для виплати"
    в рапорті вийшла порожньою (напр. категорія на кшталт "70_РТГр"/"170_РТГр" ще не
    має заповнених "grounds", або весь "general" для неї виключено) - щоб одразу було
    видно, кому саме бракує підстави, не гортаючи весь документ.

    Категорії з "required_basis": False (напр. 10к - "виконання обов'язків військової
    служби", без прив'язки до конкретного БР/наказу) свідомо пропускаються - для них
    порожня підстава ОЧІКУВАНА, а не помилка (див. resolve_category_config).

    Категорії з "use_brs_from_selected_folder": True (напр. МЕДИК пункту 100 - підстава
    береться з щоденних БР, зчитаних з обраної папки з документами) при порожній
    підставі отримують ОКРЕМЕ, зрозуміліше повідомлення - не "кому саме бракує
    підстави", а "папку з документами не було зчитано" (відповідь "Ні" на питання
    "Зчитати обрану папку з документами (для номерів БАТ / БЗ)?" чи в ній просто немає
    потрібних дат) - одне повідомлення на (пункт, категорія), а не по одному на кожну людину.

    "NOT_PAID" (і його секції "Прошу внести зміни..." - NOT_PAID_CHANGES_EXCLUDE/
    NOT_PAID_CHANGES_ADD) - ОКРЕМЕ формулювання ("Немає підстави ДЛЯ НЕ ВИПЛАТИ") -
    підтверджено користувачем: цей пункт описує причину НЕВИПЛАТИ (СЗЧ/задув.), а
    не підставу для виплати, тож типове "Немає підстави для виплати" тут мало б
    протилежний до реального сенс."""
    reported_missing_from_folder = set()
    for section_key, rows in categories:
        point = section_key.split("_", 1)[0]
        is_not_paid = section_key == _NOT_PAID_POINT or section_key.startswith(f"{_NOT_PAID_POINT}_CHANGES")
        for row in rows:
            if row["ПІДСТАВА"] or not row.get("_ПІДСТАВА_ОБОВ'ЯЗКОВА", True):
                continue
            if row.get("_ВИКОРИСТОВУЄ_ОБРАНУ_ПАПКУ"):
                key = (point, row.get("_КАТЕГОРІЯ"))
                if key not in reported_missing_from_folder:
                    reported_missing_from_folder.add(key)
                    print_red(f"⚠ Для {point} статус {row.get('_КАТЕГОРІЯ')} завдання не вдалось зчитати.")
                continue
            if is_not_paid:
                print_red(f"⚠ Немає підстави для не виплати: {row['ПІБ']} ({section_key}, період {row['ПЕРІОД']})")
            else:
                print_red(f"⚠ Немає підстави для виплати: {row['ПІБ']} ({section_key}, період {row['ПЕРІОД']})")


def _log_missing_posada(categories):
    """Попереджає (червоним у консоль) про кожну людину, чия "Посада" в рапорті
    вийшла порожньою - напр. у ОБЛІК.xlsx для неї не заповнена ця колонка, або
    (для секцій "Прошу внести зміни...", section_key закінчується на
    "_CHANGES_EXCLUDE"/"_CHANGES_ADD") жоден з файлів prev/actual пари цього
    місяця не мав її ПОСАДА, тож content.report_changes._backfill_missing_labels
    просто не мала звідки її взяти - в такому разі повідомлення уточнює, що це
    саме секція змін до наказу."""
    for section_key, rows in categories:
        is_changes_section = section_key.endswith("_CHANGES_EXCLUDE") or section_key.endswith("_CHANGES_ADD")
        for row in rows:
            if row.get("ПОСАДА"):
                continue
            suffix = " в секції зміни до наказу" if is_changes_section else ""
            print_red(f"для {row['ПІБ']} немає посади{suffix}")


def process_generate_report_for_get_money(rows_with_data, rows_with_tvo_data, column_names_personel, changes_file_pairs=()):
    month_int = int(MONTH)
    date_columns = get_date_columns(column_names_personel)
    if not date_columns:
        print_green("Немає жодної дати за поточний місяць у ОБЛІК.xlsx - рапорт на додаткову винагороду не сформовано.")
        return None

    period_str = f"{date_columns[0].strftime('%d.%m.%Y')}-{date_columns[-1].strftime('%d.%m.%Y')}"

    # ОБЛІК.xlsx більше не містить колонки СТАТУС - базового статусу з файлу
    # більше немає (тільки per-day текстове перевизначення категорії), тож сюди
    # завжди передається порожній лookup.
    status_lookup = {}

    # Оригінальний, НЕвідфільтрований список (перш ніж рядок нижче прибере з
    # нього командира/ТВО для тіла рапорту) - потрібен find_higher_commander
    # нижче (як запасний варіант, коли на сьогодні немає активного ТВО) -
    # переданий уже ВІДФІЛЬТРОВАНИЙ rows_with_data ніколи не містив би його, і
    # підпис рапорту завжди лишався б порожнім, навіть коли командир реально є
    # в ОБЛІК.xlsx.
    all_rows_with_data = rows_with_data

    # Той самий, хто підписує ЦЕЙ рапорт (find_higher_commander на дату
    # підписання - сьогодні) - визначає ще й КОГО саме прибрати з тіла рапорту
    # нижче (_exclude_commander_and_tvo) і чиї зміни за попередні місяці
    # лишити в ГОЛОВНОМУ рапорті (build_changes_entries) - підтверджено
    # користувачем: якщо підписує штатний командир (немає активного ТВО на
    # сьогодні) - жоден ТВО не прибирається з ГОЛОВНОГО рапорту (навіть якщо
    # він реально був ТВО частину місяця) - він, як і решта підлеглих, і далі
    # фігурує тут звичайним рядком, а не залишається невидимим у жодному з
    # двох рапортів одразу.
    narrator_row = find_higher_commander(rows_with_tvo_data, all_rows_with_data, HIGHER_COMMANDER_TITLE, datetime.now())

    rows_with_data = _exclude_commander_and_tvo(rows_with_data, rows_with_tvo_data, date_columns, narrator_row)

    # В/сл, у кого ХОЧ ЗА ОДИН день місяця в комірці ОБЛІК.xlsx стоїть "ПРВД"
    # (переведено в інший підрозділ/частину) - не потрапляють у рапорт взагалі,
    # незалежно від значень решти їхніх комірок дат. "СЗЧ" сюди не входить -
    # звичайна категорія пункту "NOT_PAID" (_build_categories нижче).
    rows_with_data = _exclude_by_excluded_day_value(rows_with_data, date_columns)

    # В/сл, чий СЗЧ ПОЧАВСЯ цього місяця - НЕ отримує оплату за жоден звичайний
    # пункт (30/70/100/170/10), АЛЕ й далі з'являється в "NOT_PAID" за свої дні
    # СЗЧ (підтверджено користувачем: див. docstring
    # _blank_non_szch_days_for_szch_started_this_month) - ПІСЛЯ ПРВД вище, щоб
    # людина з ОБОМА цими статусами лишалась виключеною ПРВД цілком (реальний
    # приклад: повернулась з СЗЧ, а тоді була переведена).
    rows_with_data = _blank_non_szch_days_for_szch_started_this_month(rows_with_data, date_columns)

    # Дні-статуси з NOT_PAID_IGNORED_WHEN_CONFIRMED_BY (constants.py - напр.
    # "СЗЧ", підтверджений "РОЗП" - людину офіційно оголошено в розшуку) -
    # ІГНОРУЮТЬСЯ (не потрапляють у "NOT_PAID" узагалі, ні рядком, ні
    # попередженням "Немає підстави для не виплати") - підтверджено
    # користувачем на реальному випадку.
    rows_with_data = _blank_not_paid_days_confirmed_elsewhere(rows_with_data, date_columns)

    log_file_name = f"Лог статусів {SHORT_UNIT_BATTALION.upper()} ДВ {month_nominative_upper(month_int).lower()} від {period_str}.txt"
    _log_personnel_day_value_periods(rows_with_data, date_columns, os.path.join(OUTPUT_DIR, log_file_name))

    categories = _build_categories(rows_with_data, date_columns, status_lookup)
    changes_entries = build_changes_entries(rows_with_tvo_data, changes_file_pairs=changes_file_pairs, narrator_row=narrator_row)
    combined_for_logging = categories + _flatten_changes_rows_for_logging(changes_entries)
    _log_missing_legal_basis(combined_for_logging)
    _log_missing_posada(combined_for_logging)

    # _NOT_PAID_POINT рендериться ОДРАЗУ ПІСЛЯ звичайних пунктів ОБРАНОГО
    # (серпневого) місяця, ПЕРЕД "Прошу внести зміни..." - підтверджено
    # користувачем: обраний місяць - У ПРІОРИТЕТІ (усі його пункти, включно з
    # NOT_PAID, ідуть ПЕРШИМИ), а вже ПОТІМ - зміни за ПОПЕРЕДНІ місяці (за
    # спаданням, найближчий перший) - раніше NOT_PAID рендерився В САМОМУ
    # КІНЦІ документа (після ВСІХ місяців "Прошу внести зміни..."), через що
    # виглядав як частина ЦИХ попередніх місяців, хоча стосується САМЕ
    # обраного (серпневого) місяця. Витягуємо його із загального списку тут
    # (для логування вище він і далі враховувався РАЗОМ з рештою categories).
    not_paid_entry = next((item for item in categories if item[0] == _NOT_PAID_POINT), None)
    categories = [item for item in categories if item[0] != _NOT_PAID_POINT]

    doc = Document()
    set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.3903))
    add_page_number_header(doc)

    _add_intro(doc, month_int)

    carried_over_cm = 0.0
    point_number = 1
    for template_key, rows in categories:
        carried_over_cm = _add_category_section(
            doc, point_number, template_key, rows, month_int, date_columns[0].day, date_columns[-1].day, carried_over_cm,
            **_table_kwargs_for_point(template_key),
        )
        point_number += 1

    if not_paid_entry:
        _, not_paid_rows = not_paid_entry
        carried_over_cm = _add_category_section(
            doc, point_number, _NOT_PAID_POINT, not_paid_rows, month_int, date_columns[0].day, date_columns[-1].day, carried_over_cm,
            headers=NOT_PAID_TABLE_HEADERS,
        )
        point_number += 1

    point_number, carried_over_cm = _add_changes_sections(doc, point_number, changes_entries, carried_over_cm)
    point_number, carried_over_cm = _add_extra_point_sections(doc, point_number, changes_entries, carried_over_cm)

    _add_missing_coverage_section(doc, changes_entries)
    _add_footer(doc, narrator_row)

    file_name = f"Рапорт {SHORT_UNIT_BATTALION.upper()} ДВ {month_nominative_upper(month_int).lower()} від {period_str}.docx"
    file_path = os.path.join(OUTPUT_DIR, file_name)
    save_docx_safely(doc, file_path)
    print_green(f"- {file_name}")
    return file_path


def process_generate_changes_report(rows_with_data, rows_with_tvo_data, changes_file_pairs):
    """RUN_MODE_CHANGES_ONLY - окремий документ, що містить ЛИШЕ розділи "Прошу
    внести зміни..." за попередні місяці плюс, за наявності, точки 70/170 "як у
    звичайному рапорті" (_add_extra_point_sections - поза межами "Прошу внести
    зміни...") - без решти звичайних пунктів рапорту - для подання самої
    поправки окремо від звичайного щомісячного рапорту. Використовує ту саму
    build_changes_entries/_add_changes_sections(+_add_extra_point_sections), що
    й вбудований розділ у process_generate_report_for_get_money - вміст завжди
    однаковий, лише документ, у який він потрапляє, інший."""
    narrator_row = find_higher_commander(rows_with_tvo_data, rows_with_data, HIGHER_COMMANDER_TITLE, datetime.now())

    changes_entries = build_changes_entries(rows_with_tvo_data, changes_file_pairs=changes_file_pairs, narrator_row=narrator_row)
    if not changes_entries:
        print_green("Немає підтверджених змін за попередні місяці - рапорт на поправки в наказі не сформовано.")
        return None

    changes_rows_for_logging = _flatten_changes_rows_for_logging(changes_entries)
    _log_missing_legal_basis(changes_rows_for_logging)
    _log_missing_posada(changes_rows_for_logging)

    doc = Document()
    set_margins(doc, Inches(0.7875), Inches(0.7875), Inches(1.18125), Inches(0.3903))
    add_page_number_header(doc)

    _add_changes_only_intro(doc)
    point_number, carried_over_cm = _add_changes_sections(doc, 1, changes_entries, 0.0)
    _add_extra_point_sections(doc, point_number, changes_entries, carried_over_cm)
    _add_missing_coverage_section(doc, changes_entries)
    _add_footer(doc, narrator_row)

    file_name = f"Поправки в наказ {SHORT_UNIT_BATTALION.upper()} за {month_nominative_upper(int(MONTH)).lower()} {YEAR}.docx"
    file_path = os.path.join(OUTPUT_DIR, file_name)
    save_docx_safely(doc, file_path)
    print_green(f"- {file_name}")
    return file_path
