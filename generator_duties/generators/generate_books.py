"""Пункт меню "Згенерувати документи": дві книги (функціональні й посадові
обов'язки) з одного book_builder - відрізняються лише аркуш таблиці й назви."""
import math
import os
from dataclasses import dataclass, field

from constants import BLANK_ROWS_MIN, BLANK_ROWS_SHARE, BOOKS
from content.duties_table import SheetDuties
from content.positions import find_commander
from generators.book_builder import BookInput, Signer, build_book
from generators.word_finalize import FinalizeError, finalize_with_word, finalize_without_word
from utils.docx_utils import save_docx_safely
from utils.names import signature_name


@dataclass
class BookResult:
    path: str
    sheet: str
    review: object
    unknown_sections: list = field(default_factory=list)
    stale: list = field(default_factory=list)
    missing_sets: list = field(default_factory=list)
    blank_rows: int = 0
    pages: int = 0
    sheets: int = 0
    finalize_error: str = ""


def make_signer(positions, data):
    """Комбат для підписів: зі штату (посада "Командир батальйону"), будь-яке
    непорожнє поле data.json → signer перекриває знайдене."""
    unit = data.get("unit") or {}
    override = data.get("signer") or {}
    slot = find_commander(positions)
    signer = Signer()
    if slot is not None:
        position = slot.position
        unit_genitive = unit.get("MILITARY_UNIT_GENITIVE", "")
        if unit_genitive and unit_genitive.lower() not in position.lower():
            position = f"{position} {unit_genitive}"
        signer = Signer(position, slot.rank, signature_name(slot.pib))
    return Signer(
        override.get("position") or signer.position,
        override.get("rank") or signer.rank,
        override.get("name") or signer.name,
    )


def blank_rows(data, slots_count):
    """Порожні рядки листа ознайомлення: data.json → blank_rows, інакше
    BLANK_ROWS_SHARE від кількості посад (вгору до десятків), не менше BLANK_ROWS_MIN."""
    configured = data.get("blank_rows")
    if isinstance(configured, int) and not isinstance(configured, bool) and configured >= 0:
        return configured
    return max(BLANK_ROWS_MIN, math.ceil(slots_count * BLANK_ROWS_SHARE / 10) * 10)


def generate_books(positions, table, sources, data, year, output_dir, finalize=finalize_with_word):
    signer = make_signer(positions, data)
    position_keys = {position.key for position in positions}
    extract = sources.sample.extract if sources.sample else []
    blank = blank_rows(data, sum(len(position.slots) for position in positions))
    results = []
    for spec in BOOKS:
        duties = SheetDuties(table, spec["sheet"])
        doc, review = build_book(BookInput(spec, year, data.get("unit") or {}, signer, positions, duties, extract,
                                           sources.regulation, blank))
        path = os.path.join(output_dir, spec["file_name"].format(year=year))
        save_docx_safely(doc, path)
        result = BookResult(
            path, spec["sheet"], review, duties.unknown_sections,
            [row.title for row in table.positions if row.key not in position_keys],
            [(position.title, name) for position in positions if (name := duties.missing_set_name(position.key))],
            blank,
        )
        try:
            result.pages, result.sheets = finalize(path, bool(data.get("duplex")))
        except FinalizeError as e:
            result.finalize_error = str(e)
            finalize_without_word(path)
        results.append(result)
    return results, signer


def review_lines(results, signer, warnings=()):
    """Усе, що варто перевірити, - для консолі й Потребує_ручної_перевірки.txt."""
    lines = [f"- {warning}" for warning in warnings]
    if not signer.name:
        lines.append("- Комбата не знайдено в штаті (посада «Командир батальйону») - підпис з прочерками; "
                     "можна вказати в data.json → signer.")
    for result in results:
        review = result.review
        book = f"[{result.sheet}]"
        with_wrong_set = {title for title, _ in result.missing_sets}
        missing = [title for title in review.missing if title not in with_wrong_set]
        if missing:
            lines.append(f"- {book} Посади без набору обов'язків ({len(missing)}) - запустіть «Оновити таблицю»: "
                         + "; ".join(missing))
        for title, name in result.missing_sets:
            lines.append(f"- {book} Посада «{title}» посилається на набір «{name}», якого немає на аркуші "
                         f"«{result.sheet}» - виправте назву на аркуші «Посади».")
        if review.to_fill:
            lines.append(f"- {book} Є «ЗАПОВНИТИ» ({len(review.to_fill)} посад, у книзі - жовтим): "
                         + "; ".join(review.to_fill))
        if review.without_vos:
            lines.append(f"- {book} Без ВОС у штаті ({len(review.without_vos)}): " + "; ".join(review.without_vos))
        if result.stale:
            lines.append(f"- {book} Посади з таблиці, яких немає в штаті (у книгу не потрапили): " + "; ".join(result.stale))
        for _, set_name, section in result.unknown_sections:
            what = f"невідомий розділ «{section}»" if section else "пункт без розділу"
            lines.append(f"- {book} Набір «{set_name}»: {what} - пропущено.")
        if result.finalize_error:
            lines.append(f"- {book} {result.finalize_error}: відкрийте {os.path.basename(result.path)}, оновіть ЗМІСТ "
                         f"(Ctrl+A, F9) і впишіть кількість аркушів у напис про зшивання.")
    return lines
