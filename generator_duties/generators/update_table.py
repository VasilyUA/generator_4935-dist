"""Пункт меню "Оновити таблицю обов'язків зі штату": на аркуші "Посади" -
рівно ті посади, що є в штаті, у порядку штату; кожна посада посилається на
набір обов'язків для кожної книги.

- нова посада → набір з джерел (prefill); такий самий набір, що вже є, не
  дублюється - посада посилається на нього;
- посада, чий набір ще лише "ЗАПОВНИТИ" (і більше нікому не потрібен) або
  вказаний набір не існує → заповнюється заново (могло з'явитися джерело);
- набори, на які посилаються посади, і "Для всіх" програма не змінює;
- посади, яких немає в штаті, видаляються лише за згодою (remove_stale), а
  разом з ними - набори, які після цього нікому не потрібні;
- таблиця старого формату перетворюється (migrated)."""
import os
import shutil
from dataclasses import dataclass, field

from content.duties_table import ALL_KEY, SHEET_NAMES, DutiesTable, PositionRow, read_table, write_table
from content.positions import compare_key
from content.prefill import common_set, label_names, prefill_set


@dataclass
class UpdateResult:
    path: str
    created: bool = False
    migrated: bool = False
    added: list = field(default_factory=list)
    refilled: list = field(default_factory=list)
    removed: list = field(default_factory=list)
    stale_kept: list = field(default_factory=list)
    to_fill: dict = field(default_factory=dict)
    set_counts: dict = field(default_factory=dict)


def stale_positions(table, positions):
    """Посади з аркуша "Посади", яких немає в штаті - назви як у таблиці."""
    keys = {position.key for position in positions}
    return [row.title for row in table.positions if row.key not in keys]


def _used_keys(table, sheet):
    keys = []
    for row in table.positions:
        key = compare_key(row.sets.get(sheet, ""))
        if key and key not in keys:
            keys.append(key)
    return keys


def _refresh_sets(table, row, position, sources, result):
    for sheet in SHEET_NAMES:
        current = table.set_for(sheet, row.sets.get(sheet, ""))
        if current is not None and not (current.is_to_fill and table.users(sheet, current) == [row]):
            continue
        fresh = prefill_set(position, sheet, sources)
        if current is not None and fresh.is_to_fill:
            continue
        row.sets[sheet] = table.add_set(sheet, fresh).name
        if current is not None and position.title not in result.refilled:
            result.refilled.append(position.title)


def _order_sets(table, used_before):
    """"Для всіх" → набори в порядку першого використання посадами → набори,
    якими ніхто ніколи не користувався (створені вручну). Набори, що були
    потрібні до оновлення, а тепер - ні, прибираються."""
    for sheet in SHEET_NAMES:
        sets = table.sets[sheet]
        ordered = {ALL_KEY: sets[ALL_KEY]} if ALL_KEY in sets else {}
        for key in _used_keys(table, sheet):
            if key in sets:
                ordered.setdefault(key, sets[key])
        for key, duty_set in sets.items():
            if key not in ordered and key not in used_before[sheet]:
                ordered[key] = duty_set
        table.sets[sheet] = ordered


def update_table(path, backup_path, positions, sources, remove_stale=False):
    """Синхронізує таблицю path зі штатом (positions); стара таблиця перед
    записом копіюється в backup_path."""
    created = not os.path.isfile(path)
    table = DutiesTable() if created else read_table(path, label_names(sources))
    result = UpdateResult(path, created=created, migrated=table.legacy)
    used_before = {sheet: set(_used_keys(table, sheet)) for sheet in SHEET_NAMES}
    for sheet in SHEET_NAMES:
        common = table.sets[sheet].get(ALL_KEY)
        if common is None or not common.items():
            table.sets[sheet][ALL_KEY] = common_set(sheet, sources)

    existing = {row.key: row for row in table.positions}
    staff_rows = []
    for position in positions:
        row = existing.pop(position.key, None)
        if row is None:
            row = PositionRow(position.title)
            table.positions.append(row)
            result.added.append(position.title)
        row.title = position.title
        _refresh_sets(table, row, position, sources, result)
        row.total = len(position.slots)
        row.vacant = len(position.slots) - position.occupied
        staff_rows.append(row)
    stale_rows = list(existing.values())
    (result.removed if remove_stale else result.stale_kept).extend(row.title for row in stale_rows)
    table.positions = staff_rows + ([] if remove_stale else stale_rows)
    _order_sets(table, used_before)

    for sheet in SHEET_NAMES:
        result.to_fill[sheet] = [row.title for row in staff_rows
                                 if (duty_set := table.set_for(sheet, row.sets.get(sheet, ""))) is None
                                 or duty_set.has_to_fill]
        result.set_counts[sheet] = len([key for key in table.sets[sheet] if key != ALL_KEY])
    if not created:
        shutil.copy2(path, backup_path)
    write_table(path, table)
    return result
