"""Повідомлення командирам підрозділів: що виправити у своїх файлах БЧС
(формат - за зразком користувача). Один файл на обидва звіти:
error_mis_statuses.xlsx (content/payment_mismatch_checker.py) і
error_mis_statuses_Відомість.xlsx (content/rop_vop_statement.py).

Спершу заголовок, далі групи за підрозділом (у порядку записів, з порожнім
рядком між групами). У групі - "<підрозділ> <ПІБ> - <що виправити>", а кожне
наступне виправлення - окремим рядком з табуляцією."""

import os
from datetime import timedelta

MESSAGE_HEADER = "Командири підрозділів нижчевказаних виправте табель в своїх файлах БЧС для військовослужбовців:"
NOT_IN_ROSTER_ISSUE = "ПІБ не знайдено в ОБЛІК - перевірте написання"
_NO_SUBDIVISION = "—"

# Джерела записів у тому порядку, в якому вони йдуть у файлі (за зразком
# користувача - спершу звірка з БЧС, потім з відомістю), хоч би в якому
# порядку працювали генератори.
PAYMENTS_SOURCE = "error_mis_statuses"
SCHEDULE_SOURCE = "error_mis_statuses_Відомість"
_SOURCES = (PAYMENTS_SOURCE, SCHEDULE_SOURCE)

# {абсолютний шлях файлу повідомлення: {джерело: записи}}. Звіти пишуть різні
# генератори ("Згенерувати все" - один за одним в одному запуску), тож файл
# перебудовується з усіх джерел, записаних у нього за цей запуск. Іншого
# вмісту там бути не може: index.py очищає output/ на початку кожного запуску.
_entries_by_path = {}


def date_issues(values_by_date):
    """{дата: як має бути} -> ["дд.мм.рррр-дд.мм.рррр статус має бути (...)"].
    Дні поспіль з тим самим значенням дають один проміжок."""
    ranges = []
    for date_value in sorted(values_by_date):
        shown = values_by_date[date_value]
        last = ranges[-1] if ranges else None
        if last and last[2] == shown and last[1] + timedelta(days=1) == date_value:
            last[1] = date_value
        else:
            ranges.append([date_value, date_value, shown])
    return [
        f"{start:%d.%m.%Y}{'' if start == end else f'-{end:%d.%m.%Y}'} статус має бути ({shown})"
        for start, end, shown in ranges
    ]


def compose_message(entries):
    """[(підрозділ, ПІБ, [що виправити])] -> текст повідомлення ("" - нема
    про що писати). Той самий підрозділ, записаний по-різному ("1 X" / "1 x"),
    - одна група; та сама людина в кількох записах - один рядок без повторів."""
    groups = {}  # нормалізований підрозділ -> (як показати, {ПІБ: [що виправити]})
    for subdivision, pib, issues in entries:
        if not issues:
            continue
        shown = str(subdivision or _NO_SUBDIVISION).strip()
        _shown, people = groups.setdefault(" ".join(shown.lower().split()), (shown, {}))
        person_issues = people.setdefault(pib, [])
        person_issues.extend(issue for issue in issues if issue not in person_issues)
    if not groups:
        return ""
    blocks = [
        "\n".join(
            f"{shown} {pib} - {issues[0]}" + "".join(f"\n\t{issue}" for issue in issues[1:])
            for pib, issues in people.items()
        )
        for shown, people in groups.values()
    ]
    return MESSAGE_HEADER + "\n" + "\n\n".join(blocks) + "\n"


def update_message_file(output_path, source, entries):
    """Записує записи source (одне з _SOURCES) для файлу output_path і
    перебудовує файл з усіх його джерел у порядку _SOURCES (той самий
    підрозділ чи людина з обох звітів - одна група, один рядок). Повертає
    output_path або None, якщо писати нема про що - тоді файл не створюється,
    як і самі звіти розбіжностей."""
    by_source = _entries_by_path.setdefault(os.path.abspath(output_path), {})
    by_source[source] = list(entries)
    text = compose_message(entry for name in _SOURCES for entry in by_source.get(name, []))
    if not text:
        return None
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as message_file:
        message_file.write(text)
    return output_path
