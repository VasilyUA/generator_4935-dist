"""Спільний файл повідомлення командирам (content/commander_message.py)."""

from datetime import date

import content.commander_message as cm


def test_both_reports_go_into_one_file_in_the_users_order(tmp_path):
    """Рішення користувача: ОДИН файл на обидва звіти: спершу
    звірка з БЧС, потім з відомістю (незалежно від порядку генерації); той
    самий підрозділ з обох звітів - одна група, та сама людина - один рядок."""
    path = str(tmp_path / "Повідомлення_командирам.txt")
    payments = [
        ("1 абв", "ПЕРШИЙ Перший Перший", cm.date_issues({date(2026, 10, 26): "Навч", date(2026, 10, 27): "Навч"})),
        ("вг", "ДРУГИЙ Другий Другий", ["Звання: сержант"]),
    ]
    schedule = [
        ("де", "ТРЕТІЙ Третій Третій", cm.date_issues({date(2026, 10, 26): 170})),
        ("вг", "ДРУГИЙ Другий Другий", cm.date_issues({date(2026, 10, 27): 70})),
    ]

    assert cm.update_message_file(path, cm.SCHEDULE_SOURCE, schedule) == path
    assert cm.update_message_file(path, cm.PAYMENTS_SOURCE, payments) == path

    assert (tmp_path / "Повідомлення_командирам.txt").read_text(encoding="utf-8") == (
        "Командири підрозділів нижчевказаних виправте табель в своїх файлах БЧС для військовослужбовців:\n"
        "1 абв ПЕРШИЙ Перший Перший - 26.10.2026-27.10.2026 статус має бути (Навч)\n"
        "\n"
        "вг ДРУГИЙ Другий Другий - Звання: сержант\n"
        "\t27.10.2026 статус має бути (70)\n"
        "\n"
        "де ТРЕТІЙ Третій Третій - 26.10.2026 статус має бути (170)\n"
    )


def test_nothing_to_say_writes_no_file(tmp_path):
    path = str(tmp_path / "out" / "message.txt")
    assert cm.update_message_file(path, cm.PAYMENTS_SOURCE, [("вг", "ПЕРШИЙ Перший Перший", [])]) is None
    assert not (tmp_path / "out").exists()
