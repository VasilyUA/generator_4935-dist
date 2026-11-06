from datetime import date

from openpyxl import load_workbook

import content.category_100_report as report_module
from content.oblik_timesheet import normalize_name

_FIRST = "ПЕРШИЙ Перший Перший"
_SECOND = "ДРУГИЙ Другий Другий"
_LONG_TASK = "завдання " * 40


def _violation(pib_raw, pidrozdil="1 абв", days=31, color=None):
    return {
        "reason": "текст для файлу", "pib_raw": pib_raw, "pidrozdil": pidrozdil,
        "start": date(2026, 10, 1), "end": date(2026, 10, days), "days": days, "limit": 28, "color": color,
    }


def _write(tmp_path, violations, tasks):
    path = str(tmp_path / "Перевірка_категорії_100.xlsx")
    assert report_module.write_category_100_report(violations, tasks, path) == path
    return load_workbook(path).worksheets[0]


def test_first_columns_are_subdivision_name_and_both_tasks(tmp_path):
    """Рішення користувача: 1 колонка - підрозділ, 2 - ПІБ, далі завдання з
    аркуша БЧС: ФАКТИЧНО і ПО ДОКУМЕНТАХ НА ВИПЛАТИ."""
    ws = _write(tmp_path, [_violation(_FIRST)], {
        normalize_name(_FIRST): {"actual": "Чергування", "documents": "Оборона позиції"},
    })

    assert [ws.cell(row=1, column=col).value for col in range(1, 5)] == [
        "Підрозділ", "ПІБ",
        "Завдання яке виконує військовослужбовець (ФАКТИЧНО)",
        "Бойове завдання яке виконує військовослужбовець (ПО ДОКУМЕНТАХ НА ВИПЛАТИ)",
    ]
    assert [ws.cell(row=2, column=col).value for col in range(1, 5)] == ["1 абв", _FIRST, "Чергування", "Оборона позиції"]


def test_the_violation_itself_stays_in_the_table(tmp_path):
    """Те, що було в текстовому файлі: період, скільки днів і скільки дозволено."""
    ws = _write(tmp_path, [_violation(_FIRST)], {})

    assert [ws.cell(row=1, column=col).value for col in range(5, 8)] == ["Період", "Днів поспіль", "Дозволено"]
    assert [ws.cell(row=2, column=col).value for col in range(5, 8)] == ["01.10.2026 - 31.10.2026", 31, 28]


def test_person_without_tasks_gets_a_placeholder_not_an_empty_cell(tmp_path):
    """Реальний випадок: у файлі управління немає аркуша БЧС."""
    ws = _write(tmp_path, [_violation(_FIRST)], {})

    assert [ws.cell(row=2, column=col).value for col in (3, 4)] == ["—", "—"]


def test_one_empty_task_does_not_hide_the_other(tmp_path):
    ws = _write(tmp_path, [_violation(_FIRST)], {normalize_name(_FIRST): {"actual": "Чергування", "documents": None}})

    assert [ws.cell(row=2, column=col).value for col in (3, 4)] == ["Чергування", "—"]


def test_person_without_a_subdivision_gets_a_placeholder(tmp_path):
    ws = _write(tmp_path, [{**_violation(_FIRST), "pidrozdil": None}], {})

    assert ws.cell(row=2, column=1).value == "—"


def test_every_violation_has_its_own_row_in_the_given_order(tmp_path):
    ws = _write(tmp_path, [_violation(_SECOND), _violation(_FIRST)], {})

    assert [ws.cell(row=row, column=2).value for row in (2, 3)] == [_SECOND, _FIRST]


def test_note_about_exceptions_stays_under_the_table(tmp_path):
    """Текстовий файл підказував, що робити, якщо так і має бути."""
    ws = _write(tmp_path, [_violation(_FIRST)], {})

    assert "кількість днів в EXCLUDE_NAMES" in str(ws.cell(row=4, column=1).value)


def test_person_without_100_days_gets_a_placeholder_period(tmp_path):
    ws = _write(tmp_path, [{**_violation(_FIRST, color="FFFFFF00"), "start": None, "end": None, "days": 0}], {})

    assert [ws.cell(row=2, column=col).value for col in range(5, 8)] == ["—", 0, 28]


def test_note_explains_why_a_colored_person_is_in_the_table(tmp_path):
    ws = _write(tmp_path, [_violation(_FIRST)], {})

    assert "навіть без порушення" in str(ws.cell(row=4, column=1).value)


def test_name_cell_gets_the_color_of_the_person(tmp_path):
    ws = _write(tmp_path, [_violation(_FIRST, color="FFFFFF00")], {})

    assert ws.cell(row=2, column=2).fill.fill_type == "solid"
    assert ws.cell(row=2, column=2).fill.start_color.rgb == "FFFFFF00"


def test_person_without_color_stays_without_a_fill(tmp_path):
    ws = _write(tmp_path, [_violation(_FIRST, color=None)], {})

    assert ws.cell(row=2, column=2).fill.fill_type is None


def test_color_touches_only_the_name_cell_of_that_person(tmp_path):
    ws = _write(tmp_path, [_violation(_FIRST, color="FFFFFF00"), _violation(_SECOND)], {})

    assert ws.cell(row=3, column=2).fill.fill_type is None
    assert all(ws.cell(row=2, column=col).fill.fill_type is None for col in (1, 3, 4, 5, 6, 7))


def test_long_task_is_wrapped_and_gets_a_taller_row(tmp_path):
    """Завдання бувають у кілька речень: їх має бути видно повністю."""
    ws = _write(tmp_path, [_violation(_FIRST), _violation(_SECOND)], {
        normalize_name(_FIRST): {"actual": _LONG_TASK, "documents": "коротко"},
        normalize_name(_SECOND): {"actual": "коротко", "documents": "коротко"},
    })

    assert ws.cell(row=2, column=3).alignment.wrap_text is True
    assert ws.row_dimensions[2].height > ws.row_dimensions[3].height


def test_name_column_fits_the_longest_name(tmp_path):
    long_name = "ДУЖЕДОВГЕПРІЗВИЩЕДОВГЕПРІЗВИЩЕ Перший Перший"

    ws = _write(tmp_path, [_violation(long_name)], {})

    assert ws.column_dimensions["B"].width > 30
