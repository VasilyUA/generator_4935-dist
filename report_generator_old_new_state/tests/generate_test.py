import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from generate import generate
from template import change_position as change_position_mod
from template import handed_position as handed_position_mod
from template import accepted_position as accepted_position_mod


def _row_move(from_id, to_id, commander_id, higher_id):
    return {
        "НОМЕР ПОСАДИ З ЯКОЇ ПЕРЕМІЩУЄТЬСЯ ОСОБА": from_id,
        "НОМЕР ПОСАДИ НА ЯКУ ПЕРЕМІЩУЄТЬСЯ ОСОБА": to_id,
        "КОМАНДИР В СТАРІЙ ШТАТІ ЯКИЙ КЛОПОЧЕ КОМАНДИРУ БАТАЛЬЙОНУ": commander_id,
        "КОМАНДИР В НОВОМУ ШТАТІ ЯКИЙ КЛОПОЧЕ КОМАНДИРУ БАТАЛЬЙОНУ": commander_id,
        "КОМАНДИР БАТАЛЬЙОНУ В СТАРІЙ ШТАТІ": higher_id,
        "КОМАНДИР БАТАЛЬЙОНУ В НОВОМУ ШТАТІ": higher_id,
    }


def _personel_row(pos_id, pib, posada="Стрілець", rank="сержант"):
    return {
        "№ з.п.": pos_id, "ПІБ": pib, "Посада": posada,
        "підрозділ повністю": "1 рота", "звання фактичне": rank,
        "звання за штатом": rank, "ВОС": "000000А/000",
    }


def _task_row(pib):
    return {
        "ПІБ називний": pib, "Посада називний": "Навідник", "посада називний": "навідник",
        "посада родовий": "навідника", "ПІБ родовий": pib, "Посада давальний": "Навіднику",
    }


def test_generate_creates_change_handed_and_accepted_documents(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(change_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(handed_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(accepted_position_mod, "OUTPUT_DIR", str(tmp_path))

    from_pib, to_pib = "ПЕРШИЙ Перший Перший", "ДРУГИЙ Другий Другий"
    commander_pib, higher_pib = "ТРЕТІЙ Третій Третій", "ЧЕТВЕРТИЙ Четвертий Четвертий"
    rows_personel_old = [
        _personel_row(1, from_pib),
        _personel_row(3, commander_pib, posada="Командир роти", rank="капітан"),
        _personel_row(4, higher_pib, posada="Командир батальйону", rank="підполковник"),
    ]
    rows_personel_new = [
        _personel_row(2, to_pib),
        _personel_row(3, commander_pib, posada="Командир роти", rank="капітан"),
        _personel_row(4, higher_pib, posada="Командир батальйону", rank="підполковник"),
    ]
    rows_task = [_task_row(to_pib), _task_row(commander_pib)]
    rows_move = [_row_move(1, 2, 3, 4)]

    generate(rows_move, rows_task, rows_personel_old, rows_personel_new, [], "15.08.2026")

    created = {p.name for p in tmp_path.iterdir()}
    assert "РАПОРТ НА ПЕРЕМІЩЕННЯ.docx" in created
    assert f"{from_pib} здав посаду.docx" in created
    assert f"{to_pib} прийняв посаду.docx" in created
    assert capsys.readouterr().out.count("- ") == 3


def test_generate_skips_row_when_validation_fails(tmp_path, monkeypatch, capsys):
    """DataValidation.is_valid() завжди True сьогодні (validator.py) - гілку
    "не пройшла валідація" перевіряємо примусово через monkeypatch."""
    import validator

    monkeypatch.setattr(change_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(handed_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(accepted_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(validator.DataValidation, "is_valid", lambda self: False)

    from_pib, to_pib = "ШОСТИЙ Шостий Шостий", "СЬОМИЙ Сьомий Сьомий"
    commander_pib, higher_pib = "ВОСЬМИЙ Восьмий Восьмий", "ДВАДЦЯТИЙ Двадцятий Двадцятий"
    rows_personel_old = [
        _personel_row(1, from_pib),
        _personel_row(3, commander_pib, posada="Командир роти", rank="капітан"),
        _personel_row(4, higher_pib, posada="Командир батальйону", rank="підполковник"),
    ]
    rows_personel_new = [
        _personel_row(2, to_pib),
        _personel_row(3, commander_pib, posada="Командир роти", rank="капітан"),
        _personel_row(4, higher_pib, posada="Командир батальйону", rank="підполковник"),
    ]
    rows_task = [_task_row(to_pib), _task_row(commander_pib)]
    rows_move = [_row_move(1, 2, 3, 4)]

    generate(rows_move, rows_task, rows_personel_old, rows_personel_new, [], "15.08.2026")

    # DataValidation.is_valid() підмінено на False - і власний цикл
    # change_position_report (перед основним циклом generate) теж пропускає
    # РІВНО той самий рядок, тож жоден документ узагалі не зберігається.
    created = {p.name for p in tmp_path.iterdir()}
    assert created == set()
    assert capsys.readouterr().out.count("Валідація не пройшла") == 2
