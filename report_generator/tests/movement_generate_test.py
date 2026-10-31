import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from movement_generate import mass_generate
from template import change_position as change_position_mod
from template import handed_position as handed_position_mod
from template import accepted_position as accepted_position_mod


def _personel_row(pos_id, pib):
    return {
        "№": pos_id,
        "ВОС": "000000А/000",
        "звання за штатом": "сержант",
        "звання фактичне": "сержант",
    }


def _task_row(pos_id, pib):
    return {
        "№ посади": pos_id,
        "посада називний": "Стрілець",
        "посада родовий": "стрільця",
        "посада давальний": "стрільцю",
        "Посада давальний": "Стрільцю",
        "ПІБ називний": pib,
        "ПІБ родовий": pib,
        "ПІБ давальний": pib,
    }


def _rows_move(from_id, to_id, commander_id, higher_id):
    return {
        "НОМЕР ПОСАДИ З ЯКОЇ ПЕРЕМІЩУЄТЬСЯ ОСОБА": from_id,
        "НОМЕР ПОСАДИ НА ЯКУ ПЕРЕМІЩУЄТЬСЯ ОСОБА": to_id,
        "КОМАНДИР ВЗВОДУ ЧИ РОТИ АБО ЇХ ТВО ЯКІ КЛОПОЧУТЬ КОМАНДИРУ БАТАЛЬЙОНУ": commander_id,
        "КОМАНДИР БАТАЛЬЙОНУ АБО ЙОГО ТВО": higher_id,
    }


def test_mass_generate_single_row_creates_all_three_documents(tmp_path, monkeypatch):
    monkeypatch.setattr(change_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(handed_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(accepted_position_mod, "OUTPUT_DIR", str(tmp_path))

    ids = {"from": 1, "to": 2, "commander": 3, "higher": 4}
    pibs = {1: "ПЕРШИЙ Перший Перший", 2: "ДРУГИЙ Другий Другий", 3: "ТРЕТІЙ Третій Третій", 4: "ЧЕТВЕРТИЙ Четвертий Четвертий"}
    rows_personel = [_personel_row(v, pibs[v]) for v in ids.values()]
    rows_task = [_task_row(v, pibs[v]) for v in ids.values()]
    rows_move = [_rows_move(ids["from"], ids["to"], ids["commander"], ids["higher"])]

    mass_generate(rows_move, rows_task, rows_personel)

    created = {p.name for p in tmp_path.iterdir()}
    assert "РАПОРТ НА ПЕРЕМІЩЕННЯ.docx" in created
    assert "ПЕРШИЙ Перший Перший здав посаду.docx" in created
    # data_for_soldier_to успадковує ім'я/звання ВІД data_for_soldier_from перед
    # generate_accepted_position_report (mass_generate.py) - "прийняв посаду"
    # підписаний тим самим (реальним) імʼям людини, що переміщується, а не
    # умовним "новим" ПІБ з rows_task за НОВОЮ посадою.
    assert "ПЕРШИЙ Перший Перший прийняв посаду.docx" in created


def test_mass_generate_two_rows_produces_pairwise_handed_and_accepted_docs(tmp_path, monkeypatch):
    monkeypatch.setattr(change_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(handed_position_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(accepted_position_mod, "OUTPUT_DIR", str(tmp_path))

    pibs = {
        1: "ПЕРШИЙ Перший Перший", 2: "ДРУГИЙ Другий Другий",
        3: "ТРЕТІЙ Третій Третій", 4: "ЧЕТВЕРТИЙ Четвертий Четвертий",
        5: "ШОСТИЙ Шостий Шостий", 6: "СЬОМИЙ Сьомий Сьомий",
    }
    rows_personel = [_personel_row(v, pib) for v, pib in pibs.items()]
    rows_task = [_task_row(v, pib) for v, pib in pibs.items()]
    rows_move = [
        _rows_move(1, 2, 3, 4),
        _rows_move(5, 6, 3, 4),
    ]

    mass_generate(rows_move, rows_task, rows_personel)

    created = {p.name for p in tmp_path.iterdir()}
    assert created == {
        "РАПОРТ НА ПЕРЕМІЩЕННЯ.docx",
        "ПЕРШИЙ Перший Перший здав посаду.docx",
        "ПЕРШИЙ Перший Перший прийняв посаду.docx",
        "ШОСТИЙ Шостий Шостий здав посаду.docx",
        "ШОСТИЙ Шостий Шостий прийняв посаду.docx",
    }
