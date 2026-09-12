import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

from validator import DataValidation


def test_is_valid_always_true_today():
    """Заглушка (validator.py: 'Implement validation logic') - завжди True,
    незалежно від вхідних даних. Тест фіксує ЦЮ поведінку зараз, а не те, якою
    вона МАЄ бути - якщо колись з'явиться реальна валідація, цей тест мав би
    впасти й підказати, що його треба переписати під нову логіку."""
    validator = DataValidation({
        "position_codes": {}, "rows_personel_old": [], "rows_personel_new": [],
        "rows_task": [], "rows_tvo": [],
    })
    assert validator.is_valid() is True


def test_is_valid_stores_data_as_is():
    data = {"position_codes": {"subordinate_position_code_from": 1}}
    validator = DataValidation(data)
    assert validator.data is data
