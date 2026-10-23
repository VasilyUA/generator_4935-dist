import sys
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parent.parent))


class FakePrompt:
    """Заглушка InquirerPy для тестового середовища без реальної консолі -
    повертає default (як Enter), або значення з черги inquirer_inputs."""

    _queue = []

    def __init__(self, *args, **kwargs):
        self._kwargs = kwargs

    def execute(self):
        if FakePrompt._queue:
            return FakePrompt._queue.pop(0)
        return self._kwargs.get("default")


@pytest.fixture
def inquirer_inputs(monkeypatch):
    """Скриптує InquirerPy-промпти в get_data.py. Патчимо саме прив'язані в
    get_data імена (не глобальний клас у самому пакеті InquirerPy) - інакше
    цей патч конфліктував би з аналогічним у generator_br_and_report_for_money/tests/conftest.py,
    який патчить ті самі глобальні класи для СВОЇХ тестів."""
    import get_data
    monkeypatch.setattr(get_data, "InputPrompt", FakePrompt)
    monkeypatch.setattr(get_data, "ListPrompt", FakePrompt)
    FakePrompt._queue = []
    yield FakePrompt._queue
    FakePrompt._queue = []
