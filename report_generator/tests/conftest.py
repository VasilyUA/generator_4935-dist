import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parent.parent))

import pytest
import InquirerPy.prompts.checkbox as _checkbox_mod
import InquirerPy.prompts.input as _input_mod
import InquirerPy.prompts.number as _number_mod
import InquirerPy.prompts.list as _list_mod


class _BaseFakePrompt:
    """Заглушка InquirerPy для тестового середовища без реальної консолі.

    За замовчуванням execute() повертає default (те саме, що Enter без
    введення). Тести, яким потрібна конкретна "введена" відповідь, скриптують
    чергу через фікстуру inquirer_inputs. last_kwargs зберігає аргументи
    ОСТАННЬОГО створеного промпту - дозволяє дістати й напряму викликати
    validate/invalid_message, не проходячи через реальний InquirerPy цикл."""

    _queue = []
    last_kwargs = None

    def __init__(self, *args, **kwargs):
        self.kwargs = kwargs
        _BaseFakePrompt.last_kwargs = kwargs

    def execute(self):
        if _BaseFakePrompt._queue:
            return _BaseFakePrompt._queue.pop(0)
        return self._default()

    def _default(self):
        return self.kwargs.get("default", "")


class FakeCheckboxPrompt(_BaseFakePrompt):
    """CheckboxPrompt.execute() повертає СПИСОК обраних значень (не скаляр)."""

    def _default(self):
        choices = self.kwargs.get("choices") or []
        return [choices[0]] if choices else []


FakePrompt = _BaseFakePrompt

_checkbox_mod.CheckboxPrompt = FakeCheckboxPrompt
_input_mod.InputPrompt = FakePrompt
_number_mod.NumberPrompt = FakePrompt
_list_mod.ListPrompt = FakePrompt


def _reassert_fake_prompts():
    # InquirerPy.prompts.* - модулі СПІЛЬНІ для всього pytest-процесу (не
    # копіюються per-проєкт, на відміну від "constants"/"helpers"/... у
    # кореневому conftest.py _COLLIDING_MODULES). generator_br_and_report_for_money
    # та report_generator ОБИДВА так само глобально підміняють ListPrompt/
    # InputPrompt на СВОЇ класи - який саме лишиться діючим після того, як усі
    # conftest.py прогону завантажаться, залежить від порядку (останній
    # виграє), а не від того, чий тест зараз збирається.
    _checkbox_mod.CheckboxPrompt = FakeCheckboxPrompt
    _input_mod.InputPrompt = FakePrompt
    _number_mod.NumberPrompt = FakePrompt
    _list_mod.ListPrompt = FakePrompt


def pytest_collectstart(collector):
    # Хук, як і в кореневому conftest.py (pytest_collectstart активний лише
    # для файлів У ВЛАСНОМУ піддереві - тут це tests/), перевстановлює ЦІ
    # ЧОТИРИ атрибути перед КОЖНИМ збором файлу цього проєкту - тож навіть
    # якщо інший проєкт перезаписав їх останнім, до ПЕРШОГО імпорту будь-якого
    # модуля report_generator (helpers.py тощо, на збиранні файлу) вони знову
    # вказують на класи звідси.
    _reassert_fake_prompts()


def pytest_runtest_setup(item):
    # Той самий self-heal, але перед ВИКОНАННЯМ кожного тесту, не лише перед
    # збором файлів - лише collectstart тут не ловить випадок, коли якийсь
    # модуль report_generator імпортується ВПЕРШЕ не на рівні файлу, а
    # УСЕРЕДИНІ окремої тестової функції (вже під час виконання, коли збір
    # усіх файлів давно завершився, і будь-який тест ІНШОГО проєкту, що
    # виконався тим часом, міг знову перезаписати ListPrompt/InputPrompt/etc
    # своєю заглушкою).
    _reassert_fake_prompts()


@pytest.fixture
def inquirer_inputs():
    """Дозволяє заскриптувати послідовність "введених" значень для InquirerPy-промптів."""
    _BaseFakePrompt._queue = []
    yield _BaseFakePrompt._queue
    _BaseFakePrompt._queue = []
