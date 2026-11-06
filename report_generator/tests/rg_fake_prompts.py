"""Заглушки InquirerPy для тестів report_generator.

Окремий модуль (а не conftest.py), бо тести імпортують ці класи напряму:
`from conftest import ...` у спільному pytest-процесі кількох проєктів
підхоплював conftest.py ТОГО проєкту, який завантажився останнім, - унікальна
назва модуля завжди дає ті самі класи, що й патчить conftest.py цього проєкту."""


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
