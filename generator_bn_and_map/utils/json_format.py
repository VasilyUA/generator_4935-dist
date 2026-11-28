"""Компактний, але читабельний JSON для файлів, які штаб править руками:
короткі списки й об'єкти пишуться в один рядок (як-от координати чи пара
SIDC + опис), довгі - по рядку на елемент з відступом."""
import json

LINE_WIDTH = 120


def _inline(value):
    return json.dumps(value, ensure_ascii=False, separators=(", ", ": "))


def _format(value, width, indent, level, used):
    """used - скільки символів рядка вже зайнято перед value (відступ і ключ)."""
    inline = _inline(value)
    if not isinstance(value, (dict, list)) or not value or used + len(inline) <= width:
        return inline
    pad = " " * (indent * (level + 1))
    if isinstance(value, dict):
        lines = []
        for key, item in value.items():
            head = f"{pad}{_inline(key)}: "
            lines.append(head + _format(item, width, indent, level + 1, len(head)))
        brackets = "{}"
    else:
        lines = [pad + _format(item, width, indent, level + 1, len(pad)) for item in value]
        brackets = "[]"
    return brackets[0] + "\n" + ",\n".join(lines) + "\n" + " " * (indent * level) + brackets[1]


def dumps_compact(value, width=LINE_WIDTH, indent=2):
    """value -> текст JSON: контейнер в один рядок, якщо разом з відступом і
    ключем вміщається в width символів, інакше - по елементу на рядок."""
    return _format(value, width, indent, 0, 0)
