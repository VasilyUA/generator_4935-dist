"""Створює ПРИКЛАД resources/ з вигаданими даними (ті самі синтетичні штат,
зразок і сторінка Статуту, що й у тестах - tests/duties_fixtures.py), щоб було
видно структуру файлів. Не чіпає resources/, якщо вона вже існує.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Після цього: python index.py → «Оновити таблицю обов'язків зі штату», потім
«Згенерувати документи». Справжні штат, зразок і Статут - див. README.md."""
import importlib.util
import json
import os
import sys

_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESOURCES_DIR = os.path.join(_PROJECT_DIR, "resources")
sys.path.insert(0, _PROJECT_DIR)


def _load_fixtures():
    path = os.path.join(_PROJECT_DIR, "tests", "duties_fixtures.py")
    spec = importlib.util.spec_from_file_location("duties_sample_fixtures", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    if os.path.exists(RESOURCES_DIR):
        print(f"{RESOURCES_DIR} вже існує - нічого не роблю.")
        return
    fixtures = _load_fixtures()
    os.makedirs(RESOURCES_DIR)
    fixtures.write_roster(os.path.join(RESOURCES_DIR, "Штат приклад.xlsx"))
    fixtures.write_sample(os.path.join(RESOURCES_DIR, "Зразки приклад.docx"))
    fixtures.write_statute(os.path.join(RESOURCES_DIR, "Статут приклад.htm"))
    with open(os.path.join(RESOURCES_DIR, "data.json"), "w", encoding="utf-8") as f:
        json.dump(fixtures.DATA_EXAMPLE, f, ensure_ascii=False, indent=2)
    print(f"Створено приклад {RESOURCES_DIR} (вигадані дані). Запустіть python index.py.")


if __name__ == "__main__":
    main()
