"""Створює ПРИКЛАД resources/ з вигаданими даними (ті самі синтетичні БР, ПБД,
РОП_ВОП і data.json, що й у тестах - tests/bn_fixtures.py) - щоб було видно
структуру файлів. Не чіпає resources/, якщо вона вже існує.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Після цього: python index.py, на запитання "Оновити … з ПБД" - "Так".
Реальні БР/ПБД/РОП_ВОП і тексти БН замініть своїми (див. README.md)."""
import importlib.util
import json
import os
from datetime import date

_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESOURCES_DIR = os.path.join(_PROJECT_DIR, "resources")


def _load_fixtures():
    path = os.path.join(_PROJECT_DIR, "tests", "bn_fixtures.py")
    spec = importlib.util.spec_from_file_location("bn_sample_fixtures", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    if os.path.exists(RESOURCES_DIR):
        print(f"{RESOURCES_DIR} вже існує - нічого не роблю.")
        return
    fixtures = _load_fixtures()
    folder = os.path.join(RESOURCES_DIR, f"{date.today().strftime('%d.%m.%Y')} приклад")
    os.makedirs(folder)
    fixtures.write_docx(os.path.join(RESOURCES_DIR, "БР приклад 1234.docx"), fixtures.BR_PARAGRAPHS)
    fixtures.write_docx(os.path.join(folder, "Додаток №1 ПБД приклад.docx"), fixtures.PBD_PARAGRAPHS)
    fixtures.write_rop_vop(os.path.join(folder, "Додаток 2 (РОП_ВОП) приклад.xlsx"))
    with open(os.path.join(RESOURCES_DIR, "data.json"), "w", encoding="utf-8") as f:
        json.dump(fixtures.make_data(), f, ensure_ascii=False, indent=2)
    print(f"Створено приклад {RESOURCES_DIR} (вигадані дані). Запустіть python index.py і відповідайте «Так» на оновлення з ПБД.")


if __name__ == "__main__":
    main()
