"""Створює ПРИКЛАД resources/ з вигаданими даними (ті самі синтетичні БР, ПБД,
РОП_ВОП, РР, позатермінове донесення, таблиця сигналів і дані, що й у тестах -
tests/bn_fixtures.py) - щоб було видно структуру файлів: data.json (ручні
поля) + data_auto.json (веде програма). Не чіпає resources/, якщо вона вже
існує.
Запуск: python resources_example/generate_sample.py (з кореня проєкту).

Після цього: python index.py, на запитання "Оновити … з ПБД" - "Так".
Реальні БР/ПБД/РОП_ВОП і тексти БН замініть своїми (див. README.md)."""
import importlib.util
import os
import sys
from datetime import date

_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESOURCES_DIR = os.path.join(_PROJECT_DIR, "resources")
sys.path.insert(0, _PROJECT_DIR)

from content.data_store import save_data, split_data  # noqa: E402


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
    folder = os.path.join(RESOURCES_DIR, "ПБД")
    os.makedirs(folder)
    today = date.today().strftime("%d.%m.%Y")
    fixtures.write_docx(os.path.join(RESOURCES_DIR, "БР приклад 1234.docx"), fixtures.BR_PARAGRAPHS)
    fixtures.write_docx(os.path.join(folder, f"Додаток №1 ПБД приклад {today}.docx"), fixtures.PBD_PARAGRAPHS)
    fixtures.write_rop_vop(os.path.join(folder, f"Додаток 2 (РОП_ВОП) приклад {today}.xlsx"))
    fixtures.write_docx(os.path.join(RESOURCES_DIR, "РР приклад 77.docx"), fixtures.RR_PARAGRAPHS)
    fixtures.write_docx(os.path.join(RESOURCES_DIR, "ДОНЕСЕННЯ_ПОЗАТЕРМІНОВЕ приклад.docx"), fixtures.RENAME_PARAGRAPHS)
    fixtures.write_signal_table(os.path.join(RESOURCES_DIR, "Таблиця сигналів приклад.docx"))
    manual, auto = split_data(fixtures.make_data())
    save_data(os.path.join(RESOURCES_DIR, "data_auto.json"), auto)
    save_data(os.path.join(RESOURCES_DIR, "data.json"), manual)
    print(f"Створено приклад {RESOURCES_DIR} (вигадані дані). Запустіть python index.py і відповідайте «Так» на оновлення з ПБД.")


if __name__ == "__main__":
    main()
