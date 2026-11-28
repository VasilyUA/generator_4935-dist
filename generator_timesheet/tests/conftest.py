import os
import sys
from pathlib import Path

_PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.append(str(_PROJECT_DIR))

# Логіка табеля перевіряється на стандартних статусах (utils/default_statuses.py),
# а не на тих, що користувач налаштував у resources/НАЛАШТУВАННЯ.xlsx: інакше
# зміна кольору чи виплати статусу в Excel ламала б тести (constants.py).
os.environ["TIMESHEET_STATUSES_FROM_CODE"] = "1"


def pytest_sessionstart(session):
    # constants.py/index.py читають resources/... відносно generator_timesheet/,
    # тож cwd під час тестів має бути цією директорією - незалежно від того,
    # звідки саме запущено pytest (той самий підхід, що й у сусідніх проєктах).
    os.chdir(_PROJECT_DIR)
