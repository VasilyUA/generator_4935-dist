import os
import sys
from pathlib import Path

_PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.append(str(_PROJECT_DIR))


def pytest_sessionstart(session):
    # constants.py/index.py читають resources/... відносно generator_timesheet/,
    # тож cwd під час тестів має бути цією директорією - незалежно від того,
    # звідки саме запущено pytest (той самий підхід, що й у сусідніх проєктах).
    os.chdir(_PROJECT_DIR)
