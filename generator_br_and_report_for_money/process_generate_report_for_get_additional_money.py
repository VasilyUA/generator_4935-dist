# Реекспорт - жодної власної логіки: index.py (і тести) імпортують усі три
# генератори рапортів через ЦЕЙ єдиний модуль, не звертаючись до generators/*
# напряму.
from generators.generate_report_for_get_money import (
    process_generate_report_for_get_money,
    process_generate_changes_report,
)
from generators.generate_report_for_commander_money import process_generate_report_for_commander_money
