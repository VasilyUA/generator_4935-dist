from constants import DATA_FILE_NAME
import json

# Реальний текст документа "№.. ЗАВДАННЯ .." (координати, оцінка противника, варіанти
# завдань підрозділам) винесено в resources/data.json (ключ "br_task_order") —
# resources/ не потрапляє в git (див. .gitignore), тому реальні координати не лежать
# у вихідному коді.
with open(DATA_FILE_NAME, "r", encoding="utf-8") as _f:
    _data = json.load(_f)["br_task_order"]

BR_ENEMY_ASSESSMENT_TEXT = _data["BR_ENEMY_ASSESSMENT_TEXT"]
BR_FIRE_SUPPORT_INTRO_TEMPLATE = _data["BR_FIRE_SUPPORT_INTRO_TEMPLATE"]
BR_TASK_SECTION_TEMPLATE = _data["BR_TASK_SECTION_TEMPLATE"]
BR_SITUATION_UPDATE_TEMPLATE = _data["BR_SITUATION_UPDATE_TEMPLATE"]
BR_DEFENSE_TASKS_TEXT = _data["BR_DEFENSE_TASKS_TEXT"]
BR_DEFENSE_MEASURES_TEXT = _data["BR_DEFENSE_MEASURES_TEXT"]
# Три варіанти переліку завдань підрозділам після "НАКАЗАВ:" — ротуються по черзі,
# по одному на кожен фактичний БР бат (див. get_bat_period_and_variant). Порожній
# варіант -> тимчасовий фолбек на варіант 1 у generate_documents_br_weekly_task.py.
TASK_ORDER_VARIANTS = [[tuple(pair) for pair in variant] for variant in _data["TASK_ORDER_VARIANTS"]]
