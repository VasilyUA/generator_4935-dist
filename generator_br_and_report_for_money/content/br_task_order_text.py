from constants import BR_TASK_ORDER_TEXTS, TASK_ORDER_VARIANTS

# Реальний текст документа "№.. ЗАВДАННЯ .." (координати, оцінка противника, варіанти
# завдань підрозділам) - у книзі налаштувань resources/НАЛАШТУВАННЯ.xlsx (аркуші
# "Тексти ЗАВДАННЯ" і "Варіанти ЗАВДАННЯ") - resources/ не потрапляє в git (див.
# .gitignore), тому реальні координати не лежать у вихідному коді.
BR_ENEMY_ASSESSMENT_TEXT = BR_TASK_ORDER_TEXTS["BR_ENEMY_ASSESSMENT_TEXT"]
BR_FIRE_SUPPORT_INTRO_TEMPLATE = BR_TASK_ORDER_TEXTS["BR_FIRE_SUPPORT_INTRO_TEMPLATE"]
BR_TASK_SECTION_TEMPLATE = BR_TASK_ORDER_TEXTS["BR_TASK_SECTION_TEMPLATE"]
BR_SITUATION_UPDATE_TEMPLATE = BR_TASK_ORDER_TEXTS["BR_SITUATION_UPDATE_TEMPLATE"]
BR_DEFENSE_TASKS_TEXT = BR_TASK_ORDER_TEXTS["BR_DEFENSE_TASKS_TEXT"]
BR_DEFENSE_MEASURES_TEXT = BR_TASK_ORDER_TEXTS["BR_DEFENSE_MEASURES_TEXT"]
# Варіанти переліку завдань підрозділам після "НАКАЗАВ:" — ротуються по черзі,
# по одному на кожен фактичний БР бат (див. get_bat_period_and_variant). Порожній
# варіант -> тимчасовий фолбек на варіант 1 у generate_documents_br_weekly_task.py.
# TASK_ORDER_VARIANTS ([[(підрозділ, завдання), ...], ...]) будується з книги в
# utils/project_settings.py::build_task_order_variants і тут лише реекспортується.
__all__ = [
    "BR_ENEMY_ASSESSMENT_TEXT", "BR_FIRE_SUPPORT_INTRO_TEMPLATE", "BR_TASK_SECTION_TEMPLATE",
    "BR_SITUATION_UPDATE_TEMPLATE", "BR_DEFENSE_TASKS_TEXT", "BR_DEFENSE_MEASURES_TEXT", "TASK_ORDER_VARIANTS",
]
