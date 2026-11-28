# Питання "Що зробити?" - НАВМИСНЕ окремо від user_input.py/constants.py, щоб
# index.py міг спитати режим ДО важких імпортів. constants.py САМ ставить два
# питання ("Введіть номер місяця..."/"Відкрити папку проєкту (з папками
# етапів)?") прямо під час імпорту - а user_input.py імпортує constants.py на
# своєму верхньому рівні, тож БУДЬ-ЯКИЙ імпорт з user_input.py (навіть просто
# рядкової константи режиму) уже "коштує" обох цих питань. Для деяких режимів
# (напр. RUN_MODE_CHECK_REPORT_LOG_WAR) вони геть не по темі (підтверджено
# користувачем) - тож index.py питає режим ЦИМ модулем ПЕРШИМ, і лише тоді
# вирішує, чи взагалі варто ставити ці два питання (SKIP_CONSTANTS_INTERACTIVE_PROMPTS).
from InquirerPy.prompts.list import ListPrompt

RUN_MODE_ALL = "all"
RUN_MODE_BR_ONLY = "br_only"
RUN_MODE_REPORT_ONLY = "report_only"
RUN_MODE_COMMANDER_REPORT_ONLY = "commander_report_only"
RUN_MODE_CHANGES_ONLY = "changes_only"
RUN_MODE_CHECK_ACCOUNTING = "check_accounting"
RUN_MODE_TIMETABLE_ONLY = "timetable_only"
RUN_MODE_VERIFY_REPORT_ONLY = "verify_report_only"
RUN_MODE_CHECK_REPORT_LOG_WAR = "check_report_log_war"
RUN_MODE_CHECK_REPORT_PERIODS = "check_report_periods"

# Режими, для яких MONTH/"Відкрити папку проєкту?" (обидва - constants.py, під
# час імпорту) НЕ потрібні: інструмент сам визначає потрібні дати з обраного
# файлу/папки, MONTH у ньому ніде не використовується (підтверджено
# користувачем - ці питання були геть не по темі обраного інструменту). НЕ
# включає RUN_MODE_TIMETABLE_ONLY/RUN_MODE_VERIFY_REPORT_ONLY - вони, хоч і
# так само окремі інструменти, УСЕ ОДНО читають ОБЛІК.xlsx САМЕ за
# constants.MONTH/YEAR (generators.generate_timetable._month_end_or_death_date/
# checker_accounting.report_checker) - там питання місяця лишається потрібним.
# RUN_MODE_CHECK_REPORT_PERIODS ТЕЖ тут - на відміну від VERIFY_REPORT_ONLY,
# він узагалі не читає ОБЛІК.xlsx (підтверджено користувачем: окремий легкий
# інструмент - лише сам рапорт, без звірки з жодним зовнішнім джерелом).
MODES_WITHOUT_CONSTANTS_PROMPTS = (RUN_MODE_CHECK_REPORT_LOG_WAR, RUN_MODE_CHECK_REPORT_PERIODS)


def ask_run_mode():
    return ListPrompt(
        message="Що зробити?",
        choices=[
            {"name": "Звірка обліку (RESULT_accounting.xlsx з resources/check)", "value": RUN_MODE_CHECK_ACCOUNTING},
            {"name": "Все разом (бойові розпорядження + рапорт на додаткову винагороду)", "value": RUN_MODE_ALL},
            {"name": "Тільки бойові розпорядження", "value": RUN_MODE_BR_ONLY},
            {"name": "Тільки рапорт на додаткову винагороду", "value": RUN_MODE_REPORT_ONLY},
            {"name": "Тільки рапорт на поправки в наказі (за попередні місяці)", "value": RUN_MODE_CHANGES_ONLY},
            {"name": "Тільки рапорт на командира/ТВО", "value": RUN_MODE_COMMANDER_REPORT_ONLY},
            {"name": "Тільки табель обліку робочого часу (за обраний місяць)", "value": RUN_MODE_TIMETABLE_ONLY},
            {"name": "Звірка рапорту з ОБЛІК.xlsx (періоди/дні/ПІБ)", "value": RUN_MODE_VERIFY_REPORT_ONLY},
            {"name": "Перевірка рапорту та ЖБД", "value": RUN_MODE_CHECK_REPORT_LOG_WAR},
            {"name": "Перевірка рапорту (лише періоди/дні)", "value": RUN_MODE_CHECK_REPORT_PERIODS},
        ],
        default=RUN_MODE_ALL,
    ).execute()
