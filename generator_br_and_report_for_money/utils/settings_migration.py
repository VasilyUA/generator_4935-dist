"""Одноразове перенесення колишніх налаштувань у книгу resources/НАЛАШТУВАННЯ.xlsx.

Колись налаштування жили у двох місцях: resources/data.json (назви, тексти
документів, підрозділи) і прямо в constants.py (номери документів, підстави,
пункти й категорії рапорту, посади). Нова версія constants.py цих значень уже
не має, тож стара береться з історії git: лаунчер оновлює код через
"git reset --hard" (scripts/common.ps1), і попередня версія лишається в
reflog. Виконується стара версія як звичайний модуль, без жодних питань
(SKIP_CONSTANTS_INTERACTIVE_PROMPTS), а її словники розкладаються по аркушах
книги (utils/project_settings.py).

Без старої версії constants.py (немає git чи історії) правильно перенести не
вийде - лише data.json дав би книгу без підстав і категорій, і рапорти мовчки
вийшли б без підстав. Тоді - зупинка з поясненням, як отримати книгу інакше."""
import json
import os
import runpy
import shutil
import subprocess
import tempfile
from collections import Counter
from datetime import datetime

from utils import project_settings as ps
from utils import settings_book as sb
from utils.column_names import column_names_of
from utils.logging_utils import print_red

# Рядок, що є лише в старому форматі constants.py (словник номерів прямо в коді).
LEGACY_MARKER = "NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY = {"
# Скільки версій constants.py з історії переглядати в пошуках старого формату.
_MAX_HISTORY_VERSIONS = 50
# Той самий PortableGit, що ставить лаунчер (scripts/common.ps1), якщо git немає в PATH.
_PORTABLE_GIT = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "PortableGit", "cmd", "git.exe")

# Як назвати групу аркуша "Підстави" для кожної групи колишнього BR_HIGHT_UNIT.
_BRIGADE_GROUP_NAMES = {"general": "Бригада загальні"}


class MigrationError(Exception):
    pass


def _git_executable():
    found = shutil.which("git")
    if found:
        return found
    return _PORTABLE_GIT if os.path.isfile(_PORTABLE_GIT) else None


def _git(git, project_dir, *args):
    return subprocess.run(
        [git, "-c", "safe.directory=*", "-C", project_dir, *args],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60, check=False,
    )


def find_legacy_constants_source(project_dir):
    """Текст найновішої версії constants.py старого формату (LEGACY_MARKER) з
    історії git, включно з reflog, - або None."""
    git = _git_executable()
    if git is None:
        return None
    try:
        log = _git(git, project_dir, "log", "--all", "--reflog", "--format=%H", "--", "constants.py")
        for commit in log.stdout.split()[:_MAX_HISTORY_VERSIONS]:
            shown = _git(git, project_dir, "show", f"{commit}:./constants.py")
            if shown.returncode == 0 and LEGACY_MARKER in shown.stdout:
                return shown.stdout
    except (OSError, subprocess.SubprocessError):
        return None
    return None


def load_legacy_constants(source, project_dir):
    """Глобальні змінні старого constants.py (текст source), виконаного без
    питань місяця/етапу. Старий файл сам читає resources/data.json відносно
    поточної теки, тож виконується з project_dir."""
    previous_cwd = os.getcwd()
    previous_flag = os.environ.get("SKIP_CONSTANTS_INTERACTIVE_PROMPTS")
    with tempfile.TemporaryDirectory(prefix="legacy_constants_") as temp_dir:
        path = os.path.join(temp_dir, "legacy_constants.py")
        with open(path, "w", encoding="utf-8") as legacy_file:
            legacy_file.write(source)
        os.environ["SKIP_CONSTANTS_INTERACTIVE_PROMPTS"] = "1"
        try:
            os.chdir(project_dir)
            return runpy.run_path(path, run_name="legacy_constants")
        finally:
            os.chdir(previous_cwd)
            if previous_flag is None:
                os.environ.pop("SKIP_CONSTANTS_INTERACTIVE_PROMPTS", None)
            else:
                os.environ["SKIP_CONSTANTS_INTERACTIVE_PROMPTS"] = previous_flag


def _reference_row(group, entry, today, higher_hq=False):
    """Запис {"start", "end", "lines"[, "id"]} -> рядок аркуша "Підстави".
    Кінець, що дорівнює сьогоднішній даті, - це колишнє datetime.now() ("діє
    досі"), тож у книзі він порожній."""
    end = entry.get("end") or ""
    return {
        "group": group,
        "start": entry.get("start"),
        "end": "" if end == today else end,
        "lines": list(entry.get("lines", [])),
        "higher_hq": higher_hq,
        "id": entry.get("id") or "",
    }


def _same_entries(left, right):
    return len(left) == len(right) and all(a is b or a == b for a, b in zip(left, right))


class _ReferenceGroups:
    """Групи аркуша "Підстави", з яких складаються колишні списки "general"/
    "grounds": спершу відомі (ЖБД, БН, групи BR_HIGHT_UNIT), а список, якого
    ними не скласти, стає новою групою з назвою категорії чи пункту."""

    def __init__(self, today):
        self.today = today
        self.groups = []  # [(назва, [записи])]
        self.rows = []

    def add(self, name, entries, higher_hq=False):
        self.groups.append((name, list(entries)))
        self.rows.extend(_reference_row(name, entry, self.today, higher_hq) for entry in entries)

    def names_for(self, entries, *new_group_names):
        """Назви груп, з яких по черзі складається список entries. Залишок, якого
        відомими групами не скласти, - нова група з першою вільною назвою з
        new_group_names (назва категорії, потім пункту)."""
        names, position = [], 0
        entries = list(entries)
        while position < len(entries):
            for name, group_entries in self.groups:
                if group_entries and _same_entries(entries[position:position + len(group_entries)], group_entries):
                    names.append(name)
                    position += len(group_entries)
                    break
            else:
                name = self._free_name(new_group_names)
                self.add(name, entries[position:])
                names.append(name)
                break
        return names

    def _free_name(self, candidates):
        taken = {sb.normalize_name(name) for name, _entries in self.groups}
        for candidate in candidates:
            if sb.normalize_name(candidate) not in taken:
                return str(candidate)
        name, number = str(candidates[0]), 2
        while sb.normalize_name(name) in taken:
            name, number = f"{candidates[0]} ({number})", number + 1
        return name


def _column_names(value):
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    return [str(item) for item in value]


def _exclude_text(value):
    if value == "all":
        return "всі"
    return ", ".join(str(item) for item in value or [])


def _categories_rows(categories, groups):
    """Колишній словник MONEY_REPORT_CATEGORIES -> (рядки пунктів, рядки категорій)."""
    point_rows, category_rows = [], []
    for point, config in categories.items():
        point_rows.append({
            "point": str(point),
            "general": groups.names_for(config.get("general", []), f"Загальні {point}"),
            "column": _column_names(config.get("column_basis_from_tracking")),
        })
        for name, category in config.items():
            if name in ("general", "column_basis_from_tracking") or not isinstance(category, dict):
                continue
            category_rows.append({
                "point": str(point),
                "category": name,
                "grounds": groups.names_for(category.get("grounds", []), name, str(point), f"{point} {name}"),
                "exclude_general": _exclude_text(category.get("exclude_general", [])),
                "use_brs": bool(category.get("use_brs", False)),
                "use_brs_from_selected_folder": bool(category.get("use_brs_from_selected_folder", False)),
                "default": bool(category.get("default", False)),
                "generate_br": bool(category.get("generate_br", True)),
                "include_to_report": bool(category.get("include_to_report", True)),
                "required_basis": bool(category.get("required_basis", True)),
                "column": _column_names(category.get("column_basis_from_tracking")),
            })
    return point_rows, category_rows


def _text_values(section):
    """Тексти розділу data.json -> {параметр: текст}; вкладені значення - JSON-рядком."""
    values = {}
    for key, value in section.items():
        values[key] = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
    return values


def _subdivision_rows(company_sections):
    """SEQUENCE -> рядки таблиці "Підрозділи" (порядок рядка "КОМАНДИРУ ...")."""
    return [{"name": name} for name in company_sections.get("SEQUENCE", [])]


def _subdivision_text_rows(company_sections):
    """PARAGRAPH_MAP(_FOR_EXTRACT) -> рядки таблиці "Тексти підрозділів" у
    порядку абзаців БР. Назва - як у SEQUENCE, якщо ключ абзацу відповідає
    підрозділу переліку ("1рмп" -> "1 рмп"), інакше сам ключ ("упр, штаб")."""
    paragraphs = company_sections.get("PARAGRAPH_MAP", {})
    extracts = company_sections.get("PARAGRAPH_MAP_FOR_EXTRACT", {})
    names_by_key = {ps.section_key(name): name for name in company_sections.get("SEQUENCE", [])}
    return [
        {"name": names_by_key.get(ps.section_key(key), key), "br_text": paragraphs.get(key, ""), "extract_text": extracts.get(key, "")}
        for key in [*paragraphs, *(key for key in extracts if key not in paragraphs)]
    ]


def _task_variant_rows(variants):
    rows = []
    for number, variant in enumerate(variants, start=1):
        if not variant:
            rows.append({"variant": number, "title": "", "text": ""})
        rows.extend({"variant": number, "title": title, "text": text} for title, text in variant)
    return rows


def _point_spelling(key, points):
    """Ключ тексту пункту (SECTIONS/COMMANDER_SECTIONS) так, як пункт названо в
    налаштуваннях. Реальний випадок: пункт "100_РТГр" перейменували на
    "РТГр_100" лише в constants.py, а текст у data.json лишився під старою
    назвою - і рапорт падав з KeyError. Ключ, якого серед пунктів немає, але
    який відрізняється від пункту лише порядком частин через "_", стає назвою
    цього пункту."""
    if key in points:
        return key
    parts = sorted(str(key).split("_"))
    return next((point for point in points if sorted(point.split("_")) == parts), key)


def _section_texts_by_point(texts, points):
    """{ключ: текст} з ключами в написанні пунктів (_point_spelling); ключ, що
    вже є під правильною назвою, не перезаписується."""
    result = {}
    for key, text in texts.items():
        target = _point_spelling(key, points)
        if target != key and target in texts:
            target = key
        result[target] = text
    return result


def _legacy_points(legacy):
    return [
        str(point)
        for categories in (legacy["MONEY_REPORT_CATEGORIES"], legacy["COMMANDER_MONEY_REPORT_CATEGORIES"])
        for point in categories
    ]


def settings_from_legacy(data, legacy, today):
    """Словник налаштувань (utils/project_settings.py) з колишніх data.json
    (data) і глобальних змінних старого constants.py (legacy)."""
    unit = data["unit"]
    groups = _ReferenceGroups(today)
    groups.add(ps.LOG_WAR_GROUP, legacy["MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR"])
    groups.add(ps.BN_GROUP, legacy["MONEY_REPORT_GENERAL_REFERENCES_BN"])
    for name, entries in legacy["BR_HIGHT_UNIT"].items():
        groups.add(_BRIGADE_GROUP_NAMES.get(name, name), entries, higher_hq=True)

    points, categories = _categories_rows(legacy["MONEY_REPORT_CATEGORIES"], groups)
    commander_points, commander_categories = _categories_rows(legacy["COMMANDER_MONEY_REPORT_CATEGORIES"], groups)

    statik = data["data_statik_report_for_get_additional_money"]
    point_names = _legacy_points(legacy)
    sections = _section_texts_by_point(statik.get("SECTIONS", {}), point_names)
    commander_sections = _section_texts_by_point(statik.get("COMMANDER_SECTIONS", {}), point_names)
    for old_key, new_key in {
        old: _point_spelling(old, point_names)
        for old in (*statik.get("SECTIONS", {}), *statik.get("COMMANDER_SECTIONS", {}))
    }.items():
        if old_key != new_key and (new_key in sections or new_key in commander_sections):
            print_red(f'Текст пункту "{old_key}" перенесено як "{new_key}" - так цей пункт названо в налаштуваннях.')
    task_order = data["br_task_order"]
    return {
        "general": {
            **{key: unit[key] for key in ps.UNIT_KEYS},
            "HIGHER_COMMANDER_TITLE": legacy["HIGHER_COMMANDER_TITLE"],
            "COMMANDER_TITLE": legacy["COMMANDER_TITLE"],
            "SEQUENCE_HIGHER_COMMANDER": list(legacy["SEQUENCE_HIGHER_COMMANDER"]),
        },
        "numbers": {name: ps.number_rows(legacy[name], name) for name in ps.NUMBER_TABLES},
        "references": groups.rows,
        "changes_orders": [
            {key: value for key, value in _reference_row("", entry, today).items() if key in ("start", "end", "lines", "id")}
            for entry in legacy["MONEY_REPORT_CHANGES_ORDER_REFERENCES"]
        ],
        "points": points,
        "categories": categories,
        "commander_points": commander_points,
        "commander_categories": commander_categories,
        "section_texts": [
            {"point": key, "report": sections.get(key, ""), "commander": commander_sections.get(key, "")}
            for key in [*sections, *(key for key in commander_sections if key not in sections)]
        ],
        "report_texts": _text_values({k: v for k, v in statik.items() if k not in ("SECTIONS", "COMMANDER_SECTIONS")}),
        "subdivisions": _subdivision_rows(data["company_sections"]),
        "subdivision_texts": _subdivision_text_rows(data["company_sections"]),
        "task_texts": _text_values({k: v for k, v in task_order.items() if k != "TASK_ORDER_VARIANTS"}),
        "task_variants": _task_variant_rows(task_order.get("TASK_ORDER_VARIANTS", [])),
        "not_paid_exceptions": [
            {"status": status, "confirmed_by": sorted(confirming)}
            for status, confirming in (legacy.get("NOT_PAID_IGNORED_WHEN_CONFIRMED_BY") or {}).items()
        ] if "NOT_PAID_IGNORED_WHEN_CONFIRMED_BY" in legacy else [dict(row) for row in ps.DEFAULT_NOT_PAID_EXCEPTIONS],
    }


def _normalized(value):
    """Значення для порівняння: книга обрізає пробіли на краях тексту й
    зберігає перенос рядка як "\\n"; кортеж і список - одне й те саме."""
    if isinstance(value, str):
        return value.replace("\r\n", "\n").strip()
    if isinstance(value, dict):
        return {key: _normalized(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_normalized(item) for item in value]
    return value


def _comparable_categories(categories):
    """Категорії без різниці між відсутнім "general" і [], між None/рядком/
    кортежем колонки підстав (column_names_of)."""
    result = {}
    for point, config in categories.items():
        result[point] = {
            "general": config.get("general", []),
            "column_basis_from_tracking": column_names_of(config.get("column_basis_from_tracking")),
            **{
                name: {**category, "column_basis_from_tracking": column_names_of(category.get("column_basis_from_tracking"))}
                for name, category in config.items()
                if name not in ("general", "column_basis_from_tracking") and isinstance(category, dict)
            },
        }
    return result


def _existing_days(number_dict):
    """Ключі словника номерів, що є справжніми датами (колишній constants.py
    мав і "31.11.2026"; у книзі такого дня в листопаді просто немає)."""
    result = {}
    for key, entry in number_dict.items():
        try:
            datetime.strptime(key, "%d.%m.%Y")
        except ValueError:
            continue
        result[key] = entry
    return result


def _legacy_month(legacy):
    """(місяць, рік), на які колишній constants.py будував ключі номерів: його
    MONTH/YEAR, а без них - найчастіший місяць серед ключів щоденних номерів."""
    if legacy.get("MONTH") and legacy.get("YEAR"):
        return int(legacy["MONTH"]), int(legacy["YEAR"])
    months = Counter(
        (int(month), int(year))
        for _day, month, year in (str(key).split(".") for key in legacy["NUMBER_OF_DOCUMENTS_BRS_EVERY_DAY"])
    )
    return max(months, key=lambda month_year: (months[month_year], month_year[1], month_year[0]))


def migration_differences(settings_path, data, legacy, today):
    """Що з колишніх data.json і constants.py (legacy) після читання книги
    settings_path назад вийшло інакше. Порожній список - перенесено без втрат."""
    settings = ps.read_settings(settings_path)
    built = ps.build_constants(settings, today)
    statik = data["data_statik_report_for_get_additional_money"]
    sections = data["company_sections"]
    task_order = data["br_task_order"]
    numbers = dict(zip(ps.NUMBER_TABLES, ps.build_number_dicts(settings["numbers"], *_legacy_month(legacy))))
    checks = {
        "MONEY_REPORT_CATEGORIES": (
            _comparable_categories(built.MONEY_REPORT_CATEGORIES), _comparable_categories(legacy["MONEY_REPORT_CATEGORIES"]),
        ),
        "COMMANDER_MONEY_REPORT_CATEGORIES": (
            _comparable_categories(built.COMMANDER_MONEY_REPORT_CATEGORIES), _comparable_categories(legacy["COMMANDER_MONEY_REPORT_CATEGORIES"]),
        ),
        "порядок пунктів і категорій": (
            [list(config) for config in _comparable_categories(built.MONEY_REPORT_CATEGORIES).values()],
            [list(config) for config in _comparable_categories(legacy["MONEY_REPORT_CATEGORIES"]).values()],
        ),
        "MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR": (built.MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR, legacy["MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR"]),
        "MONEY_REPORT_GENERAL_REFERENCES_BN": (built.MONEY_REPORT_GENERAL_REFERENCES_BN, legacy["MONEY_REPORT_GENERAL_REFERENCES_BN"]),
        "BR_HIGHT_UNIT": (list(built.BR_HIGHT_UNIT.values()), list(legacy["BR_HIGHT_UNIT"].values())),
        "MONEY_REPORT_CHANGES_ORDER_REFERENCES": (built.MONEY_REPORT_CHANGES_ORDER_REFERENCES, legacy["MONEY_REPORT_CHANGES_ORDER_REFERENCES"]),
        "NOT_PAID_IGNORED_WHEN_CONFIRMED_BY": (
            built.NOT_PAID_IGNORED_WHEN_CONFIRMED_BY,
            legacy.get("NOT_PAID_IGNORED_WHEN_CONFIRMED_BY", built.NOT_PAID_IGNORED_WHEN_CONFIRMED_BY),
        ),
        "посади": (
            (built.HIGHER_COMMANDER_TITLE, built.COMMANDER_TITLE, built.SEQUENCE_HIGHER_COMMANDER),
            (legacy["HIGHER_COMMANDER_TITLE"], legacy["COMMANDER_TITLE"], list(legacy["SEQUENCE_HIGHER_COMMANDER"])),
        ),
        "unit": ({key: getattr(built, key) for key in ps.UNIT_KEYS}, {key: data["unit"][key] for key in ps.UNIT_KEYS}),
        "тексти рапорту": (
            {key: built.MONEY_REPORT_STATIC.get(key) for key in statik},
            {
                **statik,
                **{
                    key: _section_texts_by_point(statik[key], _legacy_points(legacy))
                    for key in ("SECTIONS", "COMMANDER_SECTIONS") if key in statik
                },
            },
        ),
        "company_sections": (
            (built.SEQUENCE, list(built.PARAGRAPH_MAP.items()), list(built.PARAGRAPH_MAP_FOR_EXTRACT.items())),
            (
                list(sections.get("SEQUENCE", [])),
                [(ps.section_key(key), text) for key, text in sections.get("PARAGRAPH_MAP", {}).items() if text],
                [(ps.section_key(key), text) for key, text in sections.get("PARAGRAPH_MAP_FOR_EXTRACT", {}).items() if text],
            ),
        ),
        "br_task_order": (
            ({key: built.BR_TASK_ORDER_TEXTS.get(key) for key in task_order if key != "TASK_ORDER_VARIANTS"}, built.TASK_ORDER_VARIANTS),
            (
                {key: value for key, value in task_order.items() if key != "TASK_ORDER_VARIANTS"},
                [[tuple(pair) for pair in variant] for variant in task_order.get("TASK_ORDER_VARIANTS", [])],
            ),
        ),
        **{
            name: (
                numbers[name],
                {
                    key: {field: entry.get(field, "") or "" for field in ps.number_fields(name)}
                    for key, entry in _existing_days(legacy[name]).items()
                },
            )
            for name in ps.NUMBER_TABLES
        },
    }
    return [name for name, (actual, expected) in checks.items() if _normalized(actual) != _normalized(expected)]


def migrate_legacy_settings(settings_path, legacy_json_path, project_dir, legacy_source=None):
    """data.json + стара версія constants.py -> книга settings_path. Старої
    версії не знайдено чи дані не читаються - зупинка з червоним поясненням,
    що зробити (sb.stop; див. docstring модуля).

    Після запису книга читається назад і звіряється з колишніми значеннями
    (migration_differences): усе збіглося - data.json видаляється (дані вже в
    книзі), інакше лишається, а користувач бачить, які розділи перевірити."""
    source = legacy_source if legacy_source is not None else find_legacy_constants_source(project_dir)
    sample_step = (f"Або покладіть у resources/ готовий {sb.SETTINGS_FILE_NAME} (напр. від колеги), або створіть зразок: "
                   "python resources_example/generate_sample.py (з теки generator_br_and_report_for_money) і "
                   "впишіть у нього свої дані.")
    if source is None:
        raise sb.stop(
            f"Немає {settings_path}, а перенести {legacy_json_path} автоматично не вдалося: не знайдено попередньої "
            "версії constants.py (з номерами документів, підставами й категоріями рапорту) в історії git.",
            (sample_step, "Запустіть програму ще раз."),
        )
    today = datetime.now().strftime("%d.%m.%Y")
    try:
        with open(legacy_json_path, encoding="utf-8") as json_file:
            data = json.load(json_file)
        legacy = load_legacy_constants(source, project_dir)
        settings = settings_from_legacy(data, legacy, today)
    except SystemExit:
        raise
    except Exception as error:
        raise sb.stop(
            f"Не вдалося перенести налаштування з {legacy_json_path} і попереднього constants.py: {error!r}",
            (f"Виправте {legacy_json_path} (відкривається Блокнотом). {sample_step}", "Запустіть програму ще раз."),
        ) from error
    ps.write_settings(settings_path, settings)

    try:
        differences = migration_differences(settings_path, data, legacy, today)
    except Exception as error:
        differences = [f"книгу не вдалося перевірити ({error!r})"]
    if differences:
        print_red(
            f"Налаштування перенесено в {settings_path}, але не все збіглося з колишніми: {', '.join(differences)}. "
            f"{legacy_json_path} лишено - перевірте ці розділи книги."
        )
        return False
    os.remove(legacy_json_path)
    return True
