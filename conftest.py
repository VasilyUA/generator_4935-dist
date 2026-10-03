import os
import sys
from pathlib import Path

# На Windows консоль (PowerShell/cmd) за замовчуванням використовує кодову
# сторінку, що НЕ збігається з тією, у якій pytest ЗАВЖДИ кодує захоплений
# вивід внутрішньо (UTF-8, жорстко - це не залежить від sys.stdout.encoding
# і НЕ змінюється через sys.stdout.reconfigure(), перевірено) - результат:
# кожен кириличний символ "вибухає" в кілька псевдографічних символів
# ("РїС–РґСЂРѕР·РґС–Р»1" замість "підрозділ1" - типовий вигляд UTF-8, прочитаного
# як cp1251/cp437). Єдине, що реально впливає на РЕНДЕР - кодова сторінка
# САМОЇ консолі (chcp), тому переключаємо ЛИШЕ її, на 65001 (UTF-8) - те
# саме кодування, яким pytest і так усе пише.
#
# Виклик ПОВТОРЮЄТЬСЯ (не лише один раз на старті модуля): інтерактивний
# запит "Які тести запустити?" (_ask_projects нижче, через
# InquirerPy/prompt_toolkit) сам звертається до Win32 Console API для
# власного рендеру і, як побічний ефект, СКИДАЄ кодову сторінку консолі
# назад (підтверджено: chcp 1251 на старті рендерився правильно, ЛИШЕ якщо
# запит взагалі не показувався - non-tty; коли запит РЕАЛЬНО малювався й на
# нього відповідали, вивід після нього знову ламався). Тому цю саму функцію
# викликано ЩЕ РАЗ, ПІСЛЯ запиту (у _prompt_selected_projects, вже після
# capman.resume_global_capture()) - той самий клас проблеми, що й інше
# спільне глобальне (не per-проєкт) консольне/сесійне state в цьому файлі:
# потрібне повторне ствердження ПІСЛЯ кожного, хто міг його скинути, а не
# лише один раз на старті.
def _fix_windows_console_codepage():
    if sys.platform != "win32":
        return
    try:
        import subprocess
        subprocess.run(["chcp", "65001"], shell=True, capture_output=True, check=False)
    except Exception:
        pass


_fix_windows_console_codepage()

_ROOT_DIR = Path(__file__).resolve().parent

# Кореневий conftest.py потрібен, щоб корінь репозиторію був у sys.path
# незалежно від того, звідки саме запущено pytest.
os.chdir(_ROOT_DIR)

# Справжні (НЕ підмінені) класи - захоплені тут, на рівні модуля кореневого
# conftest.py, який pytest завжди завантажує ПЕРШИМ (раніше за conftest.py
# будь-якого проєкту). generator_br_and_report_for_money/tests/conftest.py й
# report_generator/tests/conftest.py ОБИДВА глобально підміняють
# InquirerPy.prompts.list.ListPrompt/.checkbox.CheckboxPrompt на власні
# заглушки (потрібні лише ВСЕРЕДИНІ їхніх власних тестів) - до моменту, коли
# нижче в цьому файлі викликається _ask_projects() (з pytest_configure, вже
# ПІСЛЯ того, як усі conftest.py відповідних testpaths встигли завантажитись),
# ці атрибути можуть вказувати на ЧУЖУ заглушку. Свіжий `from
# InquirerPy.prompts.list import ListPrompt` УСЕРЕДИНІ _ask_projects() тоді
# підхопив би ЇЇ - і замість того, щоб СПРАВДІ запитати людину, яку тести
# запустити, мовчки "відповів" би заздалегідь визначеним значенням (напр.
# FakeCheckboxPrompt.execute() без явного choices/default повертає лише
# ПЕРШИЙ проєкт зі списку) - решта проєктів тихо пропускались би, без жодної
# помилки чи попередження. Захоплюючи справжні класи ОДРАЗУ тут (до того, як
# хоч один проєктний conftest.py взагалі міг завантажитись), _ask_projects()
# нижче гарантовано працює зі СПРАВЖНІМ InquirerPy - так само, як бачив би
# його користувач у реальному, не тестовому запуску.
from InquirerPy.prompts.checkbox import CheckboxPrompt as _RealCheckboxPrompt
from InquirerPy.prompts.list import ListPrompt as _RealListPrompt

# Усі три - "пласкі" проєкти (без пакетів верхнього рівня, __init__.py) з
# однаковими іменами модулів/пакетів (constants, index, і, для
# generator_br_and_report_for_money/generator_timesheet, ще й content/
# generators/utils - обидва мають ОДНАКОВО названі підпакети). Коли кілька
# збираються в одному pytest-процесі, той, чий constants.py (чи
# content/generators/utils) імпортувався останнім, лишається закешованим у
# sys.modules - і bare `from constants import ...` (чи `import content...`) в
# ІНШОМУ проєкті підхопить ЧУЖИЙ модуль (ImportError: немає потрібного імені,
# чи гірше - тихо підхопить не той файл). Порядок збору між testpaths не
# гарантований, тож перевіряємо перед КОЖНИМ файлом/тестом, який саме проєкт
# зараз закешований, і за потреби перевантажуємо.
#
# Хуки живуть саме тут (в кореневому conftest.py), а не в conftest.py одного
# з проєктів: conftest.py-хуки активні лише для файлів у ВЛАСНОМУ піддереві -
# хук у projectA/tests/ ніколи не спрацював би для projectB/tests, і навпаки.
# Кореневий conftest.py - спільний предок усіх, тому бачить усе.
_PROJECT_DIRS = {
    "generator_br_and_report_for_money": _ROOT_DIR / "generator_br_and_report_for_money",
    "final_combat_report": _ROOT_DIR / "final_combat_report",
    "generator_timesheet": _ROOT_DIR / "generator_timesheet",
    "report_generator": _ROOT_DIR / "report_generator",
    "report_generator_old_new_state": _ROOT_DIR / "report_generator_old_new_state",
}
# "template" - report_generator і report_generator_old_new_state обидва мають
# пакет template/ з ОДНАКОВО названими підмодулями (template.change_position
# тощо), але НЕСУМІСНИМИ сигнатурами - без цього колізія дала б заплутаний
# TypeError замість чистого ImportError.
#
# helpers_test/accepted_position_test/change_position_test/handed_position_test -
# ЦЕ ВЖЕ НЕ внутрішні модулі проєкту, а самі ТЕСТОВІ файли (кілька "пласких"
# tests/ без __init__.py мають файл з ОДНАКОВОЮ базовою назвою:
# generator_br_and_report_for_money/report_generator/report_generator_old_new_state
# усі мають helpers_test.py; report_generator й report_generator_old_new_state
# додатково мають однаково названі accepted_position_test.py/
# change_position_test.py/handed_position_test.py). pytest імпортує тестові
# файли під їхньою "голою" базовою назвою так само, як і звичайні модулі
# (import_path/importtestmodule), тож без еквівалентної евікції в
# sys.modules він знаходить ЗАКЕШОВАНИЙ модуль під тим самим іменем від
# ІНШОГО проєкту і відмовляється збирати файл ("import file mismatch: ...").
# Ці записи не заважають production-колізіям вище - eviction тут лише прибирає
# застарілий кеш САМОГО тестового файлу, не зачіпаючи те, що він імпортує.
_COLLIDING_MODULES = [
    "constants", "helpers", "index", "content", "generators", "utils", "template",
    "helpers_test", "accepted_position_test", "change_position_test", "handed_position_test",
]


def _project_for_path(path):
    normalized = path.replace("\\", "/")
    return next((name for name in _PROJECT_DIRS if f"/{name}/" in normalized), None)


# Явно відстежує, на який проєкт ми ОСТАННІМ перемикались - НЕ виводиться з
# sys.modules.get("constants") (як робилось раніше через _loaded_project()).
# Той підхід ламався, коли файл проєкту, чиї тести саме збираються, сам не
# проходить збір через ІНШУ, не пов'язану з цим причину - напр. pytest-івську
# перевірку "import file mismatch" (кілька проєктів мають тестові файли з
# ОДНАКОВОЮ базовою назвою - accepted_position_test.py тощо - у "пласких"
# tests/ без __init__.py; сам збіг НЕ покривається _COLLIDING_MODULES, бо це
# назви ТЕСТОВИХ файлів, а не внутрішніх модулів проєкту). pytest у такому
# випадку відхиляє файл ще ДО виконання його коду (importlib повертає
# закешований модуль під тим самим "голим" ім'ям і виявляє невідповідний
# __file__) - "constants" усередині НІКОЛИ не імпортується, sys.modules
# лишається БЕЗ нього. _loaded_project() тоді бачила None й щоразу викликала
# _switch_to (нешкідливо, хоч і зайво) - АЛЕ щойно наступний файл ТОГО Ж
# проєкту вже МІГ успішно імпортувати "constants" одного разу, наступний файл
# з колізією знову лишав "constants" незайманим, і залежно від порядку файлів
# траплялось, що collectstart для файлу, який іде ПІСЛЯ вдалого імпорту
# "constants" ІНШОГО (вже нового) проєкту переставав викликати _switch_to
# ЗАВЧАСНО - "template" (чи інший collision-модуль) лишався закешованим від
# ПОПЕРЕДНЬОГО проєкту, і `from template.change_position import ...`
# підхоплював ЙОГО, а не проєкт, чиї файли зараз насправді збираються.
# Власна змінна тут не залежить від того, чи саме "constants" встиг
# (пере)імпортуватись - лише від того, який проєкт ми самі останнім явно
# вибрали.
_active_project = None


def _switch_to(target_project):
    global _active_project
    _active_project = target_project
    target_dir = _PROJECT_DIRS[target_project]
    other_dirs = [d for name, d in _PROJECT_DIRS.items() if name != target_project]

    # constants.py (і схожі модулі) у кожному проєкті читають resources/...
    # ВІДНОСНИМ шляхом, і роблять це вже на ІМПОРТІ (наприклад
    # PERSONEL_LIST_SHEET_NAME = json.load(open("resources/data.json"))["unit"][...]) -
    # тобто ще під час ЗБОРУ тестів (pytest_collectstart нижче), задовго до
    # того, як pytest_runtest_setup взагалі вперше зробить chdir для БУДЬ-ЯКОГО
    # тесту. Раніше chdir у момент ЗБОРУ забезпечували ОКРЕМІ pytest_sessionstart
    # хуки в generator_br_and_report_for_money/tests/conftest.py й
    # generator_timesheet/tests/conftest.py - обидва спрацьовують ОДИН РАЗ на
    # весь сеанс (не per-проєкт), тож який з двох виконався ОСТАННІМ, той cwd і
    # лишався на ввесь збір - для report_generator/report_generator_old_new_state
    # (жодного власного sessionstart-хука) constants.py під час збору міг
    # відкрити resources/data.json ЧУЖОГО проєкту (той, чий sessionstart
    # переміг) - і впасти з KeyError, бо там немає потрібного ключа (не
    # FileNotFoundError, бо ЯКИЙСЬ resources/data.json там таки був). Хук
    # нижче вже й так перемикається на ПРАВИЛЬНИЙ проєкт перед збором КОЖНОГО
    # його файлу - chdir тут, а не в окремих sessionstart-хуках, гарантує
    # правильний cwd з тієї ж миті, без залежності від порядку завантаження
    # conftest.py різних проєктів.
    os.chdir(target_dir)

    for prefix in _COLLIDING_MODULES:
        for mod_name in list(sys.modules):
            if mod_name == prefix or mod_name.startswith(prefix + "."):
                del sys.modules[mod_name]

    for p in (str(target_dir), *(str(d) for d in other_dirs)):
        if p in sys.path:
            sys.path.remove(p)
    sys.path.insert(0, str(target_dir))
    for d in other_dirs:
        sys.path.append(str(d))


def pytest_collectstart(collector):
    path = str(getattr(collector, "path", "") or getattr(collector, "fspath", ""))
    if not path.endswith(".py"):
        return
    project = _project_for_path(path)
    if project is not None and _active_project != project:
        _switch_to(project)


_ALL_LABEL = "Всі тести"


def _ask_projects():
    scope = _RealListPrompt(message="Які тести запустити?", choices=[_ALL_LABEL, "Обрати проєкти"]).execute()
    if scope == _ALL_LABEL:
        return None

    selected = _RealCheckboxPrompt(
        message="Оберіть проєкт(и) (Space → Enter):",
        choices=list(_PROJECT_DIRS),
        validate=lambda x: len(x) > 0,
        invalid_message="❌ Оберіть хоча б один проєкт!",
    ).execute()
    return set(selected)


def _prompt_selected_projects(config):
    """Питає, які проєкти тестувати - ЛИШЕ в інтерактивному терміналі.

    pytest ЗА ЗАМОВЧУВАННЯМ (--capture=fd) уже до виклику цього хука
    перенаправляє сам файловий дескриптор stdin на os.devnull (щоб тест, який
    випадково читає stdin, не завис) - тож звичайний sys.stdin.isatty() тут
    ЗАВЖДИ побачив би не-tty, навіть у справжньому інтерактивному терміналі.
    capturemanager (вбудований плагін pytest, що керує цим перенаправленням)
    вміє тимчасово це призупинити - suspend_global_capture(in_=True) поверне
    І sys.stdin, І сам fd 0 до оригінального терміналу на час запиту,
    resume_global_capture() відновить перенаправлення для самих тестів після.
    Без capturemanager (напр. `-p no:capture`/`-s`) stdin ніхто не займав,
    isatty() і так вірний одразу.

    Повертає None (усі тести - нічого фільтрувати не треба), або множину
    обраних назв проєктів (ключів _PROJECT_DIRS)."""
    capman = config.pluginmanager.getplugin("capturemanager")
    if capman is not None:
        capman.suspend_global_capture(in_=True)
    try:
        if not sys.stdin.isatty():
            return None
        return _ask_projects()
    except (Exception, KeyboardInterrupt) as error:
        # isatty() == True не гарантує, що prompt_toolkit тут ЗМОЖЕ намалювати
        # інтерактивний список - напр. git-bash/mintty на Windows видає
        # справжній tty, але без Win32 screen buffer, якого потребує
        # prompt_toolkit (NoConsoleScreenBufferError), а деякі середовища (напр.
        # пісочниця, де запускались автотести цього фікса) видають tty, який на
        # спробу прочитати ввід від prompt_toolkit відповідає КЕРУЮЧИМ
        # KeyboardInterrupt - а не звичайним Exception. Запит - зручність, не
        # критична частина запуску, тож будь-яка помилка тут (не лише ця
        # конкретна, і не лише Exception) падає назад на "усі тести", а не
        # валить увесь прогін. Сам print теж може впасти (UnicodeEncodeError,
        # якщо консоль не UTF-8, а в повідомленні є "⚠" чи нестандартні
        # символи з repr(error)) - обгорнуто в свій try/except, щоб ЦЕ теж не
        # звалило увесь прогін замість запланованого "запускаю всі тести".
        try:
            print(f"\n⚠ Не вдалося показати інтерактивний вибір проєктів ({error!r}) - запускаю всі тести.\n")
        except Exception:
            pass
        return None
    finally:
        if capman is not None:
            capman.resume_global_capture()
        # prompt_toolkit (усередині _ask_projects вище) сам малює через Win32
        # Console API і, як побічний ефект, скидає кодову сторінку консолі,
        # виставлену на самому старті цього файлу - ствердити ще раз ТУТ,
        # після запиту, інакше подальший звіт pytest знову ламається.
        _fix_windows_console_codepage()


def pytest_configure(config):
    config._selected_projects = _prompt_selected_projects(config)
    # Часткова добірка (не всі три проєкти) - загальний поріг покриття
    # (--cov-fail-under=80 у pytest.ini) рахується по ВСІХ трьох і був би
    # хибно незадоволений, коли виконано тести лише частини з них - звіт
    # покриття все одно друкується, лишень не провалює запуск через це.
    #
    # pytest-cov читає поріг НЕ з config.option (мутація якого тут запізно на
    # це не впливає), а зі СВОЄЇ окремої копії options, знятої ЩЕ РАНІШЕ, у
    # власному pytest_load_initial_conftests (до того, як цей хук взагалі
    # встиг спрацювати) - тож міняти треба САМЕ цю копію, через сам
    # зареєстрований плагін ("_cov" - ім'я, під яким pytest-cov сам себе
    # реєструє).
    if config._selected_projects is not None and len(config._selected_projects) < len(_PROJECT_DIRS):
        cov_plugin = config.pluginmanager.getplugin("_cov")
        if cov_plugin is not None:
            cov_plugin.options.cov_fail_under = 0


def pytest_ignore_collect(collection_path, config):
    # Пропускає збір НЕобраних проєктів узагалі (не лише прибирає їхні тести
    # з результату) - економить час і, для generator_br_and_report_for_money,
    # уникає зайвого імпорту його constants.py (там власний інтерактивний
    # запит місяця) для проєкту, який і так не буде запущено.
    selected = getattr(config, "_selected_projects", None)
    if selected is None:
        return None
    project = _project_for_path(str(collection_path))
    return True if project is not None and project not in selected else None


def pytest_collection_modifyitems(config, items):
    # Запасний фільтр ПІСЛЯ збору - на випадок, якщо щось із НЕобраного
    # проєкту все ж потрапило в items (pytest_ignore_collect мав би відсіяти
    # це раніше, але явний позиційний testpath з pytest.ini addopts - надійніший
    # гарант, ніж покладатись лише на один хук).
    selected = getattr(config, "_selected_projects", None)
    if selected is None:
        return

    kept, removed = [], []
    for item in items:
        project = _project_for_path(str(item.fspath))
        (kept if project in selected else removed).append(item)

    items[:] = kept
    if removed:
        config.hook.pytest_deselected(items=removed)


def pytest_runtest_setup(item):
    # Кожен проєкт читає resources/... відносно власної теки - cwd перед
    # кожним тестом виставляємо залежно від того, з якого проєкту цей тест
    # (одноразового chdir на весь сеанс не досить, коли тести кількох проєктів
    # чергуються в одному прогоні).
    #
    # Так само й закешовані collision-модулі потрібно перевіряти перед КОЖНИМ
    # тестом, а не лише при колекції: тест може імпортувати модуль свого
    # проєкту вперше вже під час виконання (а не збирання), і на той момент
    # sys.modules["constants"] могла лишитись від того проєкту, чиї файли
    # збирались останніми.
    path = str(item.fspath)
    project = _project_for_path(path)
    if project is not None:
        os.chdir(_PROJECT_DIRS[project])
        if _active_project != project:
            _switch_to(project)
