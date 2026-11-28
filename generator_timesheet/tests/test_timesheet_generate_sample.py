import importlib.util
from datetime import date
from pathlib import Path

import constants
from content.oblik_timesheet import Timesheet
from utils import project_settings

_PROJECT_DIR = Path(__file__).resolve().parent.parent


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _sample_module():
    return _load("timesheet_generate_sample", _PROJECT_DIR / "resources_example" / "generate_sample.py")


def test_generate_sample_creates_readable_resources_once(tmp_path, capsys):
    module = _sample_module()
    target = tmp_path / "resources"

    module.main(str(target))

    assert sorted(path.name for path in target.iterdir()) == sorted(["НАЛАШТУВАННЯ.xlsx", "ОБЛІК.xlsx"])
    # Статусів зразок не задає - у книзі стандартні (utils/default_statuses.py).
    assert project_settings.read_settings(str(target / "НАЛАШТУВАННЯ.xlsx")) == project_settings.with_default_statuses(
        module.settings_example()
    )
    assert len(Timesheet(str(target / "ОБЛІК.xlsx"), "Табель").person_rows()) == len(module.PEOPLE)

    settings_bytes = (target / "НАЛАШТУВАННЯ.xlsx").read_bytes()
    module.main(str(target))

    assert (target / "НАЛАШТУВАННЯ.xlsx").read_bytes() == settings_bytes
    assert capsys.readouterr().out.count("вже існує") == 2


def test_sample_settings_have_everything_constants_needs(tmp_path, monkeypatch):
    """constants.py, виконаний зі зразком замість resources/НАЛАШТУВАННЯ.xlsx,
    не падає (жодного пропущеного параметра), і кожен список читається."""
    # Свіжий імпорт, а не модуль з початку файлу: у спільній сесії кількох
    # проєктів кореневий conftest.py вивантажує utils.* перед тестом, і
    # constants.py нижче імпортує саме поточний модуль.
    import utils.project_settings as current_settings

    module = _sample_module()
    module.main(str(tmp_path))
    sample = current_settings.read_settings(str(tmp_path / "НАЛАШТУВАННЯ.xlsx"))
    monkeypatch.setattr(current_settings, "load_settings", lambda *_args, **_kwargs: sample)

    sample_constants = _load("timesheet_constants_from_sample", constants.__file__)

    first_name, second_name = module.PEOPLE[0][3], module.PEOPLE[1][3]
    first_day = date.today().replace(day=1)
    assert sample_constants.FULL_MILITARY_UNIT == "військової частини А0000"
    assert list(sample["IGNORE_LISTS"]) == list(constants.ADDON_IGNORE_LIST_KEYS)
    assert sample_constants.ADDON_MISMATCH_IGNORE_LIST == (
        ("treatment", first_name, first_day, first_day.replace(day=2)),
    )
    assert sample_constants.PAYMENT_CATEGORY_100_PEOPLE == (
        {"name": second_name, "days": 31, "color": constants.PAYMENT_CATEGORY_100_HIGHLIGHT_COLOR},
    )
    assert sample_constants.REVIEW_EXCLUDE_PAIRS == ((first_day, first_name),)
    assert sample_constants.PAYMENTS_REVIEW_EXCLUDE_PAIRS == ((first_day, first_name),)
    assert sample_constants.PAYMENT_ABSENT_FROM_FILES_EXCLUDE_NAMES == (second_name,)
    assert sample_constants.BCHS_EXPECTED_SUBDIVISIONS == ("Підрозділ 1", "Підрозділ 2")
    assert sample_constants.BRIGADE_MILITARY_UNIT_CODES == {"0000", "0001"}
    assert sample_constants.STATUS_COLORS == constants.STATUS_COLORS  # стандартні статуси з аркуша "Статуси"
    assert sample_constants.PAYMENT_STATUS_OVERRIDES == constants.PAYMENT_STATUS_OVERRIDES
