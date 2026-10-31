from datetime import date, datetime, time

from book61_fixtures import (
    ALF2_HEADER,
    PPD_HEADER,
    RREZ_HEADER,
    SHP_HEADER,
    SZCH_HEADER,
    VD_HEADER,
    VLK_HEADER,
    VP_HEADER,
    VTRATY_HEADER,
    write_book61,
    write_full_book61,
)

import content.addon_reader as ar


def test_implausible_future_dates_are_likely_year_typos():
    """Реальні випадки: виписка "05.07.2027" при лікуванні з 30.06.2026,
    відпустка до "03.12.2028" - далеко в майбутньому; відпустка, що
    закінчиться через місяць-два, - ні."""
    today = date(2026, 10, 26)
    assert ar.is_implausible_future_date(date(2027, 7, 5), today)
    assert not ar.is_implausible_future_date(date(2026, 12, 5), today)
    assert not ar.is_implausible_future_date(None, today)
    typo = {"start_date": date(2026, 6, 30), "end_date": date(2027, 7, 5)}
    open_period = {"start_date": date(2026, 10, 1), "end_date": None}
    leave = {"start_date": date(2026, 10, 7), "end_date": date(2026, 12, 5)}
    assert ar.implausible_date_records([typo, open_period, leave], today) == [typo]


def test_is_book61_workbook_requires_every_sheet_the_readers_open(tmp_path):
    assert ar.is_book61_workbook(ar.load_addon_workbook(write_full_book61(tmp_path / "full.xlsm", {})))
    partial = write_book61(tmp_path / "partial.xlsm", {"ВП": (VP_HEADER, []), "Алф2": (ALF2_HEADER, [])})
    assert not ar.is_book61_workbook(ar.load_addon_workbook(partial))


def test_as_date_handles_datetime_date_and_other_values():
    assert ar._as_date(datetime(2026, 8, 1, 12, 30)) == date(2026, 8, 1)
    assert ar._as_date(date(2026, 8, 1)) == date(2026, 8, 1)
    assert ar._as_date(None) is None
    assert ar._as_date("08.01.2026") is None


def test_order_date_from_text_extracts_the_embedded_date():
    """Реальний випадок (book61.xlsm, аркуш Алф2): колонка "Дата, наказ
    розпорядження, виключення" - ВІЛЬНИЙ текст із вбудованою датою, а не
    чиста дата-комірка."""
    assert ar._order_date_from_text("Виключений із списків 31.10.2023") == date(2023, 10, 31)
    assert ar._order_date_from_text("щось без дати") is None
    assert ar._order_date_from_text(None) is None
    assert ar._order_date_from_text(datetime(2026, 8, 1)) is None


def test_end_date_from_cell_prefers_a_plain_date():
    assert ar._end_date_from_cell(date(2026, 7, 18)) == date(2026, 7, 18)
    assert ar._end_date_from_cell(datetime(2026, 7, 18, 9, 0)) == date(2026, 7, 18)


def test_end_date_from_cell_extracts_a_single_embedded_date():
    """Реальний випадок (book61.xlsm, аркуш ВД): "Дата прибуття" - вільний
    текст "у ППД з 18.07.2026", а не чиста дата-комірка."""
    assert ar._end_date_from_cell("у ППД з 18.07.2026") == date(2026, 7, 18)
    assert ar._end_date_from_cell("у СЗЧ з 09.03.2026") == date(2026, 3, 9)


def test_end_date_from_cell_prefers_the_last_embedded_date():
    """Реальний випадок (book61.xlsm, аркуш ШП): "надходження 27.06.2026,
    виписка 15.07.2026" - релевантна (кінець періоду) дата - ДРУГА
    (виписка), а не перша (надходження)."""
    assert ar._end_date_from_cell("надходження 27.06.2026, виписка 15.07.2026") == date(2026, 7, 15)


def test_end_date_from_cell_returns_none_without_any_embedded_date():
    """Реальні випадки (book61.xlsm): "продовжено"/"до окремого
    розпорядження"/"ТВО"/"-" - дійсно немає конкретної дати завершення -
    period МАЄ лишитись відкритим (None), а не хибно щось вигадувати."""
    assert ar._end_date_from_cell("продовжено") is None
    assert ar._end_date_from_cell("до окремого розпорядження") is None
    assert ar._end_date_from_cell("-") is None
    assert ar._end_date_from_cell(None) is None


# --- read_vp_family_periods -------------------------------------------------


def test_read_vp_family_periods_resolves_status_from_type_column(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"ВП": (VP_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "від", date(2026, 7, 1), date(2026, 7, 20), date(2026, 7, 19),
         "Перша частина щорічної відпустки"],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "від", date(2026, 7, 1), date(2026, 7, 20), date(2026, 7, 19),
         "За станом здоров'я"],
        [3, "Підрозділ1", "матрос", "ТРЕТІЙ Третій Третій", "від", date(2026, 7, 1), date(2026, 7, 20), date(2026, 7, 19),
         "За сімейними обставинами"],
        [4, "Підрозділ1", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "від", date(2026, 7, 1), date(2026, 7, 20), date(2026, 7, 19),
         "Соціальна відпустка"],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vp_family_periods(wb)

    statuses = {r["pib_raw"]: r["status"] for r in records}
    assert statuses == {
        "ПЕРШИЙ Перший Перший": "ВП",
        "ДРУГИЙ Другий Другий": "ВПСЗ",
        "ТРЕТІЙ Третій Третій": "ВПС",
        "ЧЕТВЕРТИЙ Четвертий Четвертий": "ВП",  # Соціальна відпустка -> звичайний ВП
    }
    first = records[0]
    assert first["start_date"] == date(2026, 7, 1)
    assert first["end_date"] == date(2026, 7, 18)  # "Дата прибуття" 19.07 - 1 день
    assert first["source_sheet"] == "ВП"


def test_read_vp_family_periods_falls_back_to_08_ranku_minus_one_day_when_arrival_date_missing(tmp_path):
    """Рішення користувача: "Дата 08 ранку" - ранок, коли
    людина вже на місці, тож останній день відпустки - на день раніше."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВП": (VP_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "від", date(2026, 7, 1), date(2026, 7, 20), None,
         "Перша частина щорічної відпустки"],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vp_family_periods(wb)

    assert records[0]["end_date"] == date(2026, 7, 19)


def test_read_vd_periods_keeps_the_book61_end_date_as_is(tmp_path):
    """Зсув будь-якої дати кінця на день - лише для аркуша ВП (рішення користувача); на ВД чиста дата лишається як є."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВД": (VD_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 7, 1), date(2026, 7, 20), None, "10"],
    ])})
    wb = ar.load_addon_workbook(path)

    assert ar.read_vd_periods(wb)[0]["end_date"] == date(2026, 7, 20)


def test_read_vp_family_periods_skips_rows_without_a_pib_or_start_date(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"ВП": (VP_HEADER, [
        [1, "Підрозділ1", "матрос", None, "від", date(2026, 7, 1), None, None, "Перша частина щорічної відпустки"],
        [2, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "від", None, None, None, "Перша частина щорічної відпустки"],
    ])})
    wb = ar.load_addon_workbook(path)

    assert ar.read_vp_family_periods(wb) == []


# --- read_vd_periods ---------------------------------------------------------


def test_read_vd_periods_distinguishes_vd_from_navchannya_by_payment_column(tmp_path):
    """Реальний випадок (map рядки 12+15): "Виплата" = "10" -> дійсно "ВД";
    "навчання" -> ОБЛІК МАЄ казати "Навч", не "ВД"."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВД": (VD_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 7, 1), None, date(2026, 7, 10), "10"],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "вдр", date(2026, 7, 1), None, date(2026, 7, 10), "навчання"],
        [3, "Підрозділ1", "матрос", "ТРЕТІЙ Третій Третій", "вдр", date(2026, 7, 1), None, date(2026, 7, 10), None],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vd_periods(wb)

    statuses = {r["pib_raw"]: r["status"] for r in records}
    assert statuses == {
        "ПЕРШИЙ Перший Перший": "ВД",
        "ДРУГИЙ Другий Другий": "Навч",
        "ТРЕТІЙ Третій Третій": "ВД",  # порожня "Виплата" - за замовчуванням ВД
    }


def test_read_vd_periods_takes_the_status_from_the_tabel_when_not_departed(tmp_path):
    """Рішення користувача: "Виплата" = "не вибув" (чи "не
    вибула"): людина фактично не вибула, статус бере Табель (ANY_STATUS)."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВД": (VD_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 10, 1), None, date(2026, 10, 10), "не вибув"],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "вдр", date(2026, 10, 1), None, date(2026, 10, 10), " Не  вибула "],
        [3, "Підрозділ1", "матрос", "ТРЕТІЙ Третій Третій", "вдр", date(2026, 10, 1), None, date(2026, 10, 10), "10"],
    ])})
    wb = ar.load_addon_workbook(path)

    statuses = {r["pib_raw"]: r["status"] for r in ar.read_vd_periods(wb)}

    assert statuses == {
        "ПЕРШИЙ Перший Перший": ar.ANY_STATUS,
        "ДРУГИЙ Другий Другий": ar.ANY_STATUS,
        "ТРЕТІЙ Третій Третій": "ВД",
    }


def test_read_vd_periods_closes_the_period_on_a_free_text_arrival_date(tmp_path):
    """Реальний випадок: "Дата прибуття" - вільний текст "у ППД з
    18.07.2026", а не чиста дата-комірка. Період МАЄ закритись, а не
    лишитись відкритим (інакше хибно перевірявся б аж до останньої
    колонки-дати ПОТОЧНОГО Табеля) - 17.07: з 18.07 людина вже на ППД."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВД": (VD_HEADER, [
        [1, "Підрозділ1", "лейтенант", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 7, 12), None,
         "у ППД з 18.07.2026", None],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vd_periods(wb)

    assert records[0]["end_date"] == date(2026, 7, 17)


def test_arrival_date_is_set_only_for_an_unshifted_end_date(tmp_path):
    """День прибуття - лише коли кінець записано датою: для тексту "у
    <статус> з <дата>" (перехід в інший статус) - None."""
    path = write_book61(tmp_path / "book61.xlsm", {"ВД": (VD_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "вдр", date(2026, 10, 1), None, date(2026, 10, 5), "10"],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "вдр", date(2026, 10, 1), None, "у лік з 05.10.2026", "10"],
    ])})
    wb = ar.load_addon_workbook(path)

    by_pib = {r["pib_raw"]: r for r in ar.read_vd_periods(wb)}

    assert by_pib["ПЕРШИЙ Перший Перший"]["arrival_date"] == date(2026, 10, 5)
    assert by_pib["ДРУГИЙ Другий Другий"]["arrival_date"] is None


def test_read_ppd_periods_ends_the_day_before_the_departure_date(tmp_path):
    """Реальний випадок: "Дата вибуття з ППД" 09.10 - у цей день Табель уже
    показує новий статус, тож ППД - до 08.10."""
    path = write_book61(tmp_path / "book61.xlsm", {"ППД": (PPD_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 8), date(2026, 10, 9)],
    ])})
    wb = ar.load_addon_workbook(path)

    assert ar.read_ppd_periods(wb)[0]["end_date"] == date(2026, 10, 8)


def test_free_text_transition_end_closes_the_period_the_day_before(tmp_path):
    """Реальний випадок: ППД "Дата вибуття з ППД" - "у вдр з 09.10.2026" -
    09.10 уже перший день відрядження, тож ППД - до 08.10. ШП-текст
    "надходження ..., виписка ..." - не перехід, без зсуву; ВП (і так
    зсувається на день) - не зсувається вдруге."""
    path = write_book61(tmp_path / "book61.xlsm", {
        "ППД": (PPD_HEADER, [[1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 10, 1),
                              "у вдр з 09.10.2026"]]),
        "ШП": (SHP_HEADER, [[1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "лік", date(2026, 6, 27),
                             "надходження 27.06.2026, виписка 15.07.2026"]]),
        "ВП": (VP_HEADER, [[1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "від", date(2026, 7, 1), None,
                            "у лік з 17.07.2026", "Перша частина щорічної відпустки"]]),
    })
    wb = ar.load_addon_workbook(path)

    assert ar.read_ppd_periods(wb)[0]["end_date"] == date(2026, 10, 8)
    assert ar.read_shp_family_periods(wb)[0]["end_date"] == date(2026, 7, 15)
    assert ar.read_vp_family_periods(wb)[0]["end_date"] == date(2026, 7, 16)


# --- read_shp_family_periods -------------------------------------------------


def test_read_shp_family_periods_splits_shp_and_bpshp_by_bchs_column(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"ШП": (SHP_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "лік", date(2026, 7, 1), date(2026, 7, 10)],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "ЛП", date(2026, 7, 1), None],
        [3, "Підрозділ1", "матрос", "ТРЕТІЙ Третій Третій", "щось інше", date(2026, 7, 1), date(2026, 7, 10)],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_shp_family_periods(wb)

    statuses = {r["pib_raw"]: r["status"] for r in records}
    assert statuses == {"ПЕРШИЙ Перший Перший": "ШП", "ДРУГИЙ Другий Другий": "БПШП"}
    bpshp = next(r for r in records if r["pib_raw"] == "ДРУГИЙ Другий Другий")
    assert bpshp["end_date"] is None  # відкритий період - "Кінець (виписка)" відсутня


def test_read_shp_family_periods_treatment_in_own_unit_is_lmr(monkeypatch, tmp_path):
    """Рішення користувача: "Назва госпіталя чи в/ч" з кодом
    власної в/ч (реально - "Медична рота військової частини ...") - "ЛМР",
    як у книзі ЦВС, незалежно від БЧС; інша в/ч чи госпіталь - як раніше;
    рядок з іншим БЧС - пропускається й далі."""
    monkeypatch.setattr(ar, "OWN_MILITARY_UNIT_CODES", frozenset({"0000"}))
    header = SHP_HEADER + ["Назва госпіталя чи в/ч"]
    path = write_book61(tmp_path / "book61.xlsm", {"ШП": (header, [
        [1, "П1", "матрос", "ПЕРШИЙ Перший Перший", "лік", date(2026, 7, 1), date(2026, 7, 10),
         "Медична рота військової частини А0000"],
        [2, "П1", "матрос", "ДРУГИЙ Другий Другий", "ЛП", date(2026, 7, 1), None, "Медична рота в/ч A 0000"],
        [3, "П1", "матрос", "ТРЕТІЙ Третій Третій", "лік", date(2026, 7, 1), None, "Медична рота в/ч А0001"],
        [4, "П1", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "ЛП", date(2026, 7, 1), None, "Госпіталь"],
        [5, "П1", "матрос", "П'ЯТИЙ П'ятий П'ятий", "щось інше", date(2026, 7, 1), None, "в/ч А0000"],
    ])})

    records = ar.read_shp_family_periods(ar.load_addon_workbook(path))

    assert {r["pib_raw"]: r["status"] for r in records} == {
        "ПЕРШИЙ Перший Перший": "ЛМР", "ДРУГИЙ Другий Другий": "ЛМР",
        "ТРЕТІЙ Третій Третій": "ШП", "ЧЕТВЕРТИЙ Четвертий Четвертий": "БПШП",
    }


# --- read_vlk_periods / read_ppd_periods -------------------------------------

def test_read_vlk_periods(tmp_path):
    header = VLK_HEADER
    path = write_book61(tmp_path / "book61.xlsm", {"ВЛК": (header, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "ВЛК", date(2026, 7, 1), date(2026, 7, 3)],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vlk_periods(wb)

    assert records == [{
        "pib_raw": "ПЕРШИЙ Перший Перший", "status": "ВЛК", "start_date": date(2026, 7, 1),
        "end_date": date(2026, 7, 3), "source_sheet": "ВЛК", "arrival_date": date(2026, 7, 3),
    }]


def test_read_ppd_periods(tmp_path):
    header = PPD_HEADER
    path = write_book61(tmp_path / "book61.xlsm", {"ППД": (header, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "ППД", date(2026, 7, 1), None],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_ppd_periods(wb)

    assert records == [{
        "pib_raw": "ПЕРШИЙ Перший Перший", "status": "ППД", "start_date": date(2026, 7, 1),
        "end_date": None, "source_sheet": "ППД", "arrival_date": None,
    }]


# --- read_alf2_order_periods / read_unmapped_alf2_people ---------------------


def test_read_alf2_order_periods_maps_known_statuses_and_skips_unknown(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Алф2": (ALF2_HEADER, [
        [1, "ПЕРШИЙ Перший Перший", "звільнений", None, None, "Виключений із списків 31.10.2023"],
        [2, "ДРУГИЙ Другий Другий", "переведений в іншу в/ч", None, None, "Виключений із списків 01.11.2023"],
        [3, "ТРЕТІЙ Третій Третій", "переведений в середині бригади", None, None, "Виключений із списків 02.11.2023"],
        [4, "ЧЕТВЕРТИЙ Четвертий Четвертий", "полон", None, None, "Виключений із списків 03.11.2023"],
        [5, "ДЕСЯТИЙ Десятий Десятий", "самостріл", None, None, "Виключений із списків 04.11.2023"],
        [6, "ШОСТИЙ Шостий Шостий", "звільнений", None, None, "немає дати тут взагалі"],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_alf2_order_periods(wb)

    by_pib = {r["pib_raw"]: r for r in records}
    assert by_pib["ПЕРШИЙ Перший Перший"]["status"] == "Звіль"
    # Рішення користувача: замість "ПРВД" - два статуси: інша в/ч - ПРВДВЧ, бригада - ПРВДБ.
    assert by_pib["ДРУГИЙ Другий Другий"]["status"] == "ПРВДВЧ"
    assert by_pib["ТРЕТІЙ Третій Третій"]["status"] == "ПРВДБ"
    assert by_pib["ЧЕТВЕРТИЙ Четвертий Четвертий"]["status"] == "полон"
    assert "ДЕСЯТИЙ Десятий Десятий" not in by_pib  # самостріл - не зіставлено з жодним статусом
    assert "ШОСТИЙ Шостий Шостий" not in by_pib  # без дати в тексті наказу
    assert by_pib["ПЕРШИЙ Перший Перший"]["start_date"] == date(2023, 10, 31)
    assert by_pib["ПЕРШИЙ Перший Перший"]["end_date"] is None


def test_read_alf2_event_statuses_start_on_the_from_date(tmp_path):
    """Рішення користувача (реальний випадок: люди з "БЗ" у
    Табелі є лише на Алф2): "безвісті" - БЗ, "загибель" - БЗ або Загиб (без
    свідоцтва), "полон" - полон; початок - колонка "з" (наказ здебільшого
    порожній), немає - дата наказу; кінець - "по", якщо є."""
    path = write_book61(tmp_path / "book61.xlsm", {"Алф2": (ALF2_HEADER, [
        [1, "ПЕРШИЙ Перший Перший", "безвісті", date(2024, 8, 1), None, None],
        [2, "ДРУГИЙ Другий Другий", "загибель", date(2024, 7, 5), None, "Наказ про виключення із списків №1 від 26.07.2024"],
        [3, "ТРЕТІЙ Третій Третій", "полон", date(2024, 8, 13), date(2026, 10, 5), None],
        [4, "ЧЕТВЕРТИЙ Четвертий Четвертий", "безвісті", None, None, "Виключений із списків 03.11.2023"],
        [5, "П'ЯТИЙ П'ятий П'ятий", "безвісті", None, None, None],
    ])})

    records = ar.read_alf2_order_periods(ar.load_addon_workbook(path))

    assert [(r["pib_raw"], r["status"], r["start_date"], r["end_date"]) for r in records] == [
        ("ПЕРШИЙ Перший Перший", "БЗ", date(2024, 8, 1), None),
        ("ДРУГИЙ Другий Другий", "БЗ/Загиб", date(2024, 7, 5), None),  # без свідоцтва - і БЗ
        ("ТРЕТІЙ Третій Третій", "полон", date(2024, 8, 13), date(2026, 10, 5)),
        ("ЧЕТВЕРТИЙ Четвертий Четвертий", "БЗ", date(2023, 11, 3), None),
    ]


def test_read_unmapped_alf2_people_finds_only_samostril(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Алф2": (ALF2_HEADER, [
        [1, "ПЕРШИЙ Перший Перший", "звільнений", None, None, "Виключений із списків 31.10.2023"],
        [2, "ДРУГИЙ Другий Другий", "самостріл", None, None, None],
    ])})
    wb = ar.load_addon_workbook(path)

    people = ar.read_unmapped_alf2_people(wb)

    assert people == [{"pib_raw": "ДРУГИЙ Другий Другий", "alf2_status": "самостріл"}]


# --- read_szch_to_rozp_periods -----------------------------------------------


def test_read_szch_to_rozp_periods_ends_szch_the_day_before_the_return_date(tmp_path):
    """Реальний випадок, підтверджений користувачем: СЗЧ 28.02.2026,
    "Дата повернення з СЗЧ" 01.03.2026 - СЗЧ лише 28.02 (раніше - відкритий
    період аж у жовтень); без дати повернення - відкритий."""
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (SZCH_HEADER, [
            [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 2, 28), date(2026, 3, 1)],
            [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "СЗЧ", date(2026, 10, 6), None],
        ]),
        "Алф2": (ALF2_HEADER, []),
    })
    wb = ar.load_addon_workbook(path)

    by_pib = {r["pib_raw"]: r for r in ar.read_szch_to_rozp_periods(wb)}

    assert by_pib["ПЕРШИЙ Перший Перший"]["end_date"] == date(2026, 2, 28)
    assert by_pib["ДРУГИЙ Другий Другий"]["end_date"] is None


def test_read_szch_to_rozp_periods_adds_rozp_only_when_alf2_has_an_order_date(tmp_path):
    """Реальний випадок (map рядки 2+5): "СЗЧ" - завжди відкритий період;
    "Розп" - ДРУГИЙ, окремий запис, лише якщо ТА САМА людина ТЕЖ має
    рядок на Алф2 з непорожньою датою наказу."""
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (SZCH_HEADER, [
            [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 7, 1)],
            [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "СЗЧ", date(2026, 7, 5)],
        ]),
        "Алф2": (ALF2_HEADER, [
            [1, "ПЕРШИЙ Перший Перший", "СЗЧ", None, None, "Виключений із списків 10.07.2026"],
        ]),
    })
    wb = ar.load_addon_workbook(path)

    records = ar.read_szch_to_rozp_periods(wb)

    by_pib_and_sheet = {(r["pib_raw"], r["source_sheet"]): r for r in records}
    assert by_pib_and_sheet[("ПЕРШИЙ Перший Перший", "СЗЧ")]["status"] == "СЗЧ"
    assert by_pib_and_sheet[("ПЕРШИЙ Перший Перший", "Алф2 (Розп)")]["status"] == "Розп"
    assert by_pib_and_sheet[("ПЕРШИЙ Перший Перший", "Алф2 (Розп)")]["start_date"] == date(2026, 7, 10)
    assert ("ДРУГИЙ Другий Другий", "Алф2 (Розп)") not in by_pib_and_sheet  # немає рядка на Алф2 взагалі


# --- read_vtraty_periods ------------------------------------------------------


def test_read_vtraty_periods_maps_bchs_to_status(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Втрати": (VTRATY_HEADER, [
        [1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "загибель", date(2026, 7, 1)],
        [2, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "смерть", date(2026, 7, 2)],
        [3, "Підрозділ1", "матрос", "ТРЕТІЙ Третій Третій", "безвісті", date(2026, 7, 3)],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_vtraty_periods(wb)

    statuses = {r["pib_raw"]: r["status"] for r in records}
    assert statuses == {
        "ПЕРШИЙ Перший Перший": "Загиб",
        "ДРУГИЙ Другий Другий": "Смерть",
        "ТРЕТІЙ Третій Третій": "БЗ",
    }
    assert all(r["end_date"] is None for r in records)


# --- read_rrez_tvo_periods ------------------------------------------------------


def _steps(records):
    return [(r["status"], r["start_date"], r["end_date"]) for r in records]


def test_read_rrez_tvo_periods_builds_the_arrival_sequence(tmp_path):
    """Рішення користувача (реальний випадок): ТВО 22.09,
    зарахування в списки 05.10, прийом посади 07.10 - "РРез" до 04.10,
    "ППД" 05.10-06.10 (у день зарахування - лише ППД), 07.10 - будь-який
    статус."""
    path = write_book61(tmp_path / "book61.xlsm", {"Прибулі у ВД": (RREZ_HEADER, [
        [1, "idx1", "Підрозділ1", "Посада1", "матрос", "ПЕРШИЙ Перший Перший",
         date(2026, 9, 22), date(2026, 10, 5), date(2026, 10, 7)],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_rrez_tvo_periods(wb)

    assert _steps(records) == [
        ("РРез", date(2026, 9, 22), date(2026, 10, 4)),
        ("ППД", date(2026, 10, 5), date(2026, 10, 6)),
        (ar.ANY_STATUS, date(2026, 10, 7), date(2026, 10, 7)),
    ]
    assert {r["source_sheet"] for r in records} == {"Прибулі у ВД"}


def test_read_rrez_tvo_periods_handles_enrollment_and_position_on_the_same_day(tmp_path):
    """Реальні випадки: зарахування = прийом посади (і навіть = ТВО) - "РРез"
    до дня перед ним, а в сам день - і "ППД", і будь-який статус (у
    content/addon_mismatch_checker.py пізніший "будь-який" перекриває)."""
    path = write_book61(tmp_path / "book61.xlsm", {"Прибулі у ВД": (RREZ_HEADER, [
        [1, "idx1", "Підрозділ1", "Посада1", "матрос", "ПЕРШИЙ Перший Перший",
         date(2026, 9, 26), date(2026, 10, 9), date(2026, 10, 9)],
        [2, "idx2", "Підрозділ1", "Посада1", "матрос", "ДРУГИЙ Другий Другий",
         date(2026, 10, 8), date(2026, 10, 8), date(2026, 10, 8)],
    ])})
    wb = ar.load_addon_workbook(path)

    records = ar.read_rrez_tvo_periods(wb)

    assert _steps([r for r in records if r["pib_raw"].startswith("ПЕРШИЙ")]) == [
        ("РРез", date(2026, 9, 26), date(2026, 10, 8)),
        ("ППД", date(2026, 10, 9), date(2026, 10, 9)),
        (ar.ANY_STATUS, date(2026, 10, 9), date(2026, 10, 9)),
    ]
    assert _steps([r for r in records if r["pib_raw"].startswith("ДРУГИЙ")]) == [
        ("ППД", date(2026, 10, 8), date(2026, 10, 8)),
        (ar.ANY_STATUS, date(2026, 10, 8), date(2026, 10, 8)),
    ]


def test_read_rrez_tvo_periods_without_an_enrollment_date_keeps_rrez_until_the_position(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Прибулі у ВД": (RREZ_HEADER, [
        [1, "idx1", "Підрозділ1", "Посада1", "матрос", "ПЕРШИЙ Перший Перший",
         date(2026, 9, 22), None, date(2026, 10, 7)],
    ])})
    wb = ar.load_addon_workbook(path)

    assert _steps(ar.read_rrez_tvo_periods(wb)) == [
        ("РРез", date(2026, 9, 22), date(2026, 10, 6)),
        (ar.ANY_STATUS, date(2026, 10, 7), date(2026, 10, 7)),
    ]


def test_read_rrez_tvo_periods_without_an_accepted_position_leaves_the_last_step_open(tmp_path):
    """Рішення користувача (реальний випадок): ТВО 07.10, зарахування 10.10,
    "Дата прийому посади" ще не встановлена (порожня, 0 - openpyxl читає
    00:00, чи заглушка 01.01.1990) - "РРез" 07.10-09.10 і відкритий "ППД" з
    10.10. Без зарахування - відкритий "РРез"."""
    path = write_book61(tmp_path / "book61.xlsm", {"Прибулі у ВД": (RREZ_HEADER, [
        [1, "idx1", "Підрозділ1", "Посада1", "матрос", "ПЕРШИЙ Перший Перший",
         date(2026, 10, 7), date(2026, 10, 10), None],
        [2, "idx2", "Підрозділ1", "Посада1", "матрос", "ДРУГИЙ Другий Другий",
         date(2026, 10, 7), date(2026, 10, 10), time(0, 0)],
        [3, "idx3", "Підрозділ1", "Посада1", "матрос", "ТРЕТІЙ Третій Третій",
         date(2026, 10, 7), date(2026, 10, 10), date(1990, 1, 1)],
        [4, "idx4", "Підрозділ1", "Посада1", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий",
         date(2026, 10, 7), None, None],
    ])})

    records = ar.read_rrez_tvo_periods(ar.load_addon_workbook(path))

    open_ppd = [("РРез", date(2026, 10, 7), date(2026, 10, 9)), ("ППД", date(2026, 10, 10), None)]
    for name in ("ПЕРШИЙ", "ДРУГИЙ", "ТРЕТІЙ"):
        assert _steps([r for r in records if r["pib_raw"].startswith(name)]) == open_ppd
    assert _steps([r for r in records if r["pib_raw"].startswith("ЧЕТВЕРТИЙ")]) == [
        ("РРез", date(2026, 10, 7), None),
    ]


def test_read_all_period_records_aggregates_every_family(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {
        "ВП": (VP_HEADER, [[1, "Р", "матрос", "ПЕРШИЙ Перший Перший", "від", date(2026, 7, 1), None, date(2026, 7, 5), "Перша частина щорічної відпустки"]]),
        "ВД": (VD_HEADER, [[1, "Р", "матрос", "ДРУГИЙ Другий Другий", "вдр", date(2026, 7, 1), None, date(2026, 7, 5), "10"]]),
        "ШП": (SHP_HEADER, [[1, "Р", "матрос", "ТРЕТІЙ Третій Третій", "лік", date(2026, 7, 1), date(2026, 7, 5)]]),
        "ВЛК": (VLK_HEADER,
                [[1, "Р", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "ВЛК", date(2026, 7, 1), date(2026, 7, 5)]]),
        "ППД": (PPD_HEADER,
                [[1, "Р", "матрос", "ДЕСЯТИЙ Десятий Десятий", "ППД", date(2026, 7, 1), None]]),
        "Алф2": (ALF2_HEADER, [[1, "ШОСТИЙ Шостий Шостий", "звільнений", None, None, "Виключений із списків 01.07.2026"]]),
        "СЗЧ": (SZCH_HEADER, [[1, "Р", "матрос", "СЬОМИЙ Сьомий Сьомий", "СЗЧ", date(2026, 7, 1)]]),
        "Втрати": (VTRATY_HEADER, [[1, "Р", "матрос", "ВОСЬМИЙ Восьмий Восьмий", "загибель", date(2026, 7, 1)]]),
        "Прибулі у ВД": (RREZ_HEADER, [
            [1, "idx1", "Р", "Посада1", "матрос", "ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий",
             date(2026, 7, 1), date(2026, 7, 5), date(2026, 7, 10)],
        ]),
    })
    wb = ar.load_addon_workbook(path)

    records = ar.read_all_period_records(wb, "book61")

    assert {r["source_file"] for r in records} == {"book61"}

    pib_raws = {r["pib_raw"] for r in records}
    assert pib_raws == {
        "ПЕРШИЙ Перший Перший", "ДРУГИЙ Другий Другий", "ТРЕТІЙ Третій Третій", "ЧЕТВЕРТИЙ Четвертий Четвертий",
        "ДЕСЯТИЙ Десятий Десятий", "ШОСТИЙ Шостий Шостий", "СЬОМИЙ Сьомий Сьомий", "ВОСЬМИЙ Восьмий Восьмий",
        "ОДИНАДЦЯТИЙ Одинадцятий Одинадцятий",
    }


def test_szch_continues_after_a_move_to_another_units_assignment(tmp_path):
    """Рішення користувача (реальний випадок: "у вдр з 29.07.2026",
    примітка "Повернувся у в/ч ..." іншої частини): для нашої частини СЗЧ
    триває, поки немає Розп чи виключення зі списків; загибель і звичайна
    дата повернення - і далі кінець (за день до дати)."""
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (SZCH_HEADER, [
            [1, "П1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 7, 23), "у вдр з 29.07.2026"],
            [2, "П1", "матрос", "ДРУГИЙ Другий Другий", "СЗЧ", date(2026, 7, 23), "у загибель 01.08.2026"],
            [3, "П1", "матрос", "ТРЕТІЙ Третій Третій", "СЗЧ", date(2026, 7, 23), date(2026, 7, 29)],
        ]),
        "Алф2": (ALF2_HEADER, []),
    })

    records = ar.read_szch_to_rozp_periods(ar.load_addon_workbook(path))

    assert [(r["pib_raw"], r["status"], r["end_date"]) for r in records] == [
        ("ПЕРШИЙ Перший Перший", "СЗЧ", None),
        ("ДРУГИЙ Другий Другий", "СЗЧ", date(2026, 7, 31)),
        ("ТРЕТІЙ Третій Третій", "СЗЧ", date(2026, 7, 28)),
    ]


def test_szch_return_to_the_own_unit_is_followed_by_returned_status(monkeypatch, tmp_path):
    """Рішення користувача: "Примітка" / "Куди повернувся з СЗЧ"
    про повернення у ВЛАСНУ в/ч - з дати повернення "Повер" (СЗЧ - до дня
    перед нею); дата - з "Дата повернення з СЗЧ" (і "у вдр з <дата>") чи з
    самої примітки. Повернення в іншу в/ч - СЗЧ триває; код власної в/ч без
    слова повернення (реальний випадок: "перевели з ... відразу втік") - не
    повернення."""
    monkeypatch.setattr(ar, "OWN_MILITARY_UNIT_CODES", frozenset({"0000"}))
    header = SZCH_HEADER + ["Куди повернувся з СЗЧ", "Примітка"]
    event = date(2026, 7, 23)
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (header, [
            [1, "П1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", event, date(2026, 8, 10), None, "Повернувся у в/ч А0000"],
            [2, "П1", "матрос", "ДРУГИЙ Другий Другий", "СЗЧ", event, "у вдр з 29.07.2026", "Прибув в в/ч A0000", None],
            [3, "П1", "матрос", "ТРЕТІЙ Третій Третій", "СЗЧ", event, None, None, "Повернувся в А0000 15.08.2026"],
            [4, "П1", "матрос", "ЧЕТВЕРТИЙ Четвертий Четвертий", "СЗЧ", event, "у вдр з 29.07.2026", None,
             "Повернувся у в/ч А0001"],
            [5, "П1", "матрос", "П'ЯТИЙ П'ятий П'ятий", "СЗЧ", event, None, None,
             "Новенького перевели з роти в/ч А0000, відразу втік"],
        ]),
        "Алф2": (ALF2_HEADER, []),
    })

    records = ar.read_szch_to_rozp_periods(ar.load_addon_workbook(path))

    assert [(r["pib_raw"].split()[0], r["status"], r["start_date"], r["end_date"]) for r in records] == [
        ("ПЕРШИЙ", "СЗЧ", event, date(2026, 8, 9)), ("ПЕРШИЙ", "Повер", date(2026, 8, 10), None),
        ("ДРУГИЙ", "СЗЧ", event, date(2026, 7, 28)), ("ДРУГИЙ", "Повер", date(2026, 7, 29), None),
        ("ТРЕТІЙ", "СЗЧ", event, date(2026, 8, 14)), ("ТРЕТІЙ", "Повер", date(2026, 8, 15), None),
        ("ЧЕТВЕРТИЙ", "СЗЧ", event, None),
        ("П'ЯТИЙ", "СЗЧ", event, None),
    ]


def test_vp_without_a_destination_is_flagged(tmp_path):
    """Рішення користувача: "Куди вибув" порожнє чи "?" - ознака
    destination_missing (звірка з addons покаже такі дні жовтим); аркуш без
    цієї колонки - не перевіряється (жодної ознаки)."""
    header = VP_HEADER + ["Куди вибув"]
    row = lambda n, name, destination: [n, "П1", "матрос", name, "ВП", date(2026, 10, 1), date(2026, 10, 10), None,
                                        "Щорічна основна відпустка", destination]
    path = write_book61(tmp_path / "book61.xlsm", {"ВП": (header, [
        row(1, "ПЕРШИЙ Перший Перший", "м. Місто1"),
        row(2, "ДРУГИЙ Другий Другий", "?"),
        row(3, "ТРЕТІЙ Третій Третій", None),
        row(4, "ЧЕТВЕРТИЙ Четвертий Четвертий", " ?? "),
    ])})
    records = ar.read_vp_family_periods(ar.load_addon_workbook(path))
    assert [(r["pib_raw"].split()[0], r["destination_missing"]) for r in records] == [
        ("ПЕРШИЙ", False), ("ДРУГИЙ", True), ("ТРЕТІЙ", True), ("ЧЕТВЕРТИЙ", True),
    ]

    old = write_book61(tmp_path / "old.xlsm", {"ВП": (VP_HEADER, [row(1, "ПЕРШИЙ Перший Перший", None)[:-1]])})
    assert "destination_missing" not in ar.read_vp_family_periods(ar.load_addon_workbook(old))[0]


# --- рядки, які читачі пропускають ----------------------------------------------


def test_impossible_embedded_dates_are_ignored():
    assert ar._end_date_from_cell("до 31.02.2026") is None
    assert ar._order_date_from_text("Наказ від 31.02.2026") is None


def test_alf2_readers_skip_rows_without_a_pib(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Алф2": (ALF2_HEADER, [
        [1, None, "самостріл", None, None, "Виключений із списків 10.07.2026"],
    ])})
    wb = ar.load_addon_workbook(path)

    assert ar.read_alf2_order_periods(wb) == []
    assert ar.read_unmapped_alf2_people(wb) == []


def test_read_szch_to_rozp_periods_skips_rows_without_a_pib_or_an_event_date(tmp_path):
    """Розп з Алф2 - лише для тих, хто є на аркуші СЗЧ."""
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (SZCH_HEADER, [
            [1, "Підрозділ1", "матрос", None, "СЗЧ", date(2026, 7, 1)],
            [2, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", None],
        ]),
        "Алф2": (ALF2_HEADER, [
            [1, None, "СЗЧ", None, None, "Виключений із списків 10.07.2026"],
            [2, "ДРУГИЙ Другий Другий", "СЗЧ", None, None, "Виключений із списків 10.07.2026"],
        ]),
    })

    assert ar.read_szch_to_rozp_periods(ar.load_addon_workbook(path)) == []


def test_read_vtraty_periods_skips_rows_without_a_pib_a_known_status_or_a_date(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {"Втрати": (VTRATY_HEADER, [
        [1, "Підрозділ1", "матрос", None, "загибель", date(2026, 7, 1)],
        [2, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "поранення", date(2026, 7, 1)],
        [3, "Підрозділ1", "матрос", "ДРУГИЙ Другий Другий", "загибель", None],
    ])})

    assert ar.read_vtraty_periods(ar.load_addon_workbook(path)) == []


def test_read_rrez_tvo_periods_treats_a_placeholder_enrollment_date_as_missing(tmp_path):
    """01.01.1990 у "Дата зарахування в списки" - заглушка book61 ("ще не
    встановлено"), як і її відсутність: РРез аж до дня перед прийомом посади.
    Рядок без ПІБ пропускається."""
    path = write_book61(tmp_path / "book61.xlsm", {"Прибулі у ВД": (RREZ_HEADER, [
        [1, "idx1", "Підрозділ1", "Посада1", "матрос", None, date(2026, 9, 22), None, date(2026, 10, 7)],
        [2, "idx2", "Підрозділ1", "Посада1", "матрос", "ПЕРШИЙ Перший Перший",
         date(2026, 9, 22), date(1990, 1, 1), date(2026, 10, 7)],
    ])})

    assert _steps(ar.read_rrez_tvo_periods(ar.load_addon_workbook(path))) == [
        ("РРез", date(2026, 9, 22), date(2026, 10, 6)),
        (ar.ANY_STATUS, date(2026, 10, 7), date(2026, 10, 7)),
    ]


def test_rozp_needs_an_order_date_on_alf2(tmp_path):
    path = write_book61(tmp_path / "book61.xlsm", {
        "СЗЧ": (SZCH_HEADER, [[1, "Підрозділ1", "матрос", "ПЕРШИЙ Перший Перший", "СЗЧ", date(2026, 7, 1)]]),
        "Алф2": (ALF2_HEADER, [[1, "ПЕРШИЙ Перший Перший", "СЗЧ", None, None, "наказ ще не підписано"]]),
    })

    records = ar.read_szch_to_rozp_periods(ar.load_addon_workbook(path))

    assert [r["status"] for r in records] == ["СЗЧ"]
