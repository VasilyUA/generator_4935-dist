"""Назви колонок ОБЛІК.xlsx, записані в реєстрах constants.py (BASIS_REQUIRED_POINT_COLUMN_NAMES,
POINT_OWN_PIDSTAVY_COLUMN_NAMES, GENERAL_PIDSTAVY_COLUMN_NAMES) - у БУДЬ-ЯКІЙ формі.

Одна назва - рядком, з комою чи без ("ПІДСТАВИ 10К" або ("ПІДСТАВИ 10К",)), кілька - кортежем,
списком чи множиною; регістр і кількість пробілів не важливі. Усі місця, що читають реєстри,
пропускають значення через ці функції: ітерація по голому рядку дала б окремі літери ("П", "І",
...) і колонка мовчки не читалась би. Модуль без залежностей, тож його можна імпортувати звідусіль
(зокрема з constants.py) без циклічних імпортів."""


def canonical_column_name(name):
    """Канонічна форма назви колонки: регістр і кількість пробілів (включно з переносом рядка
    всередині заголовка Excel) не мають значення."""
    return " ".join(str(name).split()).upper()


def column_names_of(value):
    """Кортеж канонічних назв колонки з будь-якої форми запису реєстру. Порожні значення
    відкидаються, дублі прибираються зі збереженням порядку першої появи. Для множини порядок
    сортується (у множині його немає), щоб результат був детермінованим."""
    if value is None:
        return ()
    if isinstance(value, str):
        candidates = (value,)
    elif isinstance(value, (set, frozenset)):
        candidates = tuple(sorted(value, key=str))
    else:
        try:
            candidates = tuple(value)
        except TypeError:
            candidates = (value,)

    names = []
    for candidate in candidates:
        if candidate is None:
            continue
        name = canonical_column_name(candidate)
        if name and name not in names:
            names.append(name)
    return tuple(names)
