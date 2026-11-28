"""Назви колонок ОБЛІК.xlsx, записані в constants.py у полі "column_basis_from_tracking"
окремої категорії MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES - у БУДЬ-ЯКІЙ формі.

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


def collect_column_basis_names(categories):
    """Усі назви колонок "column_basis_from_tracking" усіх пунктів categories
    (словник у форматі constants.MONEY_REPORT_CATEGORIES/COMMANDER_MONEY_REPORT_CATEGORIES) -
    як на рівні самого пункту (categories[точка]["column_basis_from_tracking"] -
    спільна колонка для ВСІХ його категорій, якщо жодна не перевизначає свою
    власну), так і на рівні кожної названої категорії (categories[точка][категорія]
    ["column_basis_from_tracking"], окрім "general"). Поле відсутнє чи None
    (типовий випадок для пункту/категорії без власної колонки) - просто нічого
    не додає. Дублі прибираються, порядок - перша поява."""
    names = []
    for point_categories in categories.values():
        point_level_tracking = point_categories.get("column_basis_from_tracking")
        for name in column_names_of(point_level_tracking):
            if name not in names:
                names.append(name)
        for category_name, config in point_categories.items():
            if category_name == "general" or not isinstance(config, dict):
                continue
            raw_value = config.get("column_basis_from_tracking", point_level_tracking)
            for name in column_names_of(raw_value):
                if name not in names:
                    names.append(name)
    return tuple(names)
