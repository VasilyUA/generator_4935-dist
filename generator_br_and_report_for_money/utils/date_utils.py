import calendar
from datetime import datetime, timedelta


def to_date(x):
    if isinstance(x, datetime):
        return x.date()
    elif isinstance(x, str):
        # .strip() - constants.py (MONEY_REPORT_GENERAL_REFERENCES_LOG_WAR тощо) редагується
        # вручну, і зайвий пробіл на кінці дати (напр. "08.08.2026 ") - реальна помилка, яку
        # strptime інакше сприймає як "невідомий формат" замість очевидного typo.
        for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(x.strip(), fmt).date()
            except ValueError:
                continue
    raise ValueError(f"Невідомий формат або тип дати: {x!r}")


def date_to_str(date_str, action='-', days=0):
    delta = timedelta(days if action == '+' else -days)
    return (to_date(date_str) + delta).strftime('%d.%m.%Y')


def get_bat_period_and_variant(col_name, number_of_documents_brs, num_variants=3):
    """Ділить місяць на періоди за ФАКТИЧНИМИ датами БР бат (constants.NUMBER_OF_DOCUMENTS_BRS_EVERY_WEEK,
    записи з непустим полем 'бат') - а не за фіксованою 7-денною сіткою від 1-го числа:
    реальні БР видаються нерівномірно (напр. 01, 07, 14, 21, 27 числа - остання пара лише
    6 днів, а не 7), тож розпорядження "ЗАВДАННЯ" має описувати період до ФАКТИЧНОЇ дати
    наступного БР бат, а не до умовної межі тижня.

    col_name - дата розпорядження, що генерується; ОЧІКУЄТЬСЯ, що для неї в
    number_of_documents_brs є запис із непустим 'бат' (викликач сам перевіряє цю умову
    перед викликом - генерація "ЗАВДАННЯ" відбувається лише для таких дат).

    Розпорядження від 1-го числа описує період, що починається того ж дня;
    розпорядження, видане не 1-го числа, описує НАСТУПНИЙ період (від наступного дня) -
    так само, як і раніше (get_weekly_period_and_variant). Період закінчується датою
    НАСТУПНОГО (хронологічно за col_name) непустого 'бат' цього ж місяця - або останнім
    днем місяця, якщо col_name - останній непустий 'бат' місяця.

    Індекс варіанту - порядковий номер col_name серед УСІХ непустих 'бат' цього місяця
    (0-відлік, за колом від 0 до num_variants-1), а НЕ номер тижня - тож ротація завжди
    йде рівно по одному варіанту на кожне фактичне розпорядження, незалежно від того,
    скільки днів між ними насправді."""
    date = to_date(col_name)
    bat_dates = sorted(
        to_date(date_str) for date_str, entry in number_of_documents_brs.items()
        if entry.get('бат') and to_date(date_str).year == date.year and to_date(date_str).month == date.month
    )
    index = bat_dates.index(date)
    variant_index = index % num_variants

    start_date = date if date.day == 1 else date + timedelta(days=1)
    if index + 1 < len(bat_dates):
        end_date = bat_dates[index + 1]
    else:
        last_day_of_month = calendar.monthrange(date.year, date.month)[1]
        end_date = date.replace(day=last_day_of_month)

    period_text = f"з {start_date.strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')}"
    return period_text, variant_index


def get_prev_general(date_str, d):
    """Повертає ключ (дата-рядок) із НАЙПІЗНІШОЮ датою серед d, що СТРОГО РАНІШЕ за
    date_str І має непорожній 'бат' - тобто дійсно РЕАЛЬНЕ попереднє щоденне
    розпорядження, а не сам date_str (навіть якщо він теж є ключем d - як завжди
    буде для реального виклику, generate_extract_log_war_general_br_every_day.py
    передає СЬОГОДНІШНЮ дату, для якої в d вже є власний запис) і не "заглушка"
    сітки дат із порожнім 'бат' (напр. кінець попереднього місяця до першого
    реального щоденного розпорядження - на таку дату посилатись у звіті про
    виконання не можна, там немає жодного номера). None - якщо date_str не
    парситься, чи в d взагалі немає жодного запису, що відповідає обом умовам
    (напр. це найперший день, за який дійсно ведеться облік)."""
    try:
        t = datetime.strptime(date_str.strip(), "%d.%m.%Y")
    except Exception:
        return None

    earlier = [
        (k, datetime.strptime(k, "%d.%m.%Y"))
        for k, v in d.items()
        if v.get('бат') and datetime.strptime(k, "%d.%m.%Y") < t
    ]
    if not earlier:
        return None
    return max(earlier, key=lambda item: item[1])[0]
