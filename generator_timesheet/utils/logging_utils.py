from colorama import Fore, Style


# Реальний випадок: символ "⚠" у тексті тут раніше валив ВЕСЬ скрипт
# UnicodeEncodeError (термінал користувача - не UTF-8 кодова сторінка,
# напр. типова cp1251 для україномовної Windows, НЕ вміє саме цей символ,
# хоча кирилицю - вміє) - БЕЗ except навколо, помилка йшла аж до самого
# верху. За прямою вказівкою - жодних символів поза ASCII/кирилицею в
# text (сам червоний/зелений колір - вже достатній візуальний акцент).
def print_green(text):
    print(Fore.GREEN + text + Style.RESET_ALL)


def print_red(text):
    print(Fore.RED + text + Style.RESET_ALL)
