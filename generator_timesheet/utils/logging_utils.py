from colorama import Fore, Style


# Лише ASCII й кирилиця в text: символ "⚠" колись валив увесь скрипт з
# UnicodeEncodeError, бо консоль користувача в cp1251, а не UTF-8. Колір і
# так достатньо виділяє повідомлення.
def print_green(text):
    print(Fore.GREEN + text + Style.RESET_ALL)


def print_red(text):
    print(Fore.RED + text + Style.RESET_ALL)


def print_yellow(text):
    """Попередження, а не помилка."""
    print(Fore.YELLOW + text + Style.RESET_ALL)
