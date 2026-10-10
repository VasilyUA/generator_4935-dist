from colorama import Fore, Style


def print_green(text):
    print(Fore.GREEN + text + Style.RESET_ALL)


def print_red(text):
    print(Fore.RED + text + Style.RESET_ALL)
