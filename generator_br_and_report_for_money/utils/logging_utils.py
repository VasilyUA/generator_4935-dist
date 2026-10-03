from colorama import Fore, Style


def print_red(s):
    print(Fore.RED + s + Style.RESET_ALL)


def print_green(s):
    print(Fore.GREEN + s + Style.RESET_ALL)


def print_purple(s):
    print(Fore.MAGENTA + s + Style.RESET_ALL)


def print_progress(current, total, prefix=""):
    """Оновлює ОДИН рядок консолі (перезаписує його, а не додає новий) -
    "лоудинг" з відсотками для довгих операцій (напр.
    checker_accounting.report_log_war_checker: зчитування десятків .docx ЖБД +
    сотень записів рапорту) - підтверджено користувачем: без цього незрозуміло,
    чи скрипт ще працює, чи "завис". Викликати з current == total НАПРИКІНЦІ
    циклу - друкує "100%" і завершує рядок переносом ("\\n"), щоб наступні
    print_green/print_red починались із чистого рядка, а не дописувались у
    той самий."""
    percent = 100 if total <= 0 else int(current * 100 / total)
    bar_width = 24
    filled = bar_width if total <= 0 else int(bar_width * current / total)
    bar = "#" * filled + "-" * (bar_width - filled)
    end = "\n" if current >= total else ""
    print(f"\r{prefix}[{bar}] {percent}%", end=end, flush=True)
