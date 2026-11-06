from colorama import Fore

from utils.logging_utils import print_green, print_red, print_yellow


def test_each_printer_colors_its_text(capsys):
    print_green("готово")
    print_red("помилка")
    print_yellow("попередження")

    lines = capsys.readouterr().out.splitlines()
    assert [line.startswith(color) for line, color in zip(lines, (Fore.GREEN, Fore.RED, Fore.YELLOW))] == [True] * 3
    assert ["готово" in lines[0], "помилка" in lines[1], "попередження" in lines[2]] == [True] * 3
