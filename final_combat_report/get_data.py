import subprocess, os, json, sys
from datetime import datetime
from InquirerPy.prompts.input import InputPrompt
from InquirerPy.prompts.list import ListPrompt

# python -m sigexport --chats="Кречет 2.0","Діяльність","Basketball","Пейнтбол","КОМІСІЇ" --overwrite --paginate=0 --source "%APPDATA%\Signal" signal_json

def run_signal_export(*groups):
    appdata = os.getenv('APPDATA')
    if not appdata:
        raise EnvironmentError("APPDATA not set")
    source_path = os.path.join(appdata, 'Signal')
    python_executable = sys.executable
    chats_param = ",".join(groups)
    
    cmd = [python_executable, "-m", "sigexport", f"--chats={chats_param}", "--overwrite", "--paginate=0", "--source", source_path, "signal_json"]

    print(f"Запуск експорту для: {chats_param}")
    
    # Видаляємо shell=True, щоб уникнути подвійного екранування лапок Windows
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.stdout: print("STDOUT:", result.stdout)
    if result.stderr: print("STDERR:", result.stderr)

def get_static_data_json(file_path):
    if not os.path.exists(file_path):
        print(f"Файл {file_path} не знайдено.")
        return []
    try:
        # Open and read the JSON file
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

        return data
    except json.JSONDecodeError as e:
        print(f"Помилка JSON у файлі {file_path}: {e}")
        return {}
    
def get_date(hour_of_report):
    today_str = datetime.now().strftime("%d.%m.%Y")
    
    # Створюємо екземпляр класу напряму
    prompt = InputPrompt(
        message=f"Введіть дату (ДД.ММ.РРРР) для підсумкового починаючи з {hour_of_report}.00 введеної дати по {hour_of_report}.00 від вчора:",
        default=today_str,
    )
    return prompt.execute()

def get_hour_of_report():
    hour = InputPrompt(
        message="Введіть годину початку пбд (від 1 до 24): по дефолту це 18, тобто звітна доба має бути з 18 вчорашнього дня по 18 годину сьогоднішнього дня",
        validate=lambda x: x.isdigit() and 1 <= int(x) <= 24,
        invalid_message="❌ Має бути число від 1 до 24!",
        default="18"
    ).execute()

    return int(hour)

def get_actual_data_json():
    boolean_data = ListPrompt(
        message="Ви хочете використовувати актуальні дані з Signal? (Так(True) / Ні(False))",
        choices=[False, True],
        default=False
    )

    return boolean_data.execute()
