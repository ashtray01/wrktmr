import os
import sys
import time
import datetime
import openpyxl
import signal
from colorama import init, Fore, Style

init(autoreset=True)

# Папка для логов на рабочем столе пользователя
def get_log_dir():
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    today = datetime.date.today()
    return os.path.join(desktop_dir, "logs", str(today.year), f"{today.month:02d}", f"{today.day:02d}")

LOG_DIR = get_log_dir()
os.makedirs(LOG_DIR, exist_ok=True)

tasks = []
start_time = None
last_task_time = 0  # Для отслеживания времени последней задачи

# Обработчик Ctrl+C
def signal_handler(sig, frame):
    print(Fore.YELLOW + "\nОбнаружено Ctrl+C! Сохраняем данные...")
    save_backup()
    save_excel()
    print(Fore.MAGENTA + "Данные сохранены. Программа завершена.")
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def load_backup():
    """Загружаем задачи из txt, если файл существует"""
    today = datetime.date.today().strftime("%Y-%m-%d")
    log_file = os.path.join(LOG_DIR, f"{today}.txt")
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    task, link, time_str, hours_hundredths = line.strip().split(" | ")
                    # Извлекаем минуты из time_str
                    if time_str == "<1 минуты":
                        minutes = 0.5  # Примерное значение для <1 минуты
                    else:
                        mins, secs = map(int, time_str.split(":"))
                        minutes = mins + secs / 60
                    tasks.append({
                        "task": task,
                        "link": link,
                        "time_str": time_str,
                        "minutes": round(minutes, 2),
                        "hours_hundredths": float(hours_hundredths)
                    })
                except Exception as e:
                    print(Fore.RED + f"Ошибка при чтении строки из backup: {e}")
        if tasks:
            print(Fore.GREEN + f"Загружено {len(tasks)} задач из backup: {log_file}")

def clear_console():
    os.system("cls" if os.name == "nt" else "clear")

def save_backup():
    """Сохраняем в txt после каждого ввода"""
    today = datetime.date.today().strftime("%Y-%m-%d")
    log_file = os.path.join(LOG_DIR, f"{today}.txt")
    with open(log_file, "w", encoding="utf-8") as f:
        for t in tasks:
            f.write(f"{t['task']} | {t['link']} | {t['time_str']} | {t['hours_hundredths']}\n")

def save_excel():
    """Сохраняем Excel-отчет на рабочий стол"""
    today = datetime.date.today().strftime("%Y-%m-%d")
    file_name = os.path.join(LOG_DIR, f"{today}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tasks"

    ws.append(["Задача", "Ссылка", "Время (мин:сек)", "Время (часы в сотых)"])

    total_minutes = 0.0
    total_hours_hundredths = 0.0

    for t in tasks:
        ws.append([t['task'], t['link'], t['time_str'], t['hours_hundredths']])
        total_minutes += float(t['minutes'])
        total_hours_hundredths += float(t['hours_hundredths'])

    ws.append([])
    ws.append(["ИТОГО", "", f"{round(total_minutes, 2)} мин", f"{round(total_hours_hundredths, 2)} ч"])

    max_retries = 5
    retry_delay = 2  # Секунды между попытками
    for attempt in range(max_retries):
        try:
            wb.save(file_name)
            print(Fore.GREEN + f"Excel-отчёт сохранён: {file_name}")
            break
        except PermissionError:
            print(Fore.YELLOW + f"Не удалось сохранить Excel (попытка {attempt + 1}/{max_retries}): файл занят. Ожидаем {retry_delay} сек...")
            time.sleep(retry_delay)
        except Exception as e:
            print(Fore.RED + f"Ошибка при сохранении Excel: {e}")
            break
    else:
        print(Fore.RED + "Не удалось сохранить Excel после всех попыток. Данные сохранены в txt.")

def display_tasks():
    clear_console()
    print(Fore.CYAN + "=== Журнал задач ===")
    for idx, t in enumerate(tasks, start=1):
        print(
            Fore.YELLOW + f"{idx}. {t['task']} [{t['link']}]" +
            Fore.WHITE + f" — {t['time_str']} ({t['hours_hundredths']} ч)"
        )

# Загружаем задачи из backup при старте
load_backup()
display_tasks()

print(Fore.GREEN + "Вводите задачи (для завершённых задач начните с '!'). Команды: :home :h — завершить день, :s :save — сохранить в Excel, :q :quit :exit — выйти.")

while True:
    try:
        task_input = input(Fore.CYAN + "\nВведите задачу: ").strip()
        if task_input.lower() in [":home", ":h", ":q", ":quit", ":exit"]:
            save_excel()
            print(Fore.MAGENTA + "День завершён.")
            break
        elif task_input.lower() in [":s", ":save"]:
            save_excel()
            continue
        elif not task_input:
            print(Fore.RED + "Пожалуйста, введите задачу или команду.")
            continue

        is_completed = task_input.startswith("!")
        task = task_input[1:] if is_completed else task_input

        link = input(Fore.CYAN + "Введите ссылку (если есть): ").strip()

        if is_completed:
            # Для уже выполненных задач запрашиваем время в минутах
            while True:
                try:
                    minutes = float(input(Fore.CYAN + "Введите время, потраченное на задачу (в минутах): ").strip())
                    if minutes < 0:
                        print(Fore.RED + "Время не может быть отрицательным!")
                        continue
                    break
                except ValueError:
                    print(Fore.RED + "Пожалуйста, введите число!")
            
            # Проверяем, прошло ли 3 секунды с последней задачи
            current_time = time.time()
            if current_time - last_task_time < 3:
                time.sleep(3 - (current_time - last_task_time))
            
            input(Fore.CYAN + "Нажмите Enter для подтверждения...")
            last_task_time = time.time()

            elapsed = minutes * 60  # Переводим минуты в секунды
            hours_hundredths = round(minutes / 60, 2)  # Часы как десятичная дробь

            if minutes < 1:
                time_str = "<1 минуты"
            else:
                mins = int(minutes)
                secs = int((minutes % 1) * 60)
                time_str = f"{mins}:{secs:02d}"
        else:
            # Обычный режим с таймером
            start_time = time.time()

            # Проверяем, прошло ли 3 секунды с последней задачи
            current_time = time.time()
            if current_time - last_task_time < 3:
                time.sleep(3 - (current_time - last_task_time))
            
            input(Fore.CYAN + "Нажмите Enter, когда задача завершена...")
            last_task_time = time.time()

            end_time = time.time()
            elapsed = end_time - start_time
            minutes = elapsed / 60
            hours_hundredths = round(elapsed / 3600, 2)  # Часы как десятичная дробь

            if minutes < 1:
                time_str = "<1 минуты"
            else:
                mins = int(minutes)
                secs = int(elapsed % 60)
                time_str = f"{mins}:{secs:02d}"

        tasks.append({
            "task": task,
            "link": link,
            "time_str": time_str,
            "minutes": round(minutes, 2),
            "hours_hundredths": hours_hundredths
        })

        save_backup()
        display_tasks()

    except KeyboardInterrupt:
        signal_handler(None, None)