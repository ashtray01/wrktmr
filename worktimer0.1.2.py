import os
import time
import datetime
import sys
from openpyxl import Workbook
from colorama import init, Fore, Style

init()

# Путь к директории скрипта
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

today = datetime.date.today().strftime("%Y-%m-%d")
log_txt_path = os.path.join(SCRIPT_DIR, f"{today}.txt")
log_xlsx_path = os.path.join(SCRIPT_DIR, f"{today}.xlsx")

session_log = []
start_time = None
log_start_time = None
last_action_time = None
total_seconds = 0
current_task = None
task_durations = {}
task_start_times = {}

def format_timer(seconds):
    return str(datetime.timedelta(seconds=int(seconds)))

def format_totals(seconds):
    minutes = seconds / 60
    hours = seconds / 3600
    return f"{int(minutes)} мин / {hours:.2f} ч"

def recover_from_log():
    global session_log, last_action_time, total_seconds, current_task, log_start_time, task_durations, task_start_times
    if os.path.exists(log_txt_path):
        with open(log_txt_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            for i, line in enumerate(lines):
                try:
                    timestamp = line[line.find("[")+1:line.find("]")]
                    task = line[line.find("]")+2:line.rfind("|")].strip() if "|" in line else line[line.find("]")+2:].strip()
                    seconds = int(line[line.rfind("|")+1:line.rfind("сек")].strip()) if "сек" in line else 0
                    if i == 0 and log_start_time is None:
                        log_start_time = datetime.datetime.strptime(f"{today} {timestamp}", "%Y-%m-%d %H:%M:%S").timestamp()
                    if task != "Начало дня":
                        session_log.append((timestamp, task, seconds, timestamp, None))
                        total_seconds += seconds
                        current_task = task
                        task_durations[task] = task_durations.get(task, 0) + seconds
                        task_start_times[task] = timestamp
                except:
                    continue
        last_action_time = time.time()
        print(f"{Fore.YELLOW}Восстановлено из лога: {len(session_log)} записей{Style.RESET_ALL}")
    else:
        log_start_time = time.time()

def log_action(task):
    global last_action_time, total_seconds, current_task, task_durations, task_start_times

    now = time.time()
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")

    if current_task and last_action_time:
        duration = now - last_action_time
        total_seconds += duration
        if session_log and session_log[-1][1] == current_task:
            session_log[-1] = (session_log[-1][0], current_task, int(duration), session_log[-1][3], timestamp)
        else:
            session_log.append((timestamp, current_task, int(duration), task_start_times.get(current_task, timestamp), timestamp))
        task_durations[current_task] = task_durations.get(current_task, 0) + int(duration)
        with open(log_txt_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {current_task} | {int(duration)} сек\n")

    current_task = task
    last_action_time = now
    task_start_times[task] = timestamp
    if task == "Начало дня" or not session_log:
        with open(log_txt_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] Начало дня\n")

def get_unique_xlsx_path():
    base_path = os.path.join(SCRIPT_DIR, f"{today}")
    version = 1
    while os.path.exists(f"{base_path}_ver{version}.xlsx"):
        version += 1
    return f"{base_path}_ver{version}.xlsx" if version > 1 else f"{base_path}.xlsx"

def export_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    ws.append([ "Начало задачи", "Конец задачи", "Задача", "Секунды", "Минуты", "Часы"])

    for timestamp, task, seconds, start_time, end_time in session_log:
        minutes = seconds / 60
        hours = seconds / 3600
        ws.append([start_time, end_time, task, seconds, round(minutes), round(hours, 2)  or ""])

    final_path = get_unique_xlsx_path()
    try:
        wb.save(final_path)
        print(f"{Fore.GREEN}Таблица сохранена: {Fore.MAGENTA}{final_path}")
    except PermissionError:
        print(f"{Fore.RED}Ошибка: Не удалось сохранить {final_path}. Возможно, файл открыт в другой программе. Закройте файл и попробуйте снова.{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Ошибка при сохранении таблицы: {e}{Style.RESET_ALL}")

def display_status():
    now = time.time()
    current_time = datetime.datetime.now().strftime("%H:%M")
    overall_time = format_timer(now - log_start_time) if log_start_time else "00:00:00"
    total_summary = format_totals(total_seconds)
    current_task_duration = task_durations.get(current_task, 0)
    if current_task and last_action_time:
        current_task_duration += now - last_action_time
    task_time = format_totals(current_task_duration) if current_task else "0 мин / 0.00 ч"
    sys.stdout.write(f"\r{Fore.WHITE}Сейчас: {Fore.GREEN}{current_task or 'Нет задачи'}{Style.RESET_ALL}"
                     f"{Fore.CYAN} | {current_time} | {Fore.YELLOW}Всего: {total_summary} | Задача: {task_time}{Style.RESET_ALL}")
    sys.stdout.flush()

def update_display():
    start_time = time.time()
    while current_task and time.time() - start_time < 5:  # Update for 5 seconds or until new input
        display_status()
        time.sleep(1)

def reset_program():
    global session_log, start_time, last_action_time, total_seconds, current_task, log_txt_path, log_xlsx_path, log_start_time, task_durations, task_start_times
    session_log = []
    start_time = time.time()
    last_action_time = time.time()
    total_seconds = 0
    current_task = None
    task_durations = {}
    task_start_times = {}
    log_start_time = time.time()
    today_new = datetime.date.today().strftime("%Y-%m-%d")
    log_txt_path = os.path.join(SCRIPT_DIR, f"{today_new}.txt")
    log_xlsx_path = os.path.join(SCRIPT_DIR, f"{today_new}.xlsx")
    with open(log_txt_path, "w", encoding="utf-8") as f:
        f.write(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Начало дня\n")
    print(f"{Fore.YELLOW}Сброшено! Новый лог: {log_txt_path}{Style.RESET_ALL}")

def main():
    global start_time, last_action_time, log_start_time, today, log_txt_path, log_xlsx_path

    start_time = time.time()
    recover_from_log()
    last_action_time = time.time()

    print(f"\n{Fore.YELLOW}Лог {today}"
          f"\n{Fore.GREEN}Вводи текущую задачу."
          f"\n{Fore.RED}:q, :exit, :end, :выход{Fore.WHITE} выход с сохранением таблицы"
          f"\n{Fore.RED}:s, :save, :сохр{Fore.WHITE} для сохранения таблицы без выхода"
          f"\n{Fore.RED}:r, :reset; :сброс{Fore.WHITE} для сброса таймеров и лога"
          f"\n{Fore.RED}CTRL-C{Fore.WHITE} для справки и обновления таймера текущей задачи")

    while True:
        try:
            current_day = datetime.date.today().strftime("%Y-%m-%d")
            if current_day != today:
                today = current_day
                reset_program()
                continue

            display_status()
            sys.stdout.write("> ")
            sys.stdout.flush()
            task = input().strip()
            if task == "":
                display_status()
                continue
            if task.lower() in [":exit",":q",":end",":выход"]:
                log_action("Конец дня")
                export_to_excel()
                break
            if task.lower() in [":r",":reset",":сброс"]:
                reset_program()
                continue
            if task.lower() in [":сохр",":s",":save"]:
                log_action(current_task or "Сохранение")
                export_to_excel()
                continue
            log_action(task)
            # Update display for a short period to show real-time progress
            update_display()
        except KeyboardInterrupt:
            now = time.time()
            overall_time = format_timer(now - log_start_time) if log_start_time else "00:00:00"
            print(f"\n{Fore.MAGENTA}Время с начала лога: {overall_time}{Style.RESET_ALL}")
            print(f"{Fore.RED}:q, :exit, :end, :выход{Fore.WHITE} выход с сохранением таблицы"
            f"\n{Fore.RED}:s, :save, :сохр{Fore.WHITE} для сохранения таблицы без выхода"
            f"\n{Fore.RED}:r, :reset; :сброс{Fore.WHITE} для сброса таймеров и лога"
            f"\n{Fore.RED}CTRL-C{Fore.WHITE}  для справки и обновления таймера текущей задачи")
        except Exception as e:
            print(f"⚠ Ошибка: {e}")

if __name__ == "__main__":
    main()



