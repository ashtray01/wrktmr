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

def format_timer(seconds):
    return str(datetime.timedelta(seconds=int(seconds)))

def format_totals(seconds):
    minutes = seconds / 60
    hours = seconds / 3600
    return f"{int(minutes)} мин / {hours:.2f} ч"

def recover_from_log():
    global session_log, last_action_time, total_seconds, current_task, log_start_time
    if os.path.exists(log_txt_path):
        with open(log_txt_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            for i, line in enumerate(lines):
                try:
                    timestamp = line[line.find("[")+1:line.find("]")]
                    task = line[line.find("]")+2:line.rfind("|")].strip() if "|" in line else line[line.find("]")+2:].strip()
                    seconds = int(line[line.rfind("|")+1:line.rfind("сек")].strip()) if "сек" in line else 0
                    if i == 0 and log_start_time is None:  # Первая запись определяет начало лога
                        log_start_time = datetime.datetime.strptime(f"{today} {timestamp}", "%Y-%m-%d %H:%M:%S").timestamp()
                    if task != "Начало дня":
                        session_log.append((timestamp, task, seconds))
                        total_seconds += seconds
                        current_task = task
                except:
                    continue
        last_action_time = time.time()
        print(f"{Fore.YELLOW}Восстановлено из лога: {len(session_log)} записей{Style.RESET_ALL}")
    else:
        # Если лога нет, устанавливаем log_start_time на момент запуска программы
        log_start_time = time.time()

def log_action(task):
    global last_action_time, total_seconds, current_task

    now = time.time()
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")

    if current_task and last_action_time:
        duration = now - last_action_time
        total_seconds += duration
        session_log.append((timestamp, current_task, int(duration)))
        with open(log_txt_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {current_task} | {int(duration)} сек\n")

    current_task = task
    last_action_time = now
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
    ws.append(["Время", "Задача", "Секунды", "Минуты", "Часы"])

    for timestamp, task, seconds in session_log:
        minutes = seconds / 60
        hours = seconds / 3600
        ws.append([timestamp, task, seconds, round(minutes, 2), round(hours, 2)])

    final_path = get_unique_xlsx_path()
    try:
        wb.save(final_path)
        print(f"\n📄 Excel сохранён: {final_path}")
    except PermissionError:
        print(f"{Fore.RED}⚠ Ошибка: Не удалось сохранить {final_path}. Возможно, файл открыт в другой программе. Закройте файл и попробуйте снова.{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}⚠ Ошибка при сохранении Excel: {e}{Style.RESET_ALL}")

def display_status():
    now = time.time()
    overall_time = format_timer(now - log_start_time) if log_start_time else "00:00:00"
    total_summary = format_totals(total_seconds)
    print(f"\r{Fore.WHITE}Сейчас: {Fore.GREEN}{current_task or 'Нет задачи'}{Style.RESET_ALL}"
          f"{Fore.CYAN} | {overall_time} | {Fore.YELLOW}{total_summary}{Style.RESET_ALL}")

def reset_program():
    global session_log, start_time, last_action_time, total_seconds, current_task, log_txt_path, log_xlsx_path, log_start_time
    session_log = []
    start_time = time.time()
    last_action_time = time.time()
    total_seconds = 0
    current_task = None
    log_start_time = time.time()  # Устанавливаем начало лога на момент сброса
    today_new = datetime.date.today().strftime("%Y-%m-%d")
    log_txt_path = os.path.join(SCRIPT_DIR, f"{today_new}.txt")
    log_xlsx_path = os.path.join(SCRIPT_DIR, f"{today_new}.xlsx")
    with open(log_txt_path, "w", encoding="utf-8") as f:
        f.write(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Начало дня\n")
    print(f"{Fore.YELLOW}Сброшено! Новый лог: {log_txt_path}{Style.RESET_ALL}")

def main():
    global start_time, last_action_time, log_start_time, today, log_txt_path, log_xlsx_path

    start_time = time.time()
    recover_from_log()  # Восстанавливаем лог или устанавливаем log_start_time
    last_action_time = time.time()

    print(f"\n📋 Лог {today}{Fore.GREEN}. Вводи текущие задачи. Для выхода ':exit', ':e', ':1', ':q', ':excel', ':end', "
          f"':сброс' для сброса таймеров и лога, ':сохр' для сохранения Excel.\n При выходе эксель автоматически сохраняется рядом со скриптом.")

    while True:
        try:
            # Проверка смены дня
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
                continue
            if task.lower() in [":exit", ":e", ":1", ":q", ":excel", ":end"]:
                log_action("Конец дня")
                export_to_excel()
                break
            if task.lower() == ":сброс":
                reset_program()
                continue
            if task.lower() == ":сохр":
                log_action(current_task or "Сохранение")
                export_to_excel()
                continue
            log_action(task)
        except KeyboardInterrupt:
            now = time.time()
            overall_time = format_timer(now - log_start_time) if log_start_time else "00:00:00"
            print(f"\n{Fore.MAGENTA}Время с начала лога: {overall_time}{Style.RESET_ALL}")
            print("\n🚪 Используй ':exit', ':e', ':1', ':q', ':excel', ':end' для завершения, "
                  "':сброс' для сброса таймеров и лога, ':сохр' для сохранения Excel.")
        except Exception as e:
            print(f"⚠ Ошибка: {e}")

if __name__ == "__main__":
    main()