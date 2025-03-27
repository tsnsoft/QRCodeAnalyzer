import csv
import logging
import subprocess
import os
from datetime import datetime, timedelta
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QMessageBox
import sys

filename_log="duplicates_log.txt"

# Настройка логирования
logging.basicConfig(
    filename=filename_log,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)

# Очистка лог-файла перед началом работы
with open(filename_log, 'w', encoding='utf-8') as log_file:
    pass  # Открытие файла в режиме 'w' очищает его

# Функция для отображения всплывающего окна
def show_popup(message, title="Предупреждение", buttons=QMessageBox.Ok):
    app = QApplication([])  # Создание приложения
    msg = QMessageBox()  # Создание окна сообщения
    msg.setIcon(QMessageBox.Warning)
    msg.setWindowTitle(title)
    msg.setText(message)
    msg.setStandardButtons(buttons)
    return msg.exec_()

# Функция для проверки повторных отметок
def check_duplicates(csv_file, duplicates_found):
    logging.info(f"Проверка файла: {csv_file}")
    attendance_records = {}

    try:
        with open(csv_file, mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)

            # Проверка наличия нужных столбцов
            expected_columns = {"text", "date_utc", "time_utc"}
            if not expected_columns.issubset(reader.fieldnames):
                logging.error(f"Файл {csv_file} не содержит все необходимые столбцы ({expected_columns}).")
                return duplicates_found

            for row in reader:
                name = row.get("text", "").strip()
                date_utc = row.get("date_utc", "").strip()
                time_utc = row.get("time_utc", "").strip()

                # Проверка на пустые строки
                if not name or not date_utc or not time_utc:
                    logging.warning(f"Пропущенные данные в файле {csv_file}: {row}")
                    continue

                # Объединяем дату и время в один объект datetime
                try:
                    timestamp = datetime.strptime(f"{date_utc} {time_utc}", "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    logging.error(f"Ошибка в формате даты/времени в файле {csv_file}: {row}")
                    continue

                if name in attendance_records:
                    last_timestamp, last_row = attendance_records[name]
                    if timestamp - last_timestamp < timedelta(minutes=5):
                        logging.warning(
                            f"Дубликат в файле {csv_file}: {name} ({date_utc} {time_utc}). "
                            f"Предыдущая отметка: {last_row['date_utc']} {last_row['time_utc']}"
                        )
                        duplicates_found = True
                    else:
                        attendance_records[name] = (timestamp, row)
                else:
                    attendance_records[name] = (timestamp, row)

        return duplicates_found

    except Exception as e:
        logging.error(f"Ошибка при обработке файла {csv_file}: {e}")
        return duplicates_found

# Основная функция для обработки всех файлов в папке Data
def process_data_folder():
    data_folder = Path("Data")
    duplicates_found = False  # Флаг для отслеживания наличия дубликатов

    if not data_folder.exists():
        logging.error(f"Папка '{data_folder}' не найдена.")
        return

    csv_files = list(data_folder.glob("*.csv"))

    if not csv_files:
        logging.info("В папке 'Data' нет файлов CSV для обработки.")
        return

    for csv_file in csv_files:
        duplicates_found = check_duplicates(csv_file, duplicates_found)

    if duplicates_found:
        user_response = show_popup(f"Дубликаты были найдены в процессе обработки файлов. Хотите открыть лог-файл?", 
                                   title="Предупреждение", 
                                   buttons=QMessageBox.Yes | QMessageBox.No)

        if user_response == QMessageBox.Yes:
            # Открытие лог-файла для Windows, Linux и macOS
            try:
                if sys.platform == "win32":
                    # Используем os.startfile для Windows
                    os.startfile(filename_log)
                elif sys.platform == "linux" or sys.platform == "darwin":  # Для Linux и macOS
                    subprocess.run(["xdg-open", filename_log], check=True)  # Для Linux
                else:
                    logging.error(f"Неизвестная операционная система: {sys.platform}")
            except Exception as e:
                logging.error(f"Ошибка при открытии лог-файла: {e}")
                show_popup("Не удалось открыть лог-файл.")
    else:
        logging.info("Повторных отметок не обнаружено в обработанных файлах.")

if __name__ == "__main__":
    process_data_folder()
