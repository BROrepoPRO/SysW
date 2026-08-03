#!/usr/bin/env python3
"""
Запуск макроса RunAllTests в work.xlsm через COM-автоматизацию Excel.
Собирает результаты через VBA-функцию GetTestResults().
Возвращает exit code: 0 — все тесты PASS, 1 — есть FAIL.
Сохраняет результаты в logs/test_results.log.
"""
import sys
import time
import os
import win32com.client
from win32com.client import gencache

from config import WORKBOOK_PATH, TEST_LOG_FILE, LOGS_DIR
EXCEL_PATH = str(WORKBOOK_PATH)
LOG_FILE = str(TEST_LOG_FILE)


def ensure_logs_dir():
    """Создаёт директорию logs/ если её нет."""
    os.makedirs(str(LOGS_DIR), exist_ok=True)


def write_log(message: str):
    """Запись сообщения в лог-файл и вывод в консоль."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    # Создаём директорию logs/ при каждом запуске
    ensure_logs_dir()

    print("=" * 60)
    print("ЗАПУСК ТЕСТОВ (TC-01..TC-44)")
    print("=" * 60)
    print()

    excel = None
    workbook = None
    exit_code = 0

    try:
        write_log("[1/5] Создание COM-объекта Excel...")
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        write_log("[2/5] Открытие work.xlsm...")
        workbook = excel.Workbooks.Open(EXCEL_PATH)

        write_log("[3/5] Установка тестового значения B4=3 и запуск макроса RunAllTests...")

        # Устанавливаем тестовое значение 3 в ячейку B4 листа main,
        # чтобы избежать срабатывания Worksheet_Change с некорректными данными
        try:
            ws_main = workbook.Sheets("main")
            ws_main.Range("B4").Value = 3
            write_log("    Установлено B4=3 на листе main")
        except Exception as e:
            write_log(f"    [ПРЕДУПРЕЖДЕНИЕ] Не удалось установить B4: {e}")

        print()
        print("    Ожидание завершения тестов...")
        print()

        excel.Run("RunAllTests")

        # Небольшая пауза для завершения всех операций
        time.sleep(1)

        write_log("[4/5] Чтение результатов из ячейки Z1 листа main...")
        try:
            results = ws_main.Range("Z1").Value
            if results:
                write_log("    Результаты получены из ячейки Z1")
            else:
                results = ""
                write_log("    [ПРЕДУПРЕЖДЕНИЕ] Ячейка Z1 пуста")
        except Exception as e:
            write_log(f"    [ПРЕДУПРЕЖДЕНИЕ] Не удалось прочитать Z1: {e}")
            results = ""

        print()
        print("-" * 60)
        print("РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ")
        print("-" * 60)

        if results:
            print(results)
            write_log(f"Сырые результаты: {results.strip()}")

            # Парсим статистику
            total = 0
            passed = 0
            failed = 0
            skipped = 0

            for line in results.split("\n"):
                line = line.strip()
                if line.startswith("Total="):
                    parts = line.split(";")
                    for part in parts:
                        if "=" in part:
                            key, val = part.split("=")
                            val = val.strip()
                            if key == "Total":
                                total = int(val)
                            elif key == "Passed":
                                passed = int(val)
                            elif key == "Failed":
                                failed = int(val)
                            elif key == "Skipped":
                                skipped = int(val)

            print()
            print("=" * 60)
            print(f"  Всего:      {total}")
            print(f"  Пройдено:   {passed}")
            print(f"  Провалено:  {failed}")
            print(f"  Пропущено:  {skipped}")
            print("=" * 60)
            print()

            if failed > 0:
                print("  ❌ ОБНАРУЖЕНЫ ОШИБКИ! Проверьте логи для деталей.")
                exit_code = 1
            else:
                print("  ✅ Все тесты успешно пройдены!")
                exit_code = 0

            write_log(f"Итог: Total={total}, Passed={passed}, Failed={failed}, Skipped={skipped}")
        else:
            print("    [ПРЕДУПРЕЖДЕНИЕ] Не удалось получить результаты через GetTestResults.")
            print("    Проверьте Immediate Window в редакторе VBA для деталей.")
            write_log("Результаты не получены (GetTestResults недоступна)")
            exit_code = 1

        write_log("[5/5] Завершение работы.")

    except Exception as e:
        error_msg = f"ОШИБКА: {e}"
        print(error_msg, file=sys.stderr)
        write_log(error_msg)
        exit_code = 1

    finally:
        # Гарантированное закрытие Excel
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
                write_log("Книга закрыта.")
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
                write_log("Excel закрыт.")
            except Exception:
                pass

        print()
        print("=" * 60)
        if exit_code == 0:
            print("РЕЗУЛЬТАТ: ВСЕ ТЕСТЫ ПРОЙДЕНЫ")
        else:
            print("РЕЗУЛЬТАТ: ОБНАРУЖЕНЫ ОШИБКИ")
        print("=" * 60)
        print()

    return exit_code


if __name__ == "__main__":
    sys.exit(main())