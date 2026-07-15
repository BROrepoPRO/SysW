#!/usr/bin/env python3
"""
Запуск макроса RunAllTests в work.xlsm через COM-автоматизацию Excel.
Собирает результаты через VBA-функцию GetTestResults().
Возвращает exit code: 0 — все тесты PASS, 1 — есть FAIL.
Сохраняет результаты в test_results.log.
"""
import sys
import time
import os
import win32com.client
from win32com.client import gencache

EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"
LOG_FILE = os.path.join(os.path.dirname(EXCEL_PATH), "test_results.log")


def write_log(message: str):
    """Запись сообщения в лог-файл и вывод в консоль."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    print("=" * 60)
    print("ЗАПУСК ТЕСТОВ (TC-01..TC-30)")
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

        write_log("[3/5] Запуск макроса RunAllTests...")
        print()
        print("    Ожидание завершения тестов...")
        print()

        excel.Run("RunAllTests")

        # Небольшая пауза для завершения всех операций
        time.sleep(1)

        write_log("[4/5] Сбор результатов через GetTestResults()...")
        try:
            results = excel.Run("GetTestResults")
        except Exception:
            # Fallback: если функция недоступна, парсим лог
            write_log("    [ПРЕДУПРЕЖДЕНИЕ] GetTestResults не найдена, используется fallback")
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