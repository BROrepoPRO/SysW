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
import argparse
import json
import re
import html
from datetime import datetime
from pathlib import Path
import win32com.client
import win32process
from win32com.client import gencache
import gc

from config import WORKBOOK_PATH, TEST_LOG_FILE, LOGS_DIR
EXCEL_PATH = str(WORKBOOK_PATH)
LOG_FILE = str(TEST_LOG_FILE)

# Версия приложения для отчётов (единый источник версии проекта).
APP_VERSION = "1.1.1.1"


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


# Имя этапа для лог-файла PID (используется build_all.py для безопасного taskkill).
STAGE_NAME = "tests"

# Константы ретраев открытия книги.
OPEN_RETRIES = 5   # Максимальное число попыток открытия.
OPEN_PAUSE = 3     # Пауза (сек) между попытками.


def open_workbook_with_retry(excel, path, retries=OPEN_RETRIES, pause_sec=OPEN_PAUSE):
    """Открывает книгу с ретраями при None/COM-ошибке -2147352567.

    Workbooks.Open может вернуть None либо выбросить прерывистую COM-ошибку
    при модальных окнах/плохом состоянии экземпляра. Повторяет открытие до
    retries раз с паузой pause_sec и сборкой мусора между попытками. Если все
    попытки исчерпаны — возвращает None (основной try/except обработает это).
    """
    for attempt in range(1, retries + 1):
        wb = None
        try:
            # Явные параметры: ReadOnly=False (на запись), UpdateLinks=0
            # (не обновлять связи) — защита от зависаний на модальных окнах
            # обновления связей/переключения в режим только для чтения.
            # ВАЖНО: параметр ConfirmConversions НЕ используется — он был удалён
            # из сигнатуры Workbooks.Open в современных версиях Excel (pywin32
            # отклоняет его как неизвестный именованный аргумент).
            wb = excel.Workbooks.Open(
                str(path),
                ReadOnly=False,
                UpdateLinks=0,
            )
        except Exception as exc:
            msg = str(exc)
            write_log(f"    Попытка {attempt}/{retries}: ошибка открытия ({msg}); "
                      f"пауза {pause_sec}с")
            wb = None
            time.sleep(pause_sec)
            gc.collect()
            continue

        if wb is None:
            write_log(f"    Попытка {attempt}/{retries}: Workbooks.Open вернул None; "
                      f"пауза {pause_sec}с")
            time.sleep(pause_sec)
            gc.collect()
            continue

        return wb

    return None


def write_excel_pid(excel):
    """Записывает PID созданного экземпляра Excel в лог-файл для build_all.py.

    build_all.py после этапа читает logs/excel_pid_<stage>.txt и завершает
    ТОЛЬКО этот процесс (по PID), не трогая чужие сессии Excel пользователя.
    При сбое записи PID просто логируем предупреждение — не критично.
    """
    try:
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        pid_file = LOGS_DIR / f"excel_pid_{STAGE_NAME}.txt"
        pid_file.parent.mkdir(parents=True, exist_ok=True)
        pid_file.write_text(str(pid), encoding="utf-8")
        write_log(f"    Записан PID Excel {pid} -> {pid_file}")
    except Exception as exc:
        write_log(f"    [!] Не удалось записать PID Excel: {exc}")


# ============================================================
# R-16: Парсинг лога тестов и генерация отчётов JSON/HTML
# ============================================================

# Регулярное выражение строки лога тестов (формат VBA Mod_Logger.WriteTestLog):
# [yyyy-mm-dd hh:nn:ss] [LEVEL] [Module] message
_LINE_RE = re.compile(
    r"^\[(?P<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] "
    r"\[(?P<level>INFO|WARN|ERROR)\] "
    r"\[(?P<module>[^\]]+)\]\s*(?P<msg>.*)$"
)
# Идентификатор тестового модуля (TC-xx, регистронезависимо).
_TC_ID_RE = re.compile(r"^TC-[A-Za-z0-9]+$", re.IGNORECASE)
# Сводка из строки завершения сессии RunAllTests: END (Total=..;Passed=..;Failed=..;Skipped=..)
_END_SUMMARY_RE = re.compile(r"Total=(\d+);Passed=(\d+);Failed=(\d+);Skipped=(\d+)")
# Модуль-маркер сессии полного прогона тестов.
_SESSION_MODULE = "Mod_FullTestRunner"


def _read_text_robust(path):
    """Читает текстовый файл байтами с построчным автоопределением кодировки.

    Лог тестов дописывается и VBA (ANSI/CP1251 для кириллицы), и Python
    (UTF-8), из-за чего файл может содержать СМЕШАННУЮ кодировку по строкам.
    Каждая физическая строка написана ровно одним писателем, поэтому декодируем
    файл пострoчно: сначала пробуем utf-8, затем cp1251, в крайнем случае
    errors='replace'. Это корректно восстанавливает кириллицу без \ufffd.
    """
    raw = Path(path).read_bytes()
    out = []
    for ln in raw.split(b"\n"):
        ln = ln.rstrip(b"\r")
        decoded = None
        for enc in ("utf-8", "cp1251"):
            try:
                decoded = ln.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            decoded = ln.decode("utf-8", errors="replace")
        out.append(decoded)
    return "\n".join(out)


def _parse_test_message(msg):
    """Разбирает сообщение строки лога на кортеж (имя, статус, детали) или None."""
    if msg.endswith(" PASS"):
        return msg[:-5].strip(), "PASS", ""
    if msg.endswith(" FAIL"):
        return msg[:-5].strip(), "FAIL", ""
    if " SKIP: " in msg:
        name, _, detail = msg.partition(" SKIP: ")
        return name.strip(), "SKIP", detail.strip()
    if " FAIL: " in msg:
        name, _, detail = msg.partition(" FAIL: ")
        return name.strip(), "FAIL", detail.strip()
    return None


def parse_test_log(path) -> list:
    """Структурирует записи тестов из лога test_results.log.

    Возвращает список словарей:
    {"id","name","status","detail","level","timestamp"}.
    Не тестовые строки (Start/End сессии, системные) пропускаются.
    """
    text = _read_text_robust(path)
    tests = []
    for line in text.splitlines():
        m = _LINE_RE.match(line)
        if not m:
            continue
        module = m.group("module").strip()
        if not _TC_ID_RE.match(module):
            continue
        parsed = _parse_test_message(m.group("msg").strip())
        if parsed is None:
            continue
        name, status, detail = parsed
        tests.append({
            "id": module,
            "name": name,
            "status": status,
            "detail": detail,
            "level": m.group("level").strip(),
            "timestamp": m.group("ts").strip(),
        })
    return tests


def parse_end_summary(msg):
    """Извлекает сводку Total/Passed/Failed/Skipped из строки END сессии."""
    m = _END_SUMMARY_RE.search(msg)
    if not m:
        return None
    return {
        "total": int(m.group(1)),
        "passed": int(m.group(2)),
        "failed": int(m.group(3)),
        "skipped": int(m.group(4)),
    }


def parse_session_meta(path) -> dict:
    """Собирает метаданные сессии прогона (start/end и итоговая сводка)."""
    text = _read_text_robust(path)
    meta = {"start": None, "end": None, "end_summary": None}
    for line in text.splitlines():
        m = _LINE_RE.match(line)
        if not m:
            continue
        if m.group("module").strip() != _SESSION_MODULE:
            continue
        msg = m.group("msg").strip()
        if "RunAllTests: START" in msg:
            meta["start"] = m.group("ts").strip()
        elif "RunAllTests: END" in msg:
            meta["end"] = m.group("ts").strip()
            meta["end_summary"] = parse_end_summary(msg)
    return meta


def parse_summary_from_z1(results_text) -> dict:
    """Парсит сводку Total/Passed/Failed/Skipped из ячейки Z1 листа main."""
    summary = {"total": 0, "passed": 0, "failed": 0, "skipped": 0}
    if not results_text:
        return summary
    for line in results_text.split("\n"):
        line = line.strip()
        if not line.startswith("Total="):
            continue
        for part in line.split(";"):
            if "=" not in part:
                continue
            key, val = part.split("=", 1)
            val = val.strip()
            try:
                if key == "Total":
                    summary["total"] = int(val)
                elif key == "Passed":
                    summary["passed"] = int(val)
                elif key == "Failed":
                    summary["failed"] = int(val)
                elif key == "Skipped":
                    summary["skipped"] = int(val)
            except ValueError:
                pass
    return summary


def apply_module_filter(tests, pattern):
    """Фильтрует тесты по подстроке pattern против id (TC-xx) и имени теста.

    Если pattern пустой/None — возвращает список без изменений.
    """
    if not pattern:
        return tests
    low = pattern.lower()
    return [t for t in tests if low in t["id"].lower() or low in t["name"].lower()]


def _count_by_status(tests):
    return {
        "total": len(tests),
        "passed": sum(1 for t in tests if t["status"] == "PASS"),
        "failed": sum(1 for t in tests if t["status"] == "FAIL"),
        "skipped": sum(1 for t in tests if t["status"] == "SKIP"),
    }


def build_report(tests, summary, args, session_meta) -> dict:
    """Собирает JSON-структуру отчёта по спецификации R-16."""
    filtered = apply_module_filter(tests, args.module)
    return {
        "schema_version": "1.0",
        "generated_at": datetime.now().astimezone().isoformat(),
        "app_version": APP_VERSION,
        "workbook": EXCEL_PATH,
        "command_line": list(sys.argv),
        "filter": args.module,
        "session": session_meta,
        "summary": summary,
        "summary_filtered": _count_by_status(filtered),
        "tests": filtered,
    }


def save_json_report(report, out_dir):
    """Сохраняет отчёт в test_report.json (UTF-8, ensure_ascii=False)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "test_report.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


_REPORT_CSS = """
body{font-family:'Segoe UI',Arial,sans-serif;margin:20px;background:#f4f6f8;color:#222;}
h1{color:#263238;}
.cards{display:flex;gap:16px;margin:20px 0;flex-wrap:wrap;}
.card{border-radius:8px;padding:16px 26px;color:#fff;min-width:110px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.2);}
.card.total{background:#546e7a;}
.card.pass{background:#2e9e4f;}
.card.skip{background:#e0a13f;}
.card.fail{background:#d9534f;}
.card-num{font-size:28px;font-weight:bold;}
.card-label{font-size:13px;opacity:.95;}
table{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.15);}
th,td{border:1px solid #e0e0e0;padding:8px 12px;text-align:left;vertical-align:top;}
th{background:#eceff1;}
.st{font-weight:bold;}
tr.pass .st{color:#2e9e4f;}
tr.skip .st{color:#e0a13f;}
tr.fail .st{color:#d9534f;}
.footer{margin-top:20px;color:#777;font-size:12px;}
"""


def render_html(report) -> str:
    """Генерирует автономный HTML-отчёт с инлайн CSS (UTF-8)."""
    e = html.escape
    s = report.get("summary") or {}
    gen = e(str(report.get("generated_at", "")))
    flt = report.get("filter")
    flt_text = flt if flt else "нет (все тесты)"
    app_ver = e(str(report.get("app_version", "")))

    cards = (
        ("Всего", s.get("total", 0), "total"),
        ("Пройдено", s.get("passed", 0), "pass"),
        ("Провалено", s.get("failed", 0), "fail"),
        ("Пропущено", s.get("skipped", 0), "skip"),
    )
    card_html = "".join(
        '<div class="card {cls}"><div class="card-num">{num}</div>'
        '<div class="card-label">{lbl}</div></div>'.format(cls=cls, num=num, lbl=lbl)
        for lbl, num, cls in cards
    )

    rows = []
    for t in report.get("tests", []):
        cls = str(t.get("status", "")).lower()
        rows.append(
            '<tr class="{cls}"><td>{tid}</td><td>{tname}</td>'
            '<td class="st">{status}</td><td>{detail}</td></tr>'.format(
                cls=cls,
                tid=e(str(t.get("id", ""))),
                tname=e(str(t.get("name", ""))),
                status=e(str(t.get("status", ""))),
                detail=e(str(t.get("detail", ""))),
            )
        )
    tbody = "\n".join(rows) if rows else '<tr><td colspan="4">Нет данных</td></tr>'

    return (
        "<!DOCTYPE html>\n<html lang=\"ru\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<title>Отчёт о тестировании SysW</title>\n<style>" + _REPORT_CSS +
        "</style>\n</head>\n<body>\n"
        "<h1>Отчёт о тестировании SysW</h1>\n"
        '<div class="cards">' + card_html + "</div>\n"
        "<p>Сгенерировано: <strong>" + gen + "</strong> · фильтр: <strong>" +
        e(flt_text) + "</strong></p>\n"
        "<table>\n<thead><tr><th>ID</th><th>Тест</th><th>Статус</th>"
        "<th>Детали</th></tr></thead>\n<tbody>\n" + tbody +
        "\n</tbody>\n</table>\n"
        '<div class="footer">SysW v' + app_ver + " · " + gen + "</div>\n"
        "</body>\n</html>\n"
    )


def save_html_report(report, out_dir):
    """Сохраняет HTML-отчёт в test_report.html (UTF-8)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "test_report.html"
    path.write_text(render_html(report), encoding="utf-8")
    return path


def print_verbose(tests):
    """Печатает результат каждого отфильтрованного теста в консоль."""
    for t in tests:
        line = "[{}] {} {}".format(t["status"], t["id"], t["name"])
        if t.get("detail"):
            line += " — {}".format(t["detail"])
        print(line)


def build_parser():
    """Создаёт argparse-парсер аргументов командной строки (описания на русском)."""
    parser = argparse.ArgumentParser(
        prog="run_tests.py",
        description="Запуск макроса RunAllTests в work.xlsm и (опционально) "
                    "генерация отчётов JSON/HTML по логу тестов.",
    )
    parser.add_argument(
        "--module", default=None,
        help="Фильтр тестов по подстроке (id TC-xx или имя теста). "
             "Без фильтра выводятся все тесты.",
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Печатать результат каждого теста в консоль.",
    )
    parser.add_argument(
        "--output", default=None,
        help="Директория для сохранения отчётов test_report.json/.html "
             "(по умолчанию logs/).",
    )
    return parser


def parse_args(argv=None):
    """Разбирает аргументы командной строки."""
    return build_parser().parse_args(argv)


def main(argv=None) -> int:
    # Создаём директорию logs/ при каждом запуске
    ensure_logs_dir()

    args = parse_args(argv)

    print("=" * 60)
    print("ЗАПУСК ТЕСТОВ (TC-01..TC-50)")
    print("=" * 60)
    print()

    excel = None
    workbook = None
    exit_code = 0

    try:
        write_log("[1/5] Создание COM-объекта Excel...")
        # Используем DispatchEx, чтобы создать НОВЫЙ выделенный экземпляр Excel
        # вместо повторного подключения к зависшему/скрытому EXCEL.EXE,
        # который вызывает DISP_E_EXCEPTION (0x80020009) на COM-вызовах.
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            write_log("    DispatchEx OK (новый выделенный экземпляр)")
        except Exception:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            write_log("    DispatchEx недоступен, использую EnsureDispatch")
        excel.Visible = False
        excel.DisplayAlerts = False
        # ВАЖНО: НЕ устанавливаем AutomationSecurity = msoAutomationSecurityForceDisable
        # (3): этот скрипт ВЫПОЛНЯЕТ макросы (RunAllTests), поэтому принудительное
        # отключение макросов неприменимо. work.xlsm проходит по доверенным путям,
        # поэтому безопасность макросов оставляем по умолчанию.

        # Записываем PID созданного экземпляра Excel для безопасного завершения
        # зависшего процесса конвейером build_all.py (только наш процесс).
        write_excel_pid(excel)

        write_log("[2/5] Открытие work.xlsm...")
        workbook = open_workbook_with_retry(excel, EXCEL_PATH)
        if workbook is None:
            raise RuntimeError("Workbooks.Open вернул None после всех ретраев")

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

            # Парсим статистику из ячейки Z1 (Total/Passed/Failed/Skipped)
            summary = parse_summary_from_z1(results)
            total = summary["total"]
            passed = summary["passed"]
            failed = summary["failed"]
            skipped = summary["skipped"]

            print()
            print("=" * 60)
            print(f"  Всего:      {total}")
            print(f"  Пройдено:   {passed}")
            print(f"  Провалено:  {failed}")
            print(f"  Пропущено:  {skipped}")
            print("=" * 60)
            print()

            if failed > 0:
                print("  [FAIL] ОБНАРУЖЕНЫ ОШИБКИ! Проверьте логи для деталей.")
                exit_code = 1
            else:
                print("  [OK] Все тесты успешно пройдены!")
                exit_code = 0

            write_log(f"Итог: Total={total}, Passed={passed}, Failed={failed}, Skipped={skipped}")
        else:
            print("    [ПРЕДУПРЕЖДЕНИЕ] Не удалось получить результаты через GetTestResults.")
            print("    Проверьте Immediate Window в редакторе VBA для деталей.")
            write_log("Результаты не получены (GetTestResults недоступна)")
            exit_code = 1
            summary = {"total": 0, "passed": 0, "failed": 0, "skipped": 0}

        # ===== R-16: отчёты JSON/HTML (только при запросе флагов) =====
        tests = []
        try:
            tests = parse_test_log(LOG_FILE)
        except Exception as exc:
            write_log(f"    [ПРЕДУПРЕЖДЕНИЕ] Ошибка чтения лога для отчётов: {exc}")

        filtered = apply_module_filter(tests, args.module)
        session_meta = parse_session_meta(LOG_FILE)

        if args.verbose:
            print()
            print("ПОДРОБНЫЙ РЕЗУЛЬТАТ ТЕСТОВ:")
            print_verbose(filtered)

        needs_report = (args.module is not None) or args.verbose or (args.output is not None)
        if needs_report:
            try:
                report = build_report(tests, summary, args, session_meta)
                out_dir = Path(args.output) if args.output else Path(LOGS_DIR)
                out_dir.mkdir(parents=True, exist_ok=True)
                save_json_report(report, out_dir)
                save_html_report(report, out_dir)
                write_log(f"Отчёты сохранены: {out_dir}")
            except Exception as exc:
                write_log(f"    [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить отчёты: {exc}")

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

        # Освобождаем COM-ссылки и вызываем сборщик мусора,
        # чтобы гарантированно завершить процесс EXCEL.EXE
        workbook = None
        excel = None
        gc.collect()

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