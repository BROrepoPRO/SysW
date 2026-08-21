#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый конвейер сборки проекта SysW (Задача 4 плана v1.0.7).

Назначение:
    Выполняет полную пересборку проекта в строго определённом порядке,
    запуская существующие скрипты подпроцессами (без дублирования их логики).

Этапы (в этом порядке):
    1. Резервное копирование — копии work.xlsm и SysW.db в _backup/
       с меткой времени (work_YYYYmmdd_HHMMSS.xlsm, SysW_YYYYmmdd_HHMMSS.db).
    2. impVBA.py                — импорт исходников VBA из src/ в work.xlsm.
    3. build_templates.py       — пересборка шаблонов base/templates/.
    4. migrate_models_to_sqlite.py — пересборка SysW.db из base/models/*.
    5. Контроль целостности БД — PRAGMA integrity_check ("ok") и контрольные
       количества по таблицам works, parts_catalog, parts, matlib_entries.
    6. run_tests.py             — прогон тестов (после успешных этапов 2-5).

Логирование:
    Прогресс каждого этапа (старт/успех/ошибка) пишется в logs/build.log
    (append). Дублируется в консоль.

Exit code скрипта (схема кодов, отражает конкретный проваленный этап):
    0 - полный успех всех этапов;
    1 - этап 1 (резервное копирование) провален;
    2 - этап 2 (impVBA.py) провален;
    3 - этап 3 (build_templates.py) провален;
    4 - этап 4 (migrate_models_to_sqlite.py) провален;
    5 - этап 5 (контроль целостности БД) провален;
    6 - этап 6 (run_tests.py) провален;
    70+ - внутренняя ошибка самого конвейера (сеть, отсутствие файла и т.п.).

Поведение при ошибке:
    При провале ЛЮБОГО этапа конвейер останавливается (последующие этапы не
    выполняются), финальный exit code ненулевой и соответствует этапу.

Запуск (из корня проекта):
    python scripts/build_all.py

Ограничения:
    Данный скрипт НЕ изменяет существующие скрипты и бинарные файлы;
    вся работа выполняется через запуск их подпроцессами.
"""
from __future__ import annotations

import shutil
import sqlite3
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# Пути и константы из общей конфигурации (используются для согласованности
# с остальными скриптами системы).
try:
    from config import (
        DB_PATH,
        LOGS_DIR,
        PROJECT_DIR,
        WORKBOOK_PATH,
        APP_VERSION,
    )
except ImportError:  # pragma: no cover - запуск из другой директории
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from config import (  # type: ignore
        DB_PATH,
        LOGS_DIR,
        PROJECT_DIR,
        WORKBOOK_PATH,
        APP_VERSION,
    )

# ---------------------------------------------------------------------------
# Константы конвейера
# ---------------------------------------------------------------------------

# Файл лога сборки (в директории logs/)
BUILD_LOG_FILE = LOGS_DIR / "build.log"

# Каталог резервных копий (корень проекта)
BACKUP_DIR = PROJECT_DIR / "_backup"

# Файл лога тестов (используется run_tests.py, путь берётся из config)
TEST_LOG_FILE = LOGS_DIR / "test_results.log"

# Схема exit-кодов конвейера (ключ этапа -> код при его провале).
# Значение 0 используется только при полном успехе.
EXIT_CODES = {
    "backup": 1,
    "impvba": 2,
    "templates": 3,
    "migrate": 4,
    "integrity": 5,
    "tests": 6,
}

# Таблицы и их описание для контрольной проверки целостности (этап 5).
# Каждая запись: (имя_таблицы, человекочитаемое описание, критичность).
# Критичные таблицы обязаны иметь ненулевое количество строк.
INTEGRITY_TABLES = [
    ("works", "работы", True),
    ("parts_catalog", "каталог запчастей", True),
    ("parts", "привязки запчастей по группам", False),
    ("matlib_entries", "материалы библиотеки", False),
]


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def ensure_logs_dir() -> None:
    """Создаёт директорию logs/, если её нет."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)


def log_line(message: str) -> None:
    """Печать в консоль и добавление строки с меткой времени в build.log."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    with open(BUILD_LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def timestamp_label() -> str:
    """Возвращает метку времени для имени резервной копии (YYYYmmdd_HHMMSS)."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def run_step_script(name: str, script_name: str) -> bool:
    """Запускает существующий скрипт проекта подпроцессом и возвращает успех.

    Вызывается `python scripts/<script_name>` из корня проекта. stdout/stderr
    подпроцесса транслируются в консоль, чтобы пользователь видел живой вывод.
    """
    script_path = PROJECT_DIR / "scripts" / script_name
    if not script_path.is_file():
        log_line(f"[ОШИБКА] Не найден скрипт: {script_path}")
        return False

    log_line(f"[СТАРТ] Этап '{name}': python scripts/{script_name}")
    try:
        proc = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(PROJECT_DIR),
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except OSError as exc:
        log_line(f"[ОШИБКА] Не удалось запустить этап '{name}': {exc}")
        return False

    if proc.returncode == 0:
        log_line(f"[УСПЕХ] Этап '{name}' завершён (код 0)")
        return True

    log_line(
        f"[ОШИБКА] Этап '{name}' завершился с кодом {proc.returncode}"
    )
    return False


def copy_with_backup(src: Path, dst: Path) -> bool:
    """Копирует одиночный файл, логируя ошибки. Возвращает успех."""
    try:
        shutil.copy2(src, dst)
    except OSError as exc:
        log_line(f"[ОШИБКА] Не удалось скопировать {src.name}: {exc}")
        return False
    log_line(f"    Скопирован {src.name} -> {dst}")
    return True


# ---------------------------------------------------------------------------
# Отдельные этапы конвейера
# ---------------------------------------------------------------------------

def step_backup() -> bool:
    """Этап 1: резервное копирование work.xlsm и SysW.db в _backup/.

    Файлы копируются с меткой времени в имени. Если какой-либо из файлов
    отсутствует на диске, копируется только существующий, но этап считается
    успешным только при копировании хотя бы одного файла.
    """
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = timestamp_label()
    ok = False

    # Резервная копия корневого work.xlsm
    if WORKBOOK_PATH.is_file():
        dst = BACKUP_DIR / f"work_{stamp}.xlsm"
        ok = copy_with_backup(WORKBOOK_PATH, dst) or ok
    else:
        log_line(f"[ПРЕДУПРЕЖДЕНИЕ] Файл не найден, пропуск: {WORKBOOK_PATH}")

    # Резервная копия базы SysW.db (путь из config, корень проекта)
    if DB_PATH.is_file():
        dst = BACKUP_DIR / f"SysW_{stamp}.db"
        ok = copy_with_backup(DB_PATH, dst) or ok
    else:
        log_line(f"[ПРЕДУПРЕЖДЕНИЕ] Файл не найден, пропуск: {DB_PATH}")

    if not ok:
        log_line("[ОШИБКА] Этап 'backup': не создано ни одной резервной копии")
        return False
    return True


def step_integrity() -> bool:
    """Этап 5: контроль целостности базы SysW.db.

    Выполняет PRAGMA integrity_check (ожидается "ok") и контрольные количества
    по таблицам works, parts_catalog, parts, matlib_entries.
    Считается проваленным, если:
      - файл БД отсутствует;
      - integrity_check вернул не "ok";
      - критичная таблица отсутствует или имеет нулевое количество строк.
    """
    if not DB_PATH.is_file():
        log_line(f"[ОШИБКА] Файл БД не найден: {DB_PATH}")
        return False

    try:
        conn = sqlite3.connect(str(DB_PATH))
    except sqlite3.Error as exc:
        log_line(f"[ОШИБКА] Не удалось открыть БД: {exc}")
        return False

    try:
        # 1) Проверка целостности
        row = conn.execute("PRAGMA integrity_check").fetchone()
        result = row[0] if row else ""
        if result != "ok":
            log_line(f"[ОШИБКА] integrity_check вернул '{result}' (ожидалось 'ok')")
            return False
        log_line(f"    PRAGMA integrity_check = '{result}'")

        # 2) Контрольные количества по таблицам
        failed = False
        for table, description, critical in INTEGRITY_TABLES:
            try:
                count_row = conn.execute(
                    f"SELECT COUNT(*) FROM {table}"
                ).fetchone()
            except sqlite3.Error as exc:
                log_line(
                    f"[ОШИБКА] Таблица {table} ({description}) недоступна: {exc}"
                )
                failed = True
                continue

            count = count_row[0] if count_row else 0
            status = "OK" if (not critical or count > 0) else "НУЛЕВАЯ/ОТСУТСТВУЕТ"
            log_line(
                f"    {table} ({description}): {count} строк -> [{status}]"
            )
            if critical and count <= 0:
                failed = True

        if failed:
            log_line("[ОШИБКА] Этап 'integrity': контрольные количества некорректны")
            return False
        return True
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Главная функция конвейера
# ---------------------------------------------------------------------------

def main() -> int:
    """Выполняет все этапы сборки по порядку. Возвращает exit code конвейера."""
    ensure_logs_dir()
    log_line("=" * 70)
    log_line(f"НАЧАЛО СБОРКИ SysW (версия конфигурации {APP_VERSION})")
    log_line("=" * 70)

    # --- Этап 1: резервное копирование --------------------------------
    if not step_backup():
        return EXIT_CODES["backup"]

    # --- Этап 2: импорт VBA (impVBA.py) -------------------------------
    if not run_step_script("impVBA", "impVBA.py"):
        return EXIT_CODES["impvba"]

    # --- Этап 3: пересборка шаблонов (build_templates.py) --------------
    if not run_step_script("build_templates", "build_templates.py"):
        return EXIT_CODES["templates"]

    # --- Этап 4: пересборка БД (migrate_models_to_sqlite.py) -----------
    if not run_step_script("migrate_models_to_sqlite", "migrate_models_to_sqlite.py"):
        return EXIT_CODES["migrate"]

    # --- Этап 5: контроль целостности БД -------------------------------
    log_line("[СТАРТ] Этап 'integrity': проверка целостности SysW.db")
    if not step_integrity():
        return EXIT_CODES["integrity"]
    log_line("[УСПЕХ] Этап 'integrity' завершён")

    # --- Этап 6: прогон тестов (run_tests.py) --------------------------
    if not run_step_script("run_tests", "run_tests.py"):
        return EXIT_CODES["tests"]

    log_line("=" * 70)
    log_line("СБОРКА ЗАВЕРШЕНА УСПЕШНО")
    log_line("=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())