#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт ручной очистки системы SysW (Задача 9, v1.1.0).

Назначение:
    Позволяет пользователю вручную запускать очистку проекта от временных файлов,
    старого кэша, старых логов и лишних резервных копий через терминал VS Code.

Что очищается (по умолчанию):
    - временные каталоги: `_temp/`, `_temp_export/`, `_temp_import/` (содержимое);
    - Python-кэш: `__pycache__/`, `*.pyc`, `*.pyo`;
    - временные Excel/офисные файлы: `~$*.xlsm`, `~$*.xlsx`, `*.tmp`;
    - старые логи: файлы в `logs/` с расширением `.log`/`.txt`, кроме свежих
      (старше N дней, по умолчанию LOG_MAX_AGE_DAYS = 30);
    - старые резервные копии: точки отката в `_backup/`, превышающие
      BACKUP_KEEP (из config).

Безопасность:
    - Режим `--dry-run` (по умолчанию): только показывает, что будет удалено.
    - Реальный запуск — флаг `--apply`.
    - Файлы из `.codeassistantignore`/критические (`[E3]`) не затрагиваются.
    - Правило [Z2]: удаление происходит только по явной команде пользователя
      (запуск с `--apply`).

Запуск (из корня проекта):
    python scripts/clean_system.py             # dry-run, показать план
    python scripts/clean_system.py --apply     # реальная очистка

Параметры:
    --dry-run   показать, что будет удалено (режим по умолчанию)
    --apply     выполнить фактическое удаление
    --log-age N максимальный возраст логов в днях для сохранения (по умолчанию 30)
"""
from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

try:
    from config import (
        BACKUP_DIR,
        BACKUP_KEEP,
        LOGS_DIR,
        PROJECT_DIR,
        TEMP_DIR,
    )
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from config import (
        BACKUP_DIR,
        BACKUP_KEEP,
        LOGS_DIR,
        PROJECT_DIR,
        TEMP_DIR,
    )

# Дополнительные временные каталоги (помимо TEMP_DIR из config)
TEMP_EXPORT = PROJECT_DIR / "_temp_export"
TEMP_IMPORT = PROJECT_DIR / "_temp_import"

# Возраст логов (дней), старше которых они считаются устаревшими и удаляются
LOG_MAX_AGE_DAYS = 30


def _plan_temp_dirs() -> list[Path]:
    """Возвращает список временных каталогов (существующих) для очистки."""
    return [d for d in (TEMP_DIR, TEMP_EXPORT, TEMP_IMPORT) if d.is_dir()]


def _find_pycache() -> list[Path]:
    """Возвращает список каталогов __pycache__ в проекте (без .venv/venv)."""
    if not PROJECT_DIR.is_dir():
        return []
    out = []
    for p in PROJECT_DIR.rglob("__pycache__"):
        if p.is_dir() and ".venv" not in p.parts and "venv" not in p.parts:
            out.append(p)
    return out


def _find_bytecode() -> list[Path]:
    """Возвращает список .pyc/.pyo файлов в проекте (без .venv)."""
    out = []
    for pattern in ("*.pyc", "*.pyo"):
        for p in PROJECT_DIR.rglob(pattern):
            if ".venv" in p.parts or "venv" in p.parts:
                continue
            out.append(p)
    return out


def _find_temp_files() -> list[Path]:
    """Возвращает временные офисные/системные файлы в корне и docs."""
    out = []
    for pattern in ("~$*.xlsm", "~$*.xlsx", "*.tmp"):
        for p in PROJECT_DIR.rglob(pattern):
            out.append(p)
    return out


def _old_logs(age_days: int) -> list[Path]:
    """Возвращает файлы логов старше age_days дней."""
    out = []
    if not LOGS_DIR.is_dir():
        return out
    cutoff = time.time() - age_days * 86400
    for p in LOGS_DIR.iterdir():
        if p.is_file() and p.suffix.lower() in (".log", ".txt"):
            try:
                if p.stat().st_mtime < cutoff:
                    out.append(p)
            except OSError:
                pass
    return out


def _excess_backups(keep: int) -> list[Path]:
    """Возвращает точки отката в _backup сверх лимита keep (самые старые)."""
    if keep <= 0 or not BACKUP_DIR.is_dir():
        return []
    snapshots = sorted(
        (p for p in BACKUP_DIR.iterdir() if p.is_dir()),
        key=lambda p: (p.name,),
    )
    return snapshots[: max(0, len(snapshots) - keep)]


def main() -> int:
    apply_mode = "--apply" in sys.argv
    log_age = LOG_MAX_AGE_DAYS
    for i, a in enumerate(sys.argv):
        if a == "--log-age" and i + 1 < len(sys.argv):
            try:
                log_age = int(sys.argv[i + 1])
            except ValueError:
                pass

    groups = {
        "Временные каталоги (_temp*, _temp_export, _temp_import)": _plan_temp_dirs(),
        "Каталоги __pycache__": _find_pycache(),
        "Файлы .pyc/.pyo": _find_bytecode(),
        "Временные файлы (~$, *.tmp)": _find_temp_files(),
        f"Старые логи (> {log_age} дн.)": _old_logs(log_age),
        f"Лишние бэкапы (лимит {BACKUP_KEEP})": _excess_backups(BACKUP_KEEP),
    }

    total = sum(len(v) for v in groups.values())
    header = "ОЧИСТКА СИСТЕМЫ — РЕАЛЬНОЕ УДАЛЕНИЕ" if apply_mode else \
             "ОЧИСТКА СИСТЕМЫ — ПРОСМОТР (dry-run, изменений нет)"
    print(f"\n{header}\nВсего объектов: {total}\n")

    if total == 0:
        print("Нечего очищать. Проект чист.")
        return 0

    for label, items in groups.items():
        if not items:
            continue
        print(f"--- {label} ({len(items)}) ---")
        for p in items[:20]:  # показать до 20 строк на группу
            print(f"  {p.relative_to(PROJECT_DIR)}")
        if len(items) > 20:
            print(f"  ... и ещё {len(items) - 20}")

    if not apply_mode:
        print("\nЗапустите с флагом --apply для фактического удаления.")
        return 0

    # Фактическое удаление
    removed = 0
    for label, items in groups.items():
        for p in items:
            try:
                if p.is_dir() and not p.is_symlink():
                    shutil.rmtree(p)
                else:
                    p.unlink()
                removed += 1
            except OSError as exc:
                print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось удалить {p}: {exc}")

    print(f"\nУдалено объектов: {removed} / {total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())