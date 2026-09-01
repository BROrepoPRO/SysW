#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общая конфигурация путей для Python-скриптов SysW.
Все скрипты должны импортировать пути из этого модуля,
чтобы избежать дублирования абсолютных путей.
"""
from pathlib import Path

# Версия приложения (единый источник для всей системы)
APP_VERSION = "1.1.6.2"

# Корень проекта — родительская директория scripts/
PROJECT_DIR = Path(__file__).resolve().parent.parent

# Основной Excel-файл
WORKBOOK_PATH = PROJECT_DIR / "work.xlsm"

# Директория исходников VBA
SRC_DIR = PROJECT_DIR / "src"

# Временные директории
TEMP_EXPORT_DIR = PROJECT_DIR / "_temp_export"
TEMP_IMPORT_DIR = PROJECT_DIR / "_temp_import"

# Единая папка временных артефактов ассистента/скриптов (Задача 1, v1.0.17)
TEMP_DIR = PROJECT_DIR / "_temp"

# Директория логов
LOGS_DIR = PROJECT_DIR / "logs"

# Файл системного лога (в директории logs/) — общие системные события
LOG_FILE = LOGS_DIR / "log.txt"

# Файл лога тестов (в директории logs/) — расширенный тестовый лог
# (уровни INFO/WARN/ERROR; заполняется VBA Mod_Logger.WriteTestLog и run_tests.py)
TEST_LOG_FILE = LOGS_DIR / "test_results.log"

# ============================================================
# Единая база данных SQLite (миграция на SysW.db)
# ============================================================

# Версия схемы БД (должна совпадать с PRAGMA user_version в db/schema.sql)
DB_SCHEMA_VERSION = 1

# Путь к единой базе данных SysW.db (корневой уровень проекта)
DB_PATH = PROJECT_DIR / "SysW.db"

# Путь к файлу DDL-схемы
DB_SCHEMA_PATH = PROJECT_DIR / "db" / "schema.sql"

# Директория модельных файлов (легаси .xlsm)
MODELS_DIR = PROJECT_DIR / "base" / "models"

# Директория шаблонов (base/templates/)
TEMPLATES_DIR = PROJECT_DIR / "base" / "templates"

# Корневой файл отчёта Excel
REPORT_PATH = PROJECT_DIR / "report.xlsx"

# ============================================================
# Библиотека запретов редактирования (Задача 5, v1.0.17)
# ------------------------------------------------------------
# Перечень критических ключевых файлов, которые НЕЛЬЗЯ редактировать
# напрямую без резервной копии и согласования (см. правило [U5] в .ycarules).
# Единый источник для Python-скриптов сборки/проверки.
# ============================================================
PROTECTED_FILES: list[Path] = [
    WORKBOOK_PATH,   # work.xlsm
    REPORT_PATH,     # report.xlsx
    TEMPLATES_DIR,   # base/templates/* (все шаблоны)
    MODELS_DIR,      # base/models/* (все модельные книги)
    DB_PATH,         # SysW.db
]

# ============================================================
# Резервирование (Задача 10, v1.1.0)
# ------------------------------------------------------------
# Каталог резервных копий и количество хранимых «точек отката».
# При превышении лимита самая старая точка отката удаляется (ротация).
# ============================================================
BACKUP_DIR = PROJECT_DIR / "_backup"
BACKUP_KEEP = 5

# Файл отчёта миграции (в директории logs/)
MIGRATION_REPORT_FILE = LOGS_DIR / "migration_report.log"

# Файл отчёта инициации пользовательских модельных файлов (в директории logs/)
INITIATION_REPORT_FILE = LOGS_DIR / "initiation_report.log"