#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общая конфигурация путей для Python-скриптов SysW.
Все скрипты должны импортировать пути из этого модуля,
чтобы избежать дублирования абсолютных путей.
"""
from pathlib import Path

# Версия приложения (единый источник для всей системы)
APP_VERSION = "1.0.10"

# Корень проекта — родительская директория scripts/
PROJECT_DIR = Path(__file__).resolve().parent.parent

# Основной Excel-файл
WORKBOOK_PATH = PROJECT_DIR / "work.xlsm"

# Директория исходников VBA
SRC_DIR = PROJECT_DIR / "src"

# Временные директории
TEMP_EXPORT_DIR = PROJECT_DIR / "_temp_export"
TEMP_IMPORT_DIR = PROJECT_DIR / "_temp_import"

# Директория логов
LOGS_DIR = PROJECT_DIR / "logs"

# Файл лога тестов (в директории logs/)
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

# Файл отчёта миграции (в директории logs/)
MIGRATION_REPORT_FILE = LOGS_DIR / "migration_report.log"