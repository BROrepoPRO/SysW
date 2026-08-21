#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общая конфигурация путей для Python-скриптов SysW.
Все скрипты должны импортировать пути из этого модуля,
чтобы избежать дублирования абсолютных путей.
"""
from pathlib import Path

# Версия приложения (единый источник для всей системы)
APP_VERSION = "1.0.3"

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