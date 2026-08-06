#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Инструмент проверки согласованности документации проекта SysW.

Скрипт проверяет, что документация (docs/*.md, README.md) не устарела
относительно фактической структуры проекта, и выводит отчёт о расхождениях.

Режимы работы:
    --check  — проверка, вывод отчёта. Код возврата 0, если всё согласовано,
               1, если найдены расхождения.
    --fix    — автоматическое исправление простых расхождений
               (например, замена ссылок ARCHITECTURE_SQLITE.md на ARCHITECTURE.md).

Совместим с Python 3. Использует только стандартную библиотеку.
"""

import argparse
import re
import sys
from pathlib import Path

# Импорт версии из единого источника (config.py), чтобы не дублировать её
from config import APP_VERSION

# Корень проекта — родительская директория scripts/
PROJECT_DIR = Path(__file__).resolve().parent.parent

# Актуальная версия проекта (берётся из единого источника config.APP_VERSION)
CURRENT_VERSION = f"v{APP_VERSION}"

# Директории и файлы, участвующие в проверке
DOCS_DIR = PROJECT_DIR / "docs"
README_FILE = PROJECT_DIR / "README.md"
SRC_MODULES_DIR = PROJECT_DIR / "src" / "modules"
SRC_CLASSES_DIR = PROJECT_DIR / "src" / "classes"
SRC_SHEETS_DIR = PROJECT_DIR / "src" / "sheets"
SCRIPTS_DIR = PROJECT_DIR / "scripts"

# Файлы документации, которые проверяем
DOC_FILES = [
    "ARCHITECTURE.md",
    "CHANGELOG.md",
    "DEVELOPER.md",
    "git-workflow.md",
    "ROADMAP.md",
    "sourcecraft-guide.md",
    "table.md",
]

# Модули, которые были удалены из проекта и не должны упоминаться в документации
REMOVED_MODULES = ["Mod_MainButtons"]

# Модули, которые должны упоминаться в документации (новые/актуальные)
EXPECTED_MODULES = ["Mod_ModelTypes"]

# Сопоставление старых имён файлов документации на новые (для автозамены)
RENAMED_DOCS = {
    "ARCHITECTURE_SQLITE.md": "ARCHITECTURE.md",
}

# Список скриптов, которые должны вызываться с префиксом scripts/
SCRIPTS = [
    "config.ps1",
    "config.py",
    "export_vba.py",
    "impVBA.py",
    "run_tests.py",
    "Set-ExcelTrust.ps1",
]


def collect_md_files():
    """Собирает список .md-файлов для проверки (docs/*.md и README.md)."""
    files = []
    for name in DOC_FILES:
        path = DOCS_DIR / name
        if path.exists():
            files.append(path)
    if README_FILE.exists():
        files.append(README_FILE)
    return files


def read_text(path):
    """Читает текстовый файл в кодировке UTF-8 (с запасным вариантом)."""
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Запасной вариант — читаем в cp1251, если UTF-8 не подошёл
        return path.read_text(encoding="cp1251")


def check_removed_modules(files, issues):
    """Проверяет, что удалённые модули не упоминаются в документации."""
    for path in files:
        # В CHANGELOG.md удалённые модули упоминаются в исторических записях
        # (разделы Removed/Changed/Fixed прошлых версий) — это нормальная практика,
        # поэтому такие упоминания не считаем расхождением.
        if path.name == "CHANGELOG.md":
            continue
        content = read_text(path)
        for module in REMOVED_MODULES:
            # Ищем упоминание модуля как отдельного слова
            pattern = re.compile(r"\b" + re.escape(module) + r"\b")
            if pattern.search(content):
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"упоминается удалённый модуль '{module}'"
                )


def check_expected_modules(files, issues):
    """Проверяет, что новые модули упоминаются в документации, где ожидается."""
    # Ожидаем упоминание новых модулей в ARCHITECTURE.md и DEVELOPER.md
    expected_files = ["ARCHITECTURE.md", "DEVELOPER.md"]
    for name in expected_files:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        content = read_text(path)
        for module in EXPECTED_MODULES:
            pattern = re.compile(r"\b" + re.escape(module) + r"\b")
            if not pattern.search(content):
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"не упоминается новый модуль '{module}'"
                )


def check_cross_references(files, issues):
    """Проверяет, что ссылки на файлы документации указывают на существующие файлы."""
    # Собираем множество существующих .md-файлов в docs/ и корне
    existing_md = set()
    for path in DOCS_DIR.glob("*.md"):
        existing_md.add(path.name)
    if README_FILE.exists():
        existing_md.add(README_FILE.name)

    # Регулярное выражение для markdown-ссылок [текст](путь)
    link_pattern = re.compile(r"\[[^\]]*\]\(([^)]+)\)")

    for path in files:
        content = read_text(path)
        for match in link_pattern.finditer(content):
            target = match.group(1).strip()
            # Пропускаем внешние ссылки и якоря
            if target.startswith(("http://", "https://", "#", "mailto:")):
                continue
            # Пропускаем ссылки на директории
            if target.endswith("/"):
                continue
            # Нормализуем путь: убираем якорь и query
            clean_target = target.split("#")[0].split("?")[0]
            if not clean_target:
                continue
            # Проверяем только ссылки на .md-файлы
            if not clean_target.lower().endswith(".md"):
                continue
            # Ссылки внутри файлов docs/ указываются относительно самой директории docs/,
            # а ссылки с префиксом ../ — относительно корня проекта. Поэтому резолвим
            # путь относительно директории файла, в котором находится ссылка.
            abs_target = (path.parent / clean_target).resolve()
            if not abs_target.exists():
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"ссылка на несуществующий файл: '{target}'"
                )


def check_script_paths(files, issues):
    """Проверяет, что команды в документации используют корректные пути к скриптам."""
    # Ищем вызовы скриптов без префикса scripts/
    # Например: "python export_vba.py" вместо "python scripts/export_vba.py"
    for path in files:
        content = read_text(path)
        for script in SCRIPTS:
            # Ищем вызов скрипта без префикса scripts/
            # Паттерн: python <script> или python3 <script> без "scripts/"
            pattern = re.compile(
                r"\b(?:python|python3)\s+(?!scripts/)" + re.escape(script)
            )
            if pattern.search(content):
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"команда использует путь без префикса scripts/: '{script}'"
                )


def check_versions(files, issues):
    """Проверяет, что версия проекта в документации соответствует актуальной."""
    # Ищем упоминания версий вида vX.Y.Z в документации
    version_pattern = re.compile(r"\bv\d+\.\d+\.\d+\b")
    for path in files:
        content = read_text(path)
        # В исторических документах (CHANGELOG.md, ROADMAP.md) перечислены все
        # прошлые версии (v0.2.0 ... v0.18.0) — это нормальная практика, поэтому
        # такие упоминания не считаем расхождением. Проверяем только актуальную
        # версию (например, в заголовке или блоке «Версия системы»).
        if path.name in ("CHANGELOG.md", "ROADMAP.md"):
            if CURRENT_VERSION not in content:
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"не найдена актуальная версия {CURRENT_VERSION}"
                )
            continue
        for match in version_pattern.finditer(content):
            version = match.group(0)
            if version != CURRENT_VERSION:
                issues.append(
                    f"[{path.relative_to(PROJECT_DIR)}] "
                    f"устаревшая версия '{version}' (актуальная {CURRENT_VERSION})"
                )


def fix_cross_references(files, fixed):
    """Автоматически исправляет ссылки на переименованные файлы документации."""
    for path in files:
        content = read_text(path)
        new_content = content
        for old_name, new_name in RENAMED_DOCS.items():
            if old_name in new_content:
                new_content = new_content.replace(old_name, new_name)
        if new_content != content:
            path.write_text(new_content, encoding="utf-8")
            fixed.append(
                f"[{path.relative_to(PROJECT_DIR)}] "
                f"заменено '{list(RENAMED_DOCS.keys())[0]}' "
                f"на '{list(RENAMED_DOCS.values())[0]}'"
            )


def main():
    """Точка входа скрипта."""
    parser = argparse.ArgumentParser(
        description="Проверка согласованности документации проекта SysW"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--check",
        action="store_true",
        help="проверка документации и вывод отчёта о расхождениях",
    )
    group.add_argument(
        "--fix",
        action="store_true",
        help="автоматическое исправление простых расхождений",
    )
    args = parser.parse_args()

    # Собираем файлы для проверки
    files = collect_md_files()
    if not files:
        print("ОШИБКА: не найдены файлы документации для проверки.")
        return 1

    issues = []
    fixed = []

    if args.check:
        print("=== Проверка согласованности документации SysW ===")
        print(f"Актуальная версия: {CURRENT_VERSION}")
        print(f"Проверяемых файлов: {len(files)}")
        print()

        # Выполняем все проверки
        check_removed_modules(files, issues)
        check_expected_modules(files, issues)
        check_cross_references(files, issues)
        check_script_paths(files, issues)
        check_versions(files, issues)

        # Выводим отчёт
        if issues:
            print(f"Найдено расхождений: {len(issues)}")
            print()
            for issue in issues:
                print(f"  [РАСХОЖДЕНИЕ] {issue}")
            print()
            print("Для автоматического исправления простых расхождений "
                  "запустите: python scripts/check_docs.py --fix")
            return 1
        else:
            print("Расхождений не найдено. Документация согласована.")
            return 0

    elif args.fix:
        print("=== Автоматическое исправление расхождений ===")
        fix_cross_references(files, fixed)

        if fixed:
            print(f"Исправлено: {len(fixed)}")
            print()
            for item in fixed:
                print(f"  [ИСПРАВЛЕНО] {item}")
        else:
            print("Простых расхождений для автоматического исправления не найдено.")
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())