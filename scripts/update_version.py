#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт автоматического обновления версии проекта SysW.

При релизе обновляет версию во всех файлах, где она указана,
используя единый источник версии (config.APP_VERSION) как точку отсчёта.

Использование:
    python scripts/update_version.py 1.1.0

Скрипт обновляет версию в следующих файлах:
    - src/modules/Mod_Constants.bas   (константа APP_VERSION)
    - scripts/config.py               (переменная APP_VERSION)
    - scripts/config.ps1              (переменная $Script:AppVersion)
    - README.md                       (заголовок/бейдж версии)
    - docs/DEVELOPER.md               (заголовок)
    - docs/ROADMAP.md                 (блок «Версия системы»)
    - docs/ARCHITECTURE.md            (версия проекта)
    - docs/CHANGELOG.md               (новый раздел ## [vX.Y.Z] — дата)

Поддерживаемые форматы версии (v1.1.0, Задача 3):
    - X.Y.Z    (SemVer, например 1.1.0);
    - X.Y.Z.W  (расширенный, 4-я компонента W инкрементируется при изменениях
      todo/CHANGELOG/commits внутри одного промта, например 1.0.18.1).

Использует только стандартную библиотеку Python.
"""

import argparse
import re
import sys
from datetime import date
from pathlib import Path

# Корень проекта — родительская директория scripts/
PROJECT_DIR = Path(__file__).resolve().parent.parent

# Регулярное выражение для проверки формата версии (SemVer: X.Y.Z или X.Y.Z.W)
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+(\.\d+)?$")


def validate_version(version):
    """Проверяет, что версия соответствует формату X.Y.Z или X.Y.Z.W."""
    if not VERSION_RE.match(version):
        print(f"ОШИБКА: Некорректный формат версии '{version}'. "
              f"Ожидается формат X.Y.Z (например, 1.1.0) или X.Y.Z.W (например, 1.0.18.1).")
        sys.exit(1)


def replace_in_file(path, prefix, version):
    """
    Заменяет в файле первое вхождение старой версии (после заданного префикса)
    на новую, полностью удаляя остаток старого номера.

    Использует регулярное выражение: ищет префикс + семантическую версию
    (X.Y.Z или X.Y.Z.W), поэтому не оставляет «хвостов» вида 1.0.21.0.0
    независимо от длины старой и новой версий.
    Возвращает True, если замена выполнена, иначе False.
    """
    if not path.exists():
        print(f"  ПРОПУЩЕН: файл не найден — {path}")
        return False

    text = path.read_text(encoding="utf-8")

    # Паттерн: экранированный префикс + полный номер версии (X.Y.Z или X.Y.Z.W).
    # Захватываем весь номер, чтобы при замене не оставался хвост старой версии.
    pattern = re.escape(prefix) + r"(\d+\.\d+\.\d+(\.\d+)?)"
    new_text, count = re.subn(pattern, prefix + version, text, count=1)

    if count == 0:
        print(f"  ПРОПУЩЕН: шаблон '{prefix}<версия>' не найден — {path}")
        return False

    path.write_text(new_text, encoding="utf-8")
    print(f"  ОБНОВЛЕНО: {path}  ({prefix}<старая версия> -> {version})")
    return True


def update_vba_const(path, version):
    """Обновляет константу APP_VERSION в VBA-модуле Mod_Constants.bas."""
    return replace_in_file(
        path,
        'Public Const APP_VERSION As String = "',
        version,
    )


def update_python_config(path, version):
    """Обновляет переменную APP_VERSION в scripts/config.py."""
    return replace_in_file(
        path,
        'APP_VERSION = "',
        version,
    )


def update_powershell_config(path, version):
    """Обновляет переменную $Script:AppVersion в scripts/config.ps1."""
    return replace_in_file(
        path,
        '$Script:AppVersion = "',
        version,
    )


def update_readme(path, version):
    """Обновляет версию в заголовке README.md."""
    return replace_in_file(
        path,
        "**Версия:** v",
        version,
    )


def update_developer(path, version):
    """Обновляет версию в заголовке docs/DEVELOPER.md."""
    return replace_in_file(
        path,
        "SysW (v",
        version,
    )


def update_roadmap(path, version):
    """Обновляет версию в блоке «Версия системы» docs/ROADMAP.md."""
    return replace_in_file(
        path,
        "**Версия системы:** ",
        version,
    )


def update_architecture(path, version):
    """Обновляет версию проекта в docs/ARCHITECTURE.md."""
    return replace_in_file(
        path,
        "Проект: SysW v",
        version,
    )


def update_changelog(path, version):
    """
    Добавляет новый раздел ## [vX.Y.Z] — дата в начало docs/CHANGELOG.md
    (сразу после вводного блока, перед текущим первым разделом).
    """
    if not path.exists():
        print(f"  ПРОПУЩЕН: файл не найден — {path}")
        return False

    text = path.read_text(encoding="utf-8")
    today = date.today().isoformat()
    new_section = f"## [v{version}] — {today}\n\n### Added\n- **Единый источник версии и скрипт `update_version.py`:** версия проекта централизована в `config.APP_VERSION`, `Mod_Constants.APP_VERSION` и `$Script:AppVersion`; добавлен скрипт автоматического обновления версии.\n\n"

    # Вставляем новый раздел после строки "версионирование следует ..." (конец вводного блока)
    marker = "версионирование следует [Semantic Versioning](https://semver.org/lang/ru/).\n\n"
    if marker in text:
        new_text = text.replace(marker, marker + new_section, 1)
    else:
        # Запасной вариант: вставляем после первого заголовка "# История изменений"
        new_text = text.replace(
            "# История изменений\n\n",
            "# История изменений\n\n" + new_section,
            1,
        )

    path.write_text(new_text, encoding="utf-8")
    print(f"  ОБНОВЛЕНО: {path}  (добавлен раздел ## [v{version}] — {today})")
    return True


def main():
    """Главная функция скрипта."""
    parser = argparse.ArgumentParser(
        description="Автоматическое обновление версии проекта SysW."
    )
    parser.add_argument(
        "version",
        help="Новая версия в формате X.Y.Z (например, 1.1.0) или X.Y.Z.W (например, 1.0.18.1).",
    )
    args = parser.parse_args()

    new_version = args.version.strip()
    validate_version(new_version)

    print(f"Обновление версии проекта до v{new_version}...")
    print("=" * 60)

    # Список задач обновления: (описание, функция, путь)
    tasks = [
        ("VBA-константа APP_VERSION",
         update_vba_const,
         PROJECT_DIR / "src" / "modules" / "Mod_Constants.bas"),
        ("Python-переменная APP_VERSION",
         update_python_config,
         PROJECT_DIR / "scripts" / "config.py"),
        ("PowerShell-переменная $Script:AppVersion",
         update_powershell_config,
         PROJECT_DIR / "scripts" / "config.ps1"),
        ("README.md (заголовок)",
         update_readme,
         PROJECT_DIR / "README.md"),
        ("docs/DEVELOPER.md (заголовок)",
         update_developer,
         PROJECT_DIR / "docs" / "DEVELOPER.md"),
        ("docs/ROADMAP.md (блок «Версия системы»)",
         update_roadmap,
         PROJECT_DIR / "docs" / "ROADMAP.md"),
        ("docs/ARCHITECTURE.md (версия проекта)",
         update_architecture,
         PROJECT_DIR / "docs" / "ARCHITECTURE.md"),
        ("docs/CHANGELOG.md (новый раздел)",
         update_changelog,
         PROJECT_DIR / "docs" / "CHANGELOG.md"),
    ]

    updated_count = 0
    for description, func, path in tasks:
        print(f"[{description}]")
        if func(path, new_version):
            updated_count += 1

    print("=" * 60)
    print(f"Готово. Обновлено файлов: {updated_count} из {len(tasks)}.")
    print(f"Новая версия проекта: v{new_version}")


if __name__ == "__main__":
    main()