#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Статическая проверка синтаксиса VBA-исходников в src/ (без запуска Excel).

Назначение (v1.0.8, Задача 2):
    Ловить типовую ошибку компиляции VBA сразу после этапа impVBA.py и ДО
    пересборки шаблонов/прогона тестов: запрещённую inline-инициализацию
    модульных/глобальных переменных значением при объявлении.

    В VBA нельзя инициализировать переменную значением прямо в строке
    объявления на уровне модуля. Недопустимо:
        Public X As Boolean = True
        Dim Count As Long = 5
        Private Flag As Boolean = False
    Разрешено только:
        Const MAX As Long = 100        (константы)
        Public Y As Boolean            (объявление без инициализации)
        Optional ByVal silent As Boolean = False  (в сигнатуре процедуры)

    ВАЖНО: чекер срабатывает ТОЛЬКО на объявления модульных/глобальных
    переменных с инициализацией значением (строки уровня модуля вне процедур).
    Объявления внутри процедур (Sub/Function/Property) и легальные случаи
    (Const, Optional в сигнатурах) НЕ считаются ошибками.

Использование:
    python scripts/check_vba_syntax.py                # проверить src/
    python scripts/check_vba_syntax.py <путь>          # проверить файл/папку

Exit code:
    0 — ошибок не найдено;
    1 — найдены ошибки (в stdout выводится список: файл:строка:сообщение).
"""
import re
import sys
from pathlib import Path

# Корень проекта и каталог исходников по умолчанию.
PROJECT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_SRC_DIR = PROJECT_DIR / "src"

# Подкаталоги src/, содержащие VBA-исходники.
SRC_SUBDIRS = ["modules", "classes", "sheets"]

# Регулярные выражения.

# Открытие процедуры: Sub/Function/Property Get|Let|Set (с необязательными
# модификаторами Public/Private/Friend/Static).
_PROC_OPEN_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\b",
    re.IGNORECASE,
)

# Закрытие процедуры: End Sub / End Function / End Property.
_PROC_CLOSE_RE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)

# Объявление константы (разрешено инициализировать значением).
_CONST_RE = re.compile(r"^\s*(?:Public|Private)?\s*Const\b", re.IGNORECASE)

# Запрещённый паттерн: модульная/глобальная переменная с инициализацией.
# Срабатывает ТОЛЬКО на строках уровня модуля вне процедур.
_DECL_INIT_RE = re.compile(
    r"^\s*(?:Public|Private|Dim)\s+\w+\s+As\s+\w+\s*=",
    re.IGNORECASE,
)


def read_vba_text(path: Path):
    """Читает VBA-файл с автоопределением кодировки (UTF-8/CP1251).

    Возвращает содержимое как Unicode. При невозможности определить кодировку
    выбрасывает ValueError — такой файл считается ошибкой проверки.
    """
    raw = path.read_bytes()
    # UTF-8 с BOM
    if raw[:3] == b"\xef\xbb\xbf":
        return raw[3:].decode("utf-8")
    # UTF-8 без BOM
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass
    # Windows-1251 (fallback)
    return raw.decode("cp1251")


def check_file(path: Path):
    """Проверяет один VBA-файл, возвращает список найденных ошибок.

    Каждая ошибка — строка вида "путь:номер_строки:сообщение".
    """
    errors = []
    try:
        text = read_vba_text(path)
    except (UnicodeDecodeError, ValueError) as exc:
        errors.append(f"{path}:?:[ошибка чтения файла: {exc}]")
        return errors

    inside_proc = False  # Флаг «внутри процедуры».
    for lineno, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        # Обновляем состояние вложенности процедуры.
        if _PROC_CLOSE_RE.match(line):
            inside_proc = False
            continue
        if _PROC_OPEN_RE.match(line):
            inside_proc = True
            continue

        # Внутри процедуры проверки модульных объявлений не выполняем
        # (локальные объявления и Optional в сигнатурах не входят в критерий).
        if inside_proc:
            continue

        # Константы разрешено инициализировать значением — пропускаем.
        if _CONST_RE.match(line):
            continue

        # Запрещённая inline-инициализация модульной/глобальной переменной.
        if _DECL_INIT_RE.match(line):
            errors.append(
                f"{path}:{lineno}: запрещена инициализация модульной/глобальной "
                f"переменной значением при объявлении: {raw_line.strip()}"
            )

    return errors


def collect_vba_files(target: Path):
    """Возвращает список VBA-файлов (.bas/.cls) для проверки.

    Если target — файл, возвращает [target]. Если target — директория,
    рекурсивно собирает .bas/.cls внутри.
    """
    if target.is_file():
        return [target]
    result = []
    for pattern in ("*.bas", "*.cls"):
        result.extend(target.rglob(pattern))
    return sorted(result)


def main(argv):
    # Путь для проверки: аргумент CLI либо каталог src/ по умолчанию.
    if len(argv) > 1:
        target = Path(argv[1])
        if not target.exists():
            print(f"[ОШИБКА] Указанный путь не существует: {target}")
            return 1
    else:
        target = DEFAULT_SRC_DIR

    if not target.exists():
        print(f"[ОШИБКА] Каталог исходников не найден: {target}")
        return 1

    files = collect_vba_files(target)
    if not files:
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Не найдено VBA-файлов в: {target}")
        return 0

    all_errors = []
    for f in files:
        rel = f.relative_to(PROJECT_DIR) if f.is_relative_to(PROJECT_DIR) else f
        all_errors.extend(check_file(f))

    if all_errors:
        print(f"[FAIL] Найдено ошибок: {len(all_errors)}")
        for err in all_errors:
            print(f"  {err}")
        return 1

    print(f"[OK] Ошибок не найдено (проверено файлов: {len(files)})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))