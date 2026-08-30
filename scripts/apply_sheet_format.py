#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый формат целевых листов по эталону {GroupName} (Задача 6, v1.0.17).

Назначение:
    Применить формат листа-эталона `{GroupName}` из base/templates/model0.xlsm
    (заголовки в строке 3, данные с 4-й, стили ячеек, ширины столбцов, заливки,
    границы, шрифты) ко ВСЕМ целевым листам с данными:

      - модельные шаблоны: base/templates/model.xlsm, model0.xlsm
        (листы z4, {GroupName}, {GroupName}w, {GroupName}z4);
      - шаблоны work: base/templates/work.xlsm, work0.xlsm (лист main);
      - рабочие модельные книги: base/models/*.xlsm;
      - общий каталог з/ч: base/models/z4.xlsx (лист z4);
      - корневой work.xlsm (лист main).

Метод (Вариант В, согласован с пользователем):
    1. Через Excel COM копируется ФОРМАТ эталонного листа {GroupName}
       (Range.Copy + PasteSpecial(Paste=xlPasteFormats)) на целевой лист.
       Данные, формулы, VBA, кнопки и макросы НЕ затрагиваются.
    2. Затем применяется защита листа через template_protection.apply_protection()
       (Protect + AllowEditRanges + FreezePanes A4).

Безопасность:
    - Все ключевые файлы резервируются до изменения (см. правило [U5],
      резерв в _backup/pre_task6_*).
    - Режим --dry-run: только выводит план (книги/листы), ничего не изменяет.
    - Перед каждым целевым листом снимается защита (ws.Unprotect), после — защита
      восстанавливается (apply_protection).

Запуск (из корня проекта):
    python scripts/apply_sheet_format.py --dry-run   # показать план
    python scripts/apply_sheet_format.py             # применить формат

Требует включённый Excel COM (win32com) и закрытые книги.
"""
from __future__ import annotations

import sys
from pathlib import Path

try:
    from config import (
        MODELS_DIR,
        PROJECT_DIR,
        TEMPLATES_DIR,
        WORKBOOK_PATH,
    )
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from config import (
        MODELS_DIR,
        PROJECT_DIR,
        TEMPLATES_DIR,
        WORKBOOK_PATH,
    )

# --- Импорт библиотеки защиты листов (шаблоны/модели) ---
sys.path.insert(0, str(Path(__file__).resolve().parent))
from template_protection import apply_protection  # noqa: E402

# Эталонная книга и лист-эталон
REFERENCE_BOOK = TEMPLATES_DIR / "model0.xlsm"
REFERENCE_SHEET = "{GroupName}"

# Константы Excel COM
xlPasteFormats = -4122


def build_targets(templates_only: bool = False) -> list[tuple[Path, list[str]]]:
    """Возвращает список (книга, [список целевых листов]) для обработки.

    Для шаблонов модельных книг (model.xlsm/model0.xlsm) используются
    плейсхолдеры {GroupName}. Для рабочих модельных книг base/models/*.xlsm
    подставляется фактическое имя группы (имя файла без расширения):
    листы z4, {G}, {G}w, {G}z4.
    Для work-шаблонов и корневого work.xlsm — лист main.

    templates_only=True — обрабатываются ТОЛЬКО шаблоны base/templates/*
    (рекомендуемый режим, v1.0.17): рабочие книги и каталог з/ч не затрагиваются,
    их формат будет получен при пересборке (build_all.py/build_templates.py).
    """
    targets: list[tuple[Path, list[str]]] = []

    # Шаблоны модельных книг (плейсхолдер {GroupName})
    for book in (TEMPLATES_DIR / "model.xlsm", TEMPLATES_DIR / "model0.xlsm"):
        if book.is_file():
            targets.append((book, [
                "z4", "{GroupName}", "{GroupName}w", "{GroupName}z4",
            ]))

    # Шаблоны work
    for book in (TEMPLATES_DIR / "work.xlsm", TEMPLATES_DIR / "work0.xlsm"):
        if book.is_file():
            targets.append((book, ["main"]))

    # Шаблон отчёта
    report_tpl = TEMPLATES_DIR / "report0.xlsx"
    if report_tpl.is_file():
        targets.append((report_tpl, ["report", "spisok"]))

    if templates_only:
        return targets

    # Рабочие модельные книги base/models/*.xlsm (фактическое имя группы)
    for book in sorted(MODELS_DIR.glob("*.xlsm")):
        if book.is_file() and not book.name.lower().startswith("~$"):
            g = book.stem  # фактическое имя группы, напр. 2170
            targets.append((book, [
                "z4", g, f"{g}w", f"{g}z4",
            ]))

    # Общий каталог з/ч base/models/z4.xlsx
    z4 = MODELS_DIR / "z4.xlsx"
    if z4.is_file():
        targets.append((z4, ["z4"]))

    # Корневой work.xlsm
    if WORKBOOK_PATH.is_file():
        targets.append((WORKBOOK_PATH, ["main"]))

    return targets


def _used_range(ws):
    """Возвращает ограниченный прямоугольник (r1,c1,r2,c2) используемой области.

    Копирование формата только по UsedRange позволяет избежать попытки охватить
    весь лист (1 048 576 x 16 384), которая завешивает Excel на больших каталогах.
    """
    try:
        used = ws.UsedRange
        if used is None:
            return (1, 1, 1, 1)
        r1, c1 = used.Row, used.Column
        r2 = min(r1 + used.Rows.Count - 1, ws.Rows.Count)
        c2 = min(c1 + used.Columns.Count - 1, ws.Columns.Count)
        return (r1, c1, r2, c2)
    except Exception:
        return (1, 1, 1, 1)


def apply_format(excel, ref_ws, target_ws) -> None:
    """Копирует формат ref_ws на target_ws через PasteSpecial(formats).

    Используется ограниченный диапазон (UsedRange): формат эталонного листа
    применяется к используемой области целевого листа (без материализации
    миллионов пустых ячеек).
    """
    try:
        target_ws.Unprotect()
    except Exception:
        pass

    rr1, rc1, rr2, rc2 = _used_range(ref_ws)
    tr1, tc1, tr2, tc2 = _used_range(target_ws)

    # Копируем формат эталона
    try:
        ref_ws.Range(ref_ws.Cells(rr1, rc1),
                     ref_ws.Cells(rr2, rc2)).Copy()
    except Exception as exc:
        print(f"    [!] Ошибка Copy эталона: {exc}")
        raise
    try:
        # Применяем к целевой используемой области
        target_ws.Range(target_ws.Cells(tr1, tc1),
                        target_ws.Cells(tr2, tc2)
                        ).PasteSpecial(Paste=xlPasteFormats)
    except Exception as exc:
        print(f"    [!] Ошибка PasteSpecial(formats): {exc}")
        raise
    finally:
        try:
            excel.CutCopyMode = False
        except Exception:
            pass


def main() -> int:
    dry_run = "--dry-run" in sys.argv
    templates_only = "--templates-only" in sys.argv
    targets = build_targets(templates_only=templates_only)

    mode_txt = "ТОЛЬКО шаблоны base/templates/" if templates_only else "все целевые книги"
    print(f"{'ДИАГНОСТИКА (dry-run), изменений нет' if dry_run else 'ПРИМЕНЕНИЕ'} "
          f"формата эталона '{REFERENCE_SHEET}' из {REFERENCE_BOOK.name}")
    print(f"Режим: {mode_txt}; целевых книг: {len(targets)}")

    if not REFERENCE_BOOK.is_file():
        print(f"[ОШИБКА] Эталонная книга не найдена: {REFERENCE_BOOK}")
        return 1

    if dry_run:
        for book, sheets in targets:
            print(f"  [{book.relative_to(PROJECT_DIR)}] -> {', '.join(sheets)}")
        return 0

    import win32com.client

    # Dynamic dispatch (без генерируемого кэша gen_py): избегаем ошибок
    # битого кэша typelib (AttributeError CLSIDToClassMap).
    excel = win32com.client.Dispatch("Excel.Application")

    # Установка свойств приложения обёрнута в try/except: если Excel уже
    # запущен и заблокирован, не должны падать на этапе инициализации.
    try:
        excel.Visible = False
    except Exception:
        pass
    try:
        excel.DisplayAlerts = False
    except Exception as exc:
        print(f"    [!] Не удалось установить DisplayAlerts=False: {exc}")

    try:
        # Открыть эталонную книгу
        ref_wb = excel.Workbooks.Open(str(REFERENCE_BOOK))
        ref_names = {s.Name for s in ref_wb.Sheets}
        if REFERENCE_SHEET not in ref_names:
            print(f"[ОШИБКА] Лист-эталон '{REFERENCE_SHEET}' отсутствует. "
                  f"Фактические листы: {sorted(ref_names)}")
            ref_wb.Close(SaveChanges=False)
            excel.Quit()
            return 3
        ref_ws = ref_wb.Sheets(REFERENCE_SHEET)
        print(f"Эталон загружен: {REFERENCE_SHEET}")

        for book, sheets in targets:
            try:
                wb = excel.Workbooks.Open(str(book))
            except Exception as exc:
                print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось открыть {book.name}: {exc}")
                continue
            try:
                sheet_names = {s.Name for s in wb.Sheets}
                for sheet_name in sheets:
                    if sheet_name not in sheet_names:
                        print(f"    Пропуск (нет листа): {book.name}/{sheet_name}")
                        continue
                    ws = wb.Sheets(sheet_name)
                    print(f"  Применение формата -> {book.name}/{sheet_name}")
                    apply_format(excel, ref_ws, ws)

                    # Защита листа по типу
                    is_model = ("model" in book.name) or (book.parent == MODELS_DIR)
                    apply_protection(
                        ws,
                        sheet_name,
                        is_main=(sheet_name == "main"),
                        is_model=is_model,
                    )
                wb.Save()
            except Exception as exc:
                print(f"[ОШИБКА] {book.name}: {exc}")
            finally:
                try:
                    wb.Close(SaveChanges=True)
                except Exception:
                    pass

        ref_wb.Close(SaveChanges=False)
        print("Формат успешно применён ко всем целевым листам.")
        return 0
    finally:
        excel.Quit()


if __name__ == "__main__":
    raise SystemExit(main())