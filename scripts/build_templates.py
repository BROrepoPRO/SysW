#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт создания файлов-шаблонов каталога base/templates/ версии 1.0.4.

Задачи (B и C плана integration_1.0.4.md):
  B. Создать 5 шаблонов:
       work.xlsm  - шаблон work с последним обновлением кодовой базы
                      (копия корневого work.xlsm после impVBA.py)
       work0.xlsm - шаблон work пустой (без VBA-кода)
       model.xlsm - общий модельный шаблон с кодом
       model0.xlsm- модельный шаблон пустой (без VBA-кода)
       report0.xlsx - шаблон отчёта пустой (без данных, без кода)
  C. Защита листов в шаблонах: Protect + AllowEditRanges + FreezePanes A4.

Единая логика защиты вынесена в scripts/template_protection.py (v1.0.4):
  - UserInterfaceOnly=False — защита сохраняется в файле (<sheetProtection> в XML);
  - явные AllowEditRanges для зон ввода/данных;
  - FreezePanes на A4 (закреплены строки 1-3).
См. template_protection.apply_protection.

Применяемый состав листов (docs/table.md, раздел 0.11):
  - work/work0: main, spisok, models, libname (БЕЗ _SETTINGS, макетов, temp*);
  - model/model0: z4, {GroupName}, {GroupName}w, {GroupName}z4 (БЕЗ макетов {NN}M);
  - report0: report, spisok.

Использует Excel COM (win32com). Требуется включённый AccessVBOM.
"""
import re
import shutil
import time
import gc
import win32com.client
import win32process
from pathlib import Path
from win32com.client import gencache

from template_protection import (
    apply_protection,
    apply_freeze_only,
    ensure_freeze_panes_after_save,
    apply_protection_xml,
    apply_freeze_panes_xml,
    build_zone_map,
)

PROJECT_DIR = Path(__file__).resolve().parent.parent
TEMPLATES = PROJECT_DIR / "base" / "templates"
WORKBOOK = PROJECT_DIR / "work.xlsm"
REPORT = PROJECT_DIR / "report.xlsx"
MODELS = PROJECT_DIR / "base" / "models"

# Группы модельных файлов (порядок объединения)
GROUPS = ["4x4", "2170", "2180", "2190", "GAZ", "UAZ"]

# Имя этапа для лог-файла PID (используется build_all.py для безопасного taskkill).
STAGE_NAME = "templates"

# Константы ретраев открытия книги.
OPEN_RETRIES = 5   # Максимальное число попыток открытия.
OPEN_PAUSE = 3     # Пауза (сек) между попытками.


def sheet_excluded(name):
    """Проверяет, нужно ли исключить лист из шаблона.

    Исключаются: _SETTINGS (резерв, по решению не включаем), temp*, *old,
    дубликаты (2), макеты {NN}M (по решению не включаем в модельные шаблоны),
    временные/архивные ЗНW/СчетW/3M.
    """
    low = name.lower()
    if low == "_settings":
        return True  # _SETTINGS — резерв, не включаем
    if low.startswith("temp"):
        return True
    if low.endswith("old"):
        return True
    if low.endswith("(2)"):
        return True
    if low in ("знw", "счетw", "3m"):
        return True
    # Макеты {NN}M (например 1M, 2M, 001M) — не включаем в модельные шаблоны
    if re.fullmatch(r"\d+m", low):
        return True
    return False


def get_excel():
    """Создаёт и возвращает объект Excel.Application."""
    try:
        excel = gencache.EnsureDispatch("Excel.Application")
    except AttributeError:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    # Автоматизация безопасности: скрипт НЕ выполняет макросы (только
    # открывает/копирует/сохраняет книги), поэтому макросы можно принудительно
    # отключить (msoAutomationSecurityForceDisable = 3) — как в impVBA.py.
    # Обёрнуто в try/except: при сбое установки это не критично.
    try:
        excel.AutomationSecurity = 3
    except Exception as exc:
        print(f"    [!] Не удалось установить AutomationSecurity=3: {exc}")
    return excel


def open_workbook_with_retry(excel, path, retries=OPEN_RETRIES, pause_sec=OPEN_PAUSE):
    """Открывает книгу с ретраями при None/COM-ошибке -2147352567.

    Workbooks.Open может вернуть None либо выбросить прерывистую COM-ошибку
    при модальных окнах/плохом состоянии экземпляра. Повторяет открытие до
    retries раз с паузой pause_sec и сборкой мусора между попытками. Если все
    попытки исчерпаны — возвращает None (вызывающий код обработает это как
    ошибку этапа).
    """
    for attempt in range(1, retries + 1):
        wb = None
        try:
            # Явные параметры: ReadOnly=False (на запись), UpdateLinks=0
            # (не обновлять связи), ConfirmConversion=False (не спрашивать
            # про конвертацию формата) — защита от зависаний на модальных окнах.
            wb = excel.Workbooks.Open(
                str(path),
                ReadOnly=False,
                UpdateLinks=0,
                ConfirmConversion=False,
            )
        except Exception as exc:
            msg = str(exc)
            print(f"    Попытка {attempt}/{retries}: ошибка открытия ({msg}); "
                  f"пауза {pause_sec}с")
            wb = None
            time.sleep(pause_sec)
            gc.collect()
            continue

        if wb is None:
            print(f"    Попытка {attempt}/{retries}: Workbooks.Open вернул None; "
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
        pid_file = PROJECT_DIR / "logs" / f"excel_pid_{STAGE_NAME}.txt"
        pid_file.parent.mkdir(parents=True, exist_ok=True)
        pid_file.write_text(str(pid), encoding="utf-8")
        print(f"    Записан PID Excel {pid} -> {pid_file}")
    except Exception as exc:
        print(f"    [!] Не удалось записать PID Excel: {exc}")


def copy_sheets(src_wb, dst_wb, include_names):
    """Копирует листы с именами include_names из src_wb в dst_wb.

    Использует первый лист dst_wb (Sheet1) как приёмник для первого
    включённого листа, чтобы в книге всегда оставался хотя бы один лист
    (Excel запрещает удалять последний видимый лист).
    """
    first = True
    for name in include_names:
        try:
            ws = src_wb.Sheets(name)
        except Exception:
            print(f"  [!] Лист '{name}' не найден в источнике, пропуск")
            continue
        if first:
            target = dst_wb.Sheets(1)
            try:
                ws.Copy(Before=target)
                target.Delete()
                new_sheet = dst_wb.Sheets(1)
                new_sheet.Name = name
            except Exception:
                target.Name = name
            first = False
        else:
            ws.Copy(None, dst_wb.Sheets(dst_wb.Sheets.Count))
            new_sheet = dst_wb.Sheets(dst_wb.Sheets.Count)
            new_sheet.Name = name


def _reduce_to_include(wb, include):
    """Удаляет из книги все листы, кроме include, безопасно (со страховкой)."""
    guard = wb.Sheets.Add()
    guard.Name = "__guard__"
    to_delete = [s.Name for s in wb.Sheets if s.Name not in include and s.Name != "__guard__"]
    for name in to_delete:
        try:
            wb.Sheets(name).Delete()
        except Exception as e:
            print(f"  [!] Не удалён лист '{name}': {e}")
    guard.Delete()


def _remove_vba(wb):
    """Удаляет все VBA-компоненты (модули, классы, документы листов)."""
    try:
        vb = wb.VBProject
        for i in range(vb.VBComponents.Count, 0, -1):
            comp = vb.VBComponents(i)
            if comp.Type in (1, 2, 100):
                vb.VBComponents.Remove(comp)
    except Exception as e:
        print(f"    [!] Удаление VBA-компонентов: {e}")


def build_work_templates(excel):
    """Создаёт work.xlsm (копия корневого work.xlsm с кодом) и work0.xlsm (без кода)."""
    include = ["main", "spisok", "models", "libname"]

    # --- work.xlsm: копия корневого work.xlsm (уже содержит VBA после impVBA.py).
    dst_path = TEMPLATES / "work.xlsm"
    shutil.copyfile(str(WORKBOOK), str(dst_path))
    wb = open_workbook_with_retry(excel, str(dst_path))
    if wb is None:
        raise RuntimeError(f"Не удалось открыть книгу: {dst_path}")
    _reduce_to_include(wb, include)
    wb.Save()
    wb.Close()

    # --- work0.xlsm: копия без VBA-кода ---
    wb0 = open_workbook_with_retry(excel, str(dst_path))
    if wb0 is None:
        raise RuntimeError(f"Не удалось открыть книгу: {dst_path}")
    _remove_vba(wb0)
    wb0.SaveAs(str(TEMPLATES / "work0.xlsm"), FileFormat=52)
    wb0.Close()
    print("  Созданы work.xlsm и work0.xlsm")


def build_model_templates(excel):
    """Создаёт model.xlsm и model0.xlsm (общая структура листов)."""
    # Плейсхолдер группы. За основу структуры берём GAZ как эталон.
    group = "GAZ"
    src = open_workbook_with_retry(excel, str(MODELS / f"{group}.xlsm"))
    if src is None:
        raise RuntimeError(f"Не удалось открыть модель {group}.xlsm")
    dst = excel.Workbooks.Add()

    # Копируем 4 основных листа группы GAZ (первый использует Sheet1)
    first = True
    for name in [group, f"{group}w", "z4", f"{group}z4"]:
        try:
            ws = src.Sheets(name)
            if first:
                target = dst.Sheets(1)
                try:
                    ws.Copy(Before=target)
                    target.Delete()
                    dst.Sheets(1).Name = name
                except Exception:
                    target.Name = name
                first = False
            else:
                ws.Copy(None, dst.Sheets(dst.Sheets.Count))
                dst.Sheets(dst.Sheets.Count).Name = name
        except Exception as e:
            print(f"  [!] Копирование листа {name}: {e}")

    # Переименование в плейсхолдер {GroupName}
    rename_map = {
        group: "{GroupName}",
        f"{group}w": "{GroupName}w",
        f"{group}z4": "{GroupName}z4",
    }
    for old, new in rename_map.items():
        try:
            dst.Sheets(old).Name = new
        except Exception:
            pass

    # Удалить все лишние листы (макеты, temp*, архивы) со страховкой
    keep = ["{GroupName}", "{GroupName}w", "z4", "{GroupName}z4"]
    _reduce_to_include(dst, keep)

    dst.SaveAs(str(TEMPLATES / "model.xlsm"), FileFormat=52)
    dst.Close()

    # model0.xlsm: без VBA-кода
    wb = open_workbook_with_retry(excel, str(TEMPLATES / "model.xlsm"))
    if wb is None:
        raise RuntimeError("Не удалось открыть model.xlsm")
    _remove_vba(wb)
    wb.SaveAs(str(TEMPLATES / "model0.xlsm"), FileFormat=52)
    wb.Close()

    src.Close()
    print("  Созданы model.xlsm и model0.xlsm")


def build_report_template(excel):
    """Создаёт report0.xlsx (пустой, без данных, без кода)."""
    include = ["report", "spisok"]
    src = open_workbook_with_retry(excel, str(REPORT))
    if src is None:
        raise RuntimeError(f"Не удалось открыть отчёт: {REPORT}")
    dst = excel.Workbooks.Add()
    copy_sheets(src, dst, include)
    _reduce_to_include(dst, include)
    dst.SaveAs(str(TEMPLATES / "report0.xlsx"), FileFormat=51)  # xlsx
    dst.Close()
    src.Close()
    print("  Создан report0.xlsx")


def apply_all_protection(excel):
    """Применяет защиту + FreezePanes к листам шаблонов (единая логика).

    Вариант D (v1.0.7): COM задаёт Locked-флаги ячеек (формат сохраняет Excel),
    а sheetProtection/allowEditRanges/pane доставляются на XML-уровне
    (apply_protection_xml / apply_freeze_panes_xml) без пересохранения через
    openpyxl — максимальное сохранение исходного формата.
    """
    # work.xlsm
    wb = open_workbook_with_retry(excel, str(TEMPLATES / "work.xlsm"))
    if wb is None:
        raise RuntimeError("Не удалось открыть шаблон work.xlsm")
    for ws in wb.Sheets:
        apply_protection(ws, ws.Name, is_main=(ws.Name == "main"))
    wb.Save()
    wb.Close()
    apply_protection_xml(TEMPLATES / "work.xlsm", build_zone_map("work"))

    # work0.xlsm — пустой: только FreezePanes A4, без Protect/AllowEditRanges
    wb0 = open_workbook_with_retry(excel, str(TEMPLATES / "work0.xlsm"))
    if wb0 is None:
        raise RuntimeError("Не удалось открыть шаблон work0.xlsm")
    for ws in wb0.Sheets:
        apply_freeze_only(ws)
    wb0.Save()
    wb0.Close()
    apply_freeze_panes_xml(TEMPLATES / "work0.xlsm")

    # model.xlsm
    wb = open_workbook_with_retry(excel, str(TEMPLATES / "model.xlsm"))
    if wb is None:
        raise RuntimeError("Не удалось открыть шаблон model.xlsm")
    for ws in wb.Sheets:
        apply_protection(ws, ws.Name, is_model=True)
    wb.Save()
    wb.Close()
    apply_protection_xml(TEMPLATES / "model.xlsm", build_zone_map("model"))

    # model0.xlsm — пустой: только FreezePanes A4, без Protect/AllowEditRanges
    wb0 = open_workbook_with_retry(excel, str(TEMPLATES / "model0.xlsm"))
    if wb0 is None:
        raise RuntimeError("Не удалось открыть шаблон model0.xlsm")
    for ws in wb0.Sheets:
        apply_freeze_only(ws)
    wb0.Save()
    wb0.Close()
    apply_freeze_panes_xml(TEMPLATES / "model0.xlsm")

    # report0.xlsx
    wb = open_workbook_with_retry(excel, str(TEMPLATES / "report0.xlsx"))
    if wb is None:
        raise RuntimeError("Не удалось открыть шаблон report0.xlsx")
    for ws in wb.Sheets:
        apply_protection(ws, ws.Name, is_report=True)
    wb.Save()
    wb.Close()
    apply_protection_xml(TEMPLATES / "report0.xlsx", build_zone_map("report"))


def main():
    TEMPLATES.mkdir(parents=True, exist_ok=True)
    excel = get_excel()
    # Записываем PID созданного экземпляра Excel для безопасного завершения
    # зависшего процесса конвейером build_all.py (только наш процесс).
    write_excel_pid(excel)
    try:
        print("Создание шаблонов (задача B)...")
        build_work_templates(excel)
        build_model_templates(excel)
        build_report_template(excel)
        print("")
        print("Применение защиты листов (задача C)...")
        apply_all_protection(excel)
        print("Готово.")
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()