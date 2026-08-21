#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общая логика защиты листов шаблонов каталога base/templates/.

Единый источник поведения для build_templates.py и apply_protection_templates.py
(унификация v1.0.4). Согласованное правило:
  - UserInterfaceOnly=False — защита сохраняется в файле (реально пишется
    <sheetProtection> в XML), макросы работают через явные AllowEditRanges;
  - явные AllowEditRanges для зон ввода/данных (не надежда на UserInterfaceOnly);
  - FreezePanes на A4 (закреплены строки 1-3) на листах с заголовками.

Зоны (docs/table.md, раздел 0.4 / решение пользователя v1.0.4):
  work.xlsm, лист main:
      AllowEditRanges/разблокировка: B4, B5:B17, C1, Z1, данные D4:AB2000
      (доп. блокировка столбцов A:B);
  work.xlsm, spisok/libname: A2:J5000;
  work.xlsm, models:          A3:F5000;
  model.xlsm: A4:U2000 на всех листах + C1 на {GroupName} и z4;
  report0.xlsx: A2:Z5000.

Функции:
  apply_protection(ws, sheet_name, ...) — Protect + AllowEditRanges + FreezePanes;
  apply_freeze_only(ws)                  — только FreezePanes A4 (без защиты),
                                           для пустых work0/model0;
  verify_sheet_protection(path)          — проверка XML книги.
"""
import re
import zipfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- Зоны ввода/данных по шаблонам ---
WORK_MAIN_EDIT = ["B4", "B5:B17", "C1", "Z1"]
WORK_MAIN_DATA = "D4:AB2000"     # данные работ/запчастей листа main
WORK_LIST_EDIT = "A2:J5000"      # spisok, libname
WORK_MODELS_EDIT = "A3:F5000"    # models

MODEL_DATA = "A4:U2000"          # все листы модельного шаблона
MODEL_EDIT_EXTRA = ["C1"]        # доп. зона на {GroupName} и z4

REPORT_EDIT = "A2:Z5000"         # report0.xlsx

# Заголовок AllowEditRanges
EDIT_RANGE_TITLE = "Ввод"


def _set_freeze_a4(ws):
    """Закрепляет строки 1-3 (FreezePanes на A4) через COM-окно Excel.

    Использует SplitRow=3 (закреплены 3 верхние строки) + FreezePanes=True.
    Для невидимой книги окно активируется принудительно.
    """
    try:
        ws.Activate()
        win = ws.Application.ActiveWindow
        # Сброс закрепления, затем фиксация 3 верхних строк (A4)
        win.SplitRow = 0
        win.SplitColumn = 0
        win.FreezePanes = False
        win.SplitRow = 3
        win.FreezePanes = True
    except Exception:
        # Закрепление через окно не всегда применимо к невидимой книге —
        # гарантированно устанавливается позднее через openpyxl на XML-уровне
        # (см. ensure_freeze_panes_after_save).
        pass


def _unlock_used(ws):
    """Разблокирует использованный диапазон (сброс предыдущих Locked=True)."""
    try:
        used = ws.UsedRange
        if used is not None:
            used.Locked = False
    except Exception:
        pass


def _resolve_zones(sheet_name, is_main, is_model, is_report):
    """Возвращает (edit_ranges) по типу шаблона и листа."""
    if is_report:
        # report0.xlsx: все листы (report, spisok)
        return [REPORT_EDIT]
    if is_model:
        # model.xlsm: A4:U2000 на всех листах + C1 на {GroupName} и z4
        if sheet_name in ("{GroupName}", "z4"):
            return MODEL_EDIT_EXTRA + [MODEL_DATA]
        return [MODEL_DATA]
    if is_main:
        # work.xlsm, лист main
        return [*WORK_MAIN_EDIT, WORK_MAIN_DATA]
    if sheet_name in ("spisok", "libname"):
        return [WORK_LIST_EDIT]
    if sheet_name == "models":
        return [WORK_MODELS_EDIT]
    # Прочие листы work-шаблона (например, _SETTINGS, если присутствует)
    return ["A4:AB2000"]


def _apply_protect(ws):
    """Включает защиту листа с UserInterfaceOnly=False (сохраняется в файле)."""
    try:
        ws.Protect(Password="", DrawingObjects=False, Contents=True,
                   Scenarios=False, AllowFormattingCells=False,
                   AllowFormattingColumns=False, AllowFormattingRows=False)
    except Exception as e:
        print(f"    [!] Protect на листе '{ws.Name}': {e}")


def _apply_allow_edit_ranges(ws, edit_ranges):
    """Добавляет явные AllowEditRanges для зон ввода/данных."""
    if not edit_ranges:
        return
    try:
        # Удаляем ранее существовавшие диапазоны
        for ar in list(ws.Protection.AllowEditRanges):
            ar.Delete()
        for rng in edit_ranges:
            ws.Protection.AllowEditRanges.Add(
                Title=EDIT_RANGE_TITLE, Range=ws.Range(rng))
    except Exception as e:
        print(f"    [!] AllowEditRanges на листе '{ws.Name}': {e}")


def apply_protection(ws, sheet_name, is_main=False, is_model=False, is_report=False):
    """Применяет Protect + AllowEditRanges + FreezePanes к листу ws.

    is_main   — лист main (доп. блокировка столбцов A:B).
    is_model  — листы модельного шаблона.
    is_report — листы шаблона отчёта.
    """
    # Снять защиту, если была
    try:
        ws.Unprotect()
    except Exception:
        pass

    # Сброс блокировок, затем блокировка строк 1-3
    _unlock_used(ws)
    ws.Rows("1:3").Locked = True

    # Зоны ввода/данных
    edit_ranges = _resolve_zones(sheet_name, is_main, is_model, is_report)

    # Разблокируем зоны ввода/данных
    for rng in edit_ranges:
        try:
            ws.Range(rng).Locked = False
        except Exception:
            pass

    # Для main: дополнительно блокируем столбцы A:B (кроме B4/B5:B17)
    if is_main:
        ws.Range("A:B").Locked = True
        for rng in ["B4", "B5:B17"]:
            try:
                ws.Range(rng).Locked = False
            except Exception:
                pass

    # FreezePanes на A4 (закреплены строки 1-3)
    _set_freeze_a4(ws)

    # Protect (UserInterfaceOnly=False — защита сохраняется в файле)
    _apply_protect(ws)

    # Явные AllowEditRanges
    _apply_allow_edit_ranges(ws, edit_ranges)


def apply_freeze_only(ws):
    """Только FreezePanes A4 без защиты — для пустых шаблонов work0/model0."""
    try:
        ws.Unprotect()
    except Exception:
        pass
    _set_freeze_a4(ws)


# ---------------------------------------------------------------------------
# Проверка итогового XML (без Excel, через zip + стандартную библиотеку)
# ---------------------------------------------------------------------------

_SHEET_XML_RE = re.compile(r'^xl/worksheets/sheet\d+\.xml$')
_PANE_RE = re.compile(r'<pane[^>]*xSplit="0"[^>]*ySplit="(\d+)"[^>]*/>')


def _extract_sheet_xmls(path):
    """Возвращает список (sheet_name, xml_text) из xlsx/xlsm-архива."""
    wb = load_workbook(str(path), keep_vba=True)
    # Порядок листов и их внутренние имена
    names = [ws.title for ws in wb.worksheets]

    results = []
    with zipfile.ZipFile(str(path)) as zf:
        # Сортируем листы по номеру в имени файла
        sheet_files = sorted(
            (n for n in zf.namelist() if _SHEET_XML_RE.match(n)),
            key=lambda n: int(re.search(r'sheet(\d+)\.xml$', n).group(1)),
        )
        for i, sheet_file in enumerate(sheet_files):
            xml_text = zf.read(sheet_file).decode("utf-8")
            name = names[i] if i < len(names) else sheet_file
            results.append((name, xml_text))
    return results


def ensure_freeze_panes_after_save(path):
    """Гарантированно закрепляет строки 1-3 (FreezePanes A4) на XML-уровне.

    Вызывается ПОСЛЕ сохранения книги через Excel COM. Открывает .xlsm/.xlsx
    через openpyxl (keep_vba=True) и задаёт freeze_panes = 'A4' каждому листу.
    Важно: openpyxl переписывает xl/worksheets/*.xml и часть архивных записей,
    но сохраняет vbaProject.bin для .xlsm (keep_vba=True). Используется как
    страховка, когда COM-окно не смогло установить закрепление для невидимой
    книги. При необходимости шрифт/формат сохраняются; допустимо запускать
    повторно (идемпотентно).
    """
    is_xlsm = str(path).lower().endswith(".xlsm")
    wb = load_workbook(str(path), keep_vba=is_xlsm)
    changed = False
    for ws in wb.worksheets:
        if ws.freeze_panes != "A4":
            ws.freeze_panes = "A4"
            changed = True
    if changed:
        wb.save(str(path))
    return changed


def verify_sheet_protection(path):
    """Проверяет XML книги: <sheetProtection>, <allowEditRanges>, <pane> frozen A4.

    Не требует lxml — используется только стандартная библиотека (zipfile, re)
    и openpyxl для получения имён листов. Возвращает список строк-отчётов вида:
        main | protect=да | allowEdit=да (B4, ...) | pane=A4
    """
    reports = []
    for name, xml_text in _extract_sheet_xmls(path):
        prot = "да" if "<sheetProtection" in xml_text else "нет"

        # AllowEditRanges: ищем блок <allowEditRanges>...</allowEditRanges>
        aem = re.search(r'<allowEditRanges[^>]*>(.*?)</allowEditRanges>',
                        xml_text, re.S)
        if aem and re.search(r'<rangeEdit[^>]*>', aem.group(1)):
            sqrefs = re.findall(r'sqref="([^"]+)"', aem.group(1))
            allow_edit = "да (" + ", ".join(sqrefs) + ")"
        else:
            allow_edit = "нет"

        # FreezePanes: pane с ySplit (строки закрепления); frozen A4 => ySplit=3
        pane = "нет"
        pm = re.search(_PANE_RE, xml_text)
        if pm:
            ysplit = int(pm.group(1))
            pane = "A4" if ysplit == 3 else f"frozen ySplit={ysplit}"
        else:
            if "<pane" in xml_text:
                pane = "присутствует (не frozen A4)"

        reports.append(f"{name} | protect={prot} | allowEdit={allow_edit} | pane={pane}")
    return reports


def list_allowed_edits(path):
    """Краткий список зон AllowEditRanges по листам (для резюме)."""
    out = []
    for name, xml_text in _extract_sheet_xmls(path):
        aem = re.search(r'<allowEditRanges[^>]*>(.*?)</allowEditRanges>',
                        xml_text, re.S)
        sqrefs = re.findall(r'sqref="([^"]+)"', aem.group(1)) if aem else []
        out.append(f"{name}: {sqrefs}")
    return out