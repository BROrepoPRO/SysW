#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт применения защиты листов (задача C) к шаблонам base/templates/.

Единая логика защиты вынесена в scripts/template_protection.py (v1.0.4):
  - UserInterfaceOnly=False — защита сохраняется в файле (<sheetProtection> в XML);
  - явные AllowEditRanges для зон ввода/данных;
  - FreezePanes на A4 (закреплены строки 1-3).
См. template_protection.apply_protection.

Требования (docs/table.md раздел 0.4, срочное требование v1.0.4):
  - Protect + AllowEditRanges для зон ввода/данных
  - FreezePanes на A4 (закреплены строки 1-3) на листах с заголовками
  - для main: столбцы A:B только защита, закрепление только на A4
  - 3 верхние строки сохраняют формат шаблона (блокируются)

Пустые шаблоны (work0.xlsm, model0.xlsm) получают ТОЛЬКО FreezePanes A4
(apply_freeze_only) — без Protect/AllowEditRanges, т.к. защита применяется
при наполнении (docs/table.md, разделы 0.11.2/0.11.4).

Использует Excel COM (win32com). Нужен включённый AccessVBOM для *.xlsm.
"""
from pathlib import Path
import win32com.client
from win32com.client import gencache

from template_protection import (
    apply_protection,
    apply_freeze_only,
    ensure_freeze_panes_after_save,
)

PROJECT_DIR = Path(__file__).resolve().parent.parent
TEMPLATES = PROJECT_DIR / "base" / "templates"


def get_excel():
    try:
        excel = gencache.EnsureDispatch("Excel.Application")
    except AttributeError:
        excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def main():
    excel = get_excel()
    try:
        # --- work.xlsm: полная защита (Protect + AllowEditRanges + FreezePanes) ---
        print("Защита work.xlsm ...")
        wb = excel.Workbooks.Open(str(TEMPLATES / "work.xlsm"))
        for ws in wb.Sheets:
            apply_protection(ws, ws.Name, is_main=(ws.Name == "main"))
        wb.Save()
        wb.Close()
        ensure_freeze_panes_after_save(TEMPLATES / "work.xlsm")

        # --- work0.xlsm: только FreezePanes A4 (защита при наполнении) ---
        print("Защита work0.xlsm (только FreezePanes A4) ...")
        wb0 = excel.Workbooks.Open(str(TEMPLATES / "work0.xlsm"))
        for ws in wb0.Sheets:
            apply_freeze_only(ws)
        wb0.Save()
        wb0.Close()
        ensure_freeze_panes_after_save(TEMPLATES / "work0.xlsm")

        # --- model.xlsm: полная защита ---
        print("Защита model.xlsm ...")
        wb = excel.Workbooks.Open(str(TEMPLATES / "model.xlsm"))
        for ws in wb.Sheets:
            apply_protection(ws, ws.Name, is_model=True)
        wb.Save()
        wb.Close()
        ensure_freeze_panes_after_save(TEMPLATES / "model.xlsm")

        # --- model0.xlsm: только FreezePanes A4 (защита при наполнении) ---
        print("Защита model0.xlsm (только FreezePanes A4) ...")
        wb0 = excel.Workbooks.Open(str(TEMPLATES / "model0.xlsm"))
        for ws in wb0.Sheets:
            apply_freeze_only(ws)
        wb0.Save()
        wb0.Close()
        ensure_freeze_panes_after_save(TEMPLATES / "model0.xlsm")

        # --- report0.xlsx: полная защита ---
        print("Защита report0.xlsx ...")
        wb = excel.Workbooks.Open(str(TEMPLATES / "report0.xlsx"))
        for ws in wb.Sheets:
            apply_protection(ws, ws.Name, is_report=True)
        wb.Save()
        wb.Close()
        ensure_freeze_panes_after_save(TEMPLATES / "report0.xlsx")

        print("Готово.")
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()