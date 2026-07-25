#!/usr/bin/env python3
"""
Удаляет встроенную книгу work26.xlsm из UAZ.xlsm.
Удаляет VBA-компоненты: Mod_Import, Mod_Main, Mod_Search, Mod_Settings, Mod_ZN и 33 листа.
"""
import sys
import win32com.client
from win32com.client import gencache
from pathlib import Path

# Путь к UAZ.xlsm
UAZ_PATH = Path(__file__).resolve().parent.parent / "base" / "models" / "UAZ.xlsm"

# Компоненты work26 для удаления
WORK26_MODULES = [
    "Mod_Import", "Mod_Main", "Mod_Search", "Mod_Settings", "Mod_ZN",
]

# Листы work26 для удаления (по именам)
WORK26_SHEETS = [f"Лист{i}" for i in range(1, 34)]  # Лист1..Лист33


def main():
    if not UAZ_PATH.exists():
        print(f"Error: File not found: {UAZ_PATH}")
        sys.exit(1)

    print(f"Opening: {UAZ_PATH}")
    excel = gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(UAZ_PATH))
        vb_project = workbook.VBProject

        # Удаление модулей
        print("\n--- Removing work26 modules ---")
        for mod_name in WORK26_MODULES:
            try:
                comp = vb_project.VBComponents.Item(mod_name)
                vb_project.VBComponents.Remove(comp)
                print(f"  Removed: {mod_name}")
            except Exception:
                print(f"  Not found: {mod_name}")

        # Удаление листов
        print("\n--- Removing work26 sheets ---")
        for sheet_name in WORK26_SHEETS:
            try:
                ws = workbook.Sheets(sheet_name)
                ws.Delete()
                print(f"  Removed sheet: {sheet_name}")
            except Exception:
                pass  # лист может отсутствовать

        workbook.Save()
        print(f"\n=== Work26 successfully removed from {UAZ_PATH.name} ===")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        if workbook:
            workbook.Close()
        excel.Quit()


if __name__ == "__main__":
    main()