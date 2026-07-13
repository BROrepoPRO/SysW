#!/usr/bin/env python3
"""
Import all VBA modules (bas/cls) into work.xlsm via Excel COM.
Strips ALL Attribute lines before import to avoid VBA syntax errors.
Converts UTF-8 source to Windows-1251 for VBA compatibility.
"""
import os
import sys
import shutil
import re
import win32com.client
from win32com.client import gencache
from pathlib import Path

EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"
MODULES_PATH = Path(r"L:\PROject\SysW")
TEMP_DIR = Path(r"L:\PROject\SysW\_temp_import")

FILES = [
    "Mod_Utils.bas",
    "Mod_OrderHeader.bas",
    "Mod_Import.bas",
    "Mod_ButtonDispatcher.bas",
    "Mod_MinimalTestRunner.bas",
    "Sheet1_main.cls",
]


def strip_all_attribute_lines(text, is_cls=False):
    """Remove Attribute lines from VBA source.

    For .bas files: remove ALL Attribute lines (VB_Name only).
    For .cls files: keep critical Attribute lines (VB_Name, VB_PredeclaredId, VB_Exposed),
                    remove optional ones (VB_GlobalNameSpace, VB_Creatable, etc.)
    """
    if is_cls:
        # For .cls files, keep only critical attributes
        # Remove all Attribute lines EXCEPT VB_Name, VB_PredeclaredId, VB_Exposed
        lines = text.split('\n')
        filtered = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith('Attribute '):
                # Keep only critical attributes
                if ('VB_Name' in stripped or
                    'VB_PredeclaredId' in stripped or
                    'VB_Exposed' in stripped):
                    filtered.append(line)
                # else: skip this Attribute line
            else:
                filtered.append(line)
        return '\n'.join(filtered)
    else:
        # For .bas files, remove ALL Attribute lines (as before)
        return re.sub(r'^Attribute\s+\w+\s*=\s*"[^"]*"\s*$', '', text, flags=re.MULTILINE)


def main():
    # Ensure temp dir
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(parents=True)

    print("Creating Excel COM object...")
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    success = False

    try:
        print(f"Opening workbook: {EXCEL_PATH}")
        workbook = excel.Workbooks.Open(EXCEL_PATH)

        print("Accessing VBA project...")
        vb_project = workbook.VBProject

        comp_count = vb_project.VBComponents.Count
        print(f"VBA project accessible. Components count: {comp_count}")

        # First pass: remove all our components
        print("")
        print("--- Removing existing components ---")
        for file_name in FILES:
            component_name = Path(file_name).stem
            try:
                existing = vb_project.VBComponents.Item(component_name)
                if existing:
                    vb_project.VBComponents.Remove(existing)
                    print(f"  Removed: {component_name}")
            except Exception:
                print(f"  Not found: {component_name}")

        # Second pass: import all components
        print("")
        print("--- Importing components ---")
        for file_name in FILES:
            file_path = MODULES_PATH / file_name
            print("")
            print(f"  Processing: {file_name}")

            if not file_path.exists():
                print(f"    [!] File not found: {file_path}")
                continue

            component_name = file_path.stem

            # Read UTF-8, strip Attribute lines, convert to Windows-1251
            temp_file = TEMP_DIR / file_name
            with open(file_path, "r", encoding="utf-8") as f:
                text = f.read()

            # Strip Attribute lines (different handling for .bas vs .cls)
            is_cls = file_name.lower().endswith('.cls')
            text = strip_all_attribute_lines(text, is_cls=is_cls)
            print(f"    Stripped Attribute lines (is_cls={is_cls})")

            # Remove leading blank lines
            text = text.lstrip('\n\r')

            with open(temp_file, "w", encoding="cp1251") as f:
                f.write(text)

            print(f"    Converted encoding: UTF-8 -> Windows-1251")

            # Import
            print(f"    Calling Import on: {temp_file}")
            imported = vb_project.VBComponents.Import(str(temp_file))
            if imported is None:
                print(f"    [!] Import returned None!")
            else:
                print(f"    [+] Successfully imported: {file_name}")

        print("")
        print("Saving workbook...")
        workbook.Save()
        print("Workbook saved.")
        success = True

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        print("")
        print("Possible causes:")
        print("  1. 'Trust access to the VBA project object model' is not enabled")
        print("     (File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings)")
        print("  2. Make sure work.xlsm is not open in Excel")
        print("  3. Install pywin32: pip install pywin32")
    finally:
        # Clean up temp
        if TEMP_DIR.exists():
            shutil.rmtree(TEMP_DIR, ignore_errors=True)

        if workbook is not None:
            print("Closing workbook...")
            workbook.Close()

        print("Closing Excel...")
        excel.Quit()

    if success:
        print("")
        print("=" * 40)
        print("ALL MODULES IMPORTED SUCCESSFULLY!")
        print("=" * 40)
    else:
        print("")
        print("=" * 40)
        print("IMPORT FAILED")
        print("=" * 40)


if __name__ == "__main__":
    main()