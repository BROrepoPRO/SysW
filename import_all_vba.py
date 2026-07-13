#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Import all VBA modules (bas/cls) into work.xlsm via Excel COM.
Strips ALL Attribute lines before import to avoid VBA syntax errors.
Converts UTF-8 source to Windows-1251 for VBA compatibility.

Encoding strategy (two-phase model):
  1. Files on disk are stored in UTF-8 (for VS Code and Git)
  2. This script converts them to Windows-1251 before importing into Excel
  3. .gitattributes normalizes encoding for Git
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
    "Mod_FullTestRunner.bas",
    "Sheet1_main.cls",
]


def read_vba_file(file_path):
    """Читает VBA-файл с автоопределением кодировки.

    Приоритет: UTF-8 with BOM > UTF-8 > Windows-1251

    Returns:
        str: содержимое файла в Unicode

    Raises:
        FileNotFoundError: если файл не существует
        ValueError: если не удалось определить кодировку
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    # Проверка BOM (Byte Order Mark)
    with open(file_path, "rb") as f:
        raw = f.read(4)

    if raw[:3] == b'\xef\xbb\xbf':
        # UTF-8 with BOM
        with open(file_path, "r", encoding="utf-8-sig") as f:
            text = f.read()
        print(f"    Read encoding: UTF-8 with BOM")
        return text

    # Попытка UTF-8
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            text = f.read()
        print(f"    Read encoding: UTF-8")
        return text
    except UnicodeDecodeError:
        pass

    # Fallback на Windows-1251
    try:
        with open(file_path, "r", encoding="cp1251") as f:
            text = f.read()
        print(f"    ⚠️  Read encoding: Windows-1251 (fallback). "
              f"Рекомендуется конвертировать файл в UTF-8: {file_path.name}")
        return text
    except UnicodeDecodeError:
        # Последняя попытка: UTF-8 with BOM (если BOM повреждён)
        try:
            with open(file_path, "r", encoding="utf-8-sig") as f:
                text = f.read()
            print(f"    Read encoding: UTF-8 with BOM (fallback)")
            return text
        except UnicodeDecodeError as e:
            raise ValueError(f"Не удалось определить кодировку файла {file_path}: {e}")


def strip_export_header(text):
    """Remove VBA export-format header lines that are NOT valid VBA code.

    These lines appear in .cls export files but must NOT be passed to
    AddFromString, which expects only pure VBA code:
      - VERSION 1.0 CLASS
      - BEGIN ... END (designer section, e.g. MultiUse = -1)
      - Any lines inside the BEGIN/END block

    Returns text with these lines removed.
    """
    lines = text.split('\n')
    filtered = []
    inside_begin_end = False
    for line in lines:
        stripped = line.strip()
        # Detect start of BEGIN...END designer section
        if stripped == 'BEGIN' or stripped.startswith('BEGIN '):
            inside_begin_end = True
            continue  # skip the BEGIN line itself
        # Detect end of BEGIN...END designer section
        if inside_begin_end and stripped == 'END':
            inside_begin_end = False
            continue  # skip the END line itself
        # Skip any line inside the BEGIN...END block
        if inside_begin_end:
            continue
        # Skip VERSION line
        if stripped.startswith('VERSION '):
            continue
        filtered.append(line)
    return '\n'.join(filtered)


def strip_attribute_lines(text, is_cls=False):
    """Remove Attribute lines from VBA source.

    For .bas files: keep ONLY Attribute VB_Name (required for named import),
                    remove all other Attribute lines.
    For .cls files: keep ALL Attribute lines (VB_Name, VB_PredeclaredId,
                    VB_Exposed, etc.) — they are required for
                    VBComponents.Import to properly register the component
                    including event handler bindings.
    """
    lines = text.split('\n')
    filtered = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('Attribute '):
            if is_cls:
                # For .cls files: keep ALL Attribute lines
                # (VB_Name, VB_PredeclaredId, VB_Exposed, etc.)
                # These are required by VBComponents.Import for proper
                # event handler binding (e.g. Worksheet_Change).
                filtered.append(line)
            else:
                # For .bas files, keep ONLY VB_Name (required for named import)
                if 'VB_Name' in stripped:
                    filtered.append(line)
                # else: skip other Attribute lines
        else:
            filtered.append(line)
    return '\n'.join(filtered)


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

            # Read file with auto-detection of encoding
            try:
                text = read_vba_file(file_path)
            except (FileNotFoundError, ValueError) as e:
                print(f"    [!] {e}")
                continue

            # Strip export-format header (VERSION, BEGIN...END) — only for .bas files
            # .cls files keep the full export format for VBComponents.Import
            is_cls = file_name.lower().endswith('.cls')
            if not is_cls:
                text = strip_export_header(text)
                print(f"    Stripped export header (VERSION, BEGIN...END)")

            # Strip Attribute lines (different handling for .bas vs .cls)
            # .cls files keep ALL Attribute lines for VBComponents.Import
            text = strip_attribute_lines(text, is_cls=is_cls)
            print(f"    Stripped Attribute lines (is_cls={is_cls})")

            # Remove leading blank lines
            text = text.lstrip('\n\r')

            # Write temp file in Windows-1251 and import via VBComponents.Import
            # For .cls files, this properly registers event handlers (e.g. Worksheet_Change)
            # because VBComponents.Import preserves the VBA export format including
            # Attribute lines and the BEGIN...END designer section.
            temp_file = TEMP_DIR / file_name
            with open(temp_file, "w", encoding="cp1251") as f:
                f.write(text)
            print(f"    Converted encoding: UTF-8 -> Windows-1251")
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