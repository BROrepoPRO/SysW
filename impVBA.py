#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Import all VBA modules (bas/cls) into work.xlsm via Excel COM.

Encoding strategy (two-phase model):
  1. Files on disk are stored in UTF-8 (for VS Code and Git)
  2. This script converts them to Windows-1251 before importing into Excel
  3. .gitattributes normalizes encoding for Git

Sheet handling:
  Sheet1_main.cls is a Worksheet document (not a Class Module).
  It must be imported using VBComponents.Add with type vbext_ct_Document (100)
  to place it in "Microsoft Excel Objects" folder, not "Class Modules".
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

# VBA component type constants
VBEXT_CT_STDMODULE = 1    # Standard module (.bas)
VBEXT_CT_CLASSMODULE = 2  # Class module (.cls)
VBEXT_CT_MSFORM = 3       # Form
VBEXT_CT_DOCUMENT = 100   # Document (Worksheet, Workbook, etc.)

FILES = [
    "Mod_Utils.bas",
    "Mod_OrderHeader.bas",
    "Mod_Import.bas",
    "Mod_ButtonDispatcher.bas",
    "Mod_FullTestRunner.bas",
    "Sheet1_main.cls",
]

# Sheet components that should be imported as Documents (not Class Modules)
# These go into "Microsoft Excel Objects" folder
SHEET_COMPONENTS = {"Sheet1_main"}


def read_vba_file(file_path):
    """Reads a VBA file with auto-detection of encoding.

    Priority: UTF-8 with BOM > UTF-8 > Windows-1251

    Returns:
        str: file contents as Unicode

    Raises:
        FileNotFoundError: if file does not exist
        ValueError: if encoding cannot be determined
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Check BOM (Byte Order Mark)
    with open(file_path, "rb") as f:
        raw = f.read(4)

    if raw[:3] == b'\xef\xbb\xbf':
        # UTF-8 with BOM
        with open(file_path, "r", encoding="utf-8-sig") as f:
            text = f.read()
        print(f"    Read encoding: UTF-8 with BOM")
        return text

    # Try UTF-8
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            text = f.read()
        print(f"    Read encoding: UTF-8")
        return text
    except UnicodeDecodeError:
        pass

    # Fallback to Windows-1251
    try:
        with open(file_path, "r", encoding="cp1251") as f:
            text = f.read()
        print(f"    Read encoding: Windows-1251 (fallback). "
              f"Recommend converting to UTF-8: {file_path.name}")
        return text
    except UnicodeDecodeError:
        # Last attempt: UTF-8 with BOM (if BOM is corrupted)
        try:
            with open(file_path, "r", encoding="utf-8-sig") as f:
                text = f.read()
            print(f"    Read encoding: UTF-8 with BOM (fallback)")
            return text
        except UnicodeDecodeError as e:
            raise ValueError(f"Cannot determine encoding for {file_path}: {e}")


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


def get_vba_code_from_cls(text):
    """Extract pure VBA code from a .cls file, removing export header and Attribute lines.

    This is used for importing Sheet components as Documents via AddFromString.
    The export header (VERSION, BEGIN...END) and Attribute lines are not valid
    VBA code and must be removed before passing to AddFromString.

    Returns:
        str: pure VBA code (Option Explicit, Subs, Functions, etc.)
    """
    # Remove export header (VERSION, BEGIN...END)
    text = strip_export_header(text)

    # Remove ALL Attribute lines (they are set automatically by Excel)
    lines = text.split('\n')
    filtered = []
    for line in lines:
        stripped = line.strip()
        if not stripped.startswith('Attribute '):
            filtered.append(line)
    return '\n'.join(filtered)


def import_sheet_component(vb_project, file_path, component_name):
    """Import a Sheet component as a Document (Worksheet) into VBA project.

    Sheet components must be added as vbext_ct_Document (type 100) to
    appear in "Microsoft Excel Objects" folder, not "Class Modules".

    Args:
        vb_project: VBA project object
        file_path: Path to the .cls file
        component_name: Name of the component (e.g. "Sheet1_main")

    Returns:
        The imported component object, or None on failure
    """
    print(f"    Importing as Document (Worksheet)...")

    # Read the file
    try:
        text = read_vba_file(file_path)
    except (FileNotFoundError, ValueError) as e:
        print(f"    [!] {e}")
        return None

    # Extract pure VBA code (remove header and attributes)
    vba_code = get_vba_code_from_cls(text)
    vba_code = vba_code.lstrip('\n\r')

    # Create temp file in UTF-8 for the code
    temp_file = TEMP_DIR / f"{component_name}_code.txt"
    with open(temp_file, "w", encoding="utf-8") as f:
        f.write(vba_code)
    print(f"    Using UTF-8 encoding")

    # Add a new Document component
    # vbext_ct_Document = 100 creates a document module (Worksheet-like)
    try:
        new_comp = vb_project.VBComponents.Add(100)  # vbext_ct_Document
        print(f"    Created Document component")
    except Exception as e:
        print(f"    [!] Failed to create Document component: {e}")
        # Fallback: try vbext_ct_MSForm (3)
        try:
            new_comp = vb_project.VBComponents.Add(3)  # vbext_ct_MSForm
            print(f"    Created MSForm component (fallback)")
        except Exception as e2:
            print(f"    [!] Failed to create MSForm component: {e2}")
            return None

    # Rename to the target name
    try:
        new_comp.Name = component_name
        print(f"    Renamed to: {component_name}")
    except Exception as e:
        print(f"    [!] Failed to rename component: {e}")
        return None

    # Import the code from temp file
    try:
        # Read the UTF-8 temp file
        with open(temp_file, "r", encoding="utf-8") as f:
            code_text = f.read()
        new_comp.CodeModule.AddFromString(code_text)
        print(f"    [+] Successfully imported code: {file_path.name}")
        return new_comp
    except Exception as e:
        print(f"    [!] Failed to import code: {e}")
        return None


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
            is_cls = file_name.lower().endswith('.cls')
            is_sheet = component_name in SHEET_COMPONENTS

            if is_sheet:
                # Import Sheet as Document (Worksheet)
                result = import_sheet_component(vb_project, file_path, component_name)
                if result is None:
                    print(f"    [!] Failed to import sheet: {file_name}")
                continue

            # Read file with auto-detection of encoding
            try:
                text = read_vba_file(file_path)
            except (FileNotFoundError, ValueError) as e:
                print(f"    [!] {e}")
                continue

            # Strip export-format header (VERSION, BEGIN...END) — only for .bas files
            # .cls files keep the full export format for VBComponents.Import
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
            # VBComponents.Import properly registers Public Subs as macros.
            # On Russian Windows, cp1251 is the correct encoding for VBA.
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