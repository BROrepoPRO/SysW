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
  It is imported via VBComponents.Import() like any other .cls file,
  but contains the required Attribute VB_Base line that tells Excel
  to place it in "Microsoft Excel Objects" folder.
"""
import os
import sys
import shutil
import re
import win32com.client
from win32com.client import gencache
from pathlib import Path

EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"
MODULES_PATH = Path(r"L:\PROject\SysW\src")
TEMP_DIR = Path(r"L:\PROject\SysW\_temp_import")

# VBA component type constants
VBEXT_CT_STDMODULE = 1    # Standard module (.bas)
VBEXT_CT_CLASSMODULE = 2  # Class module (.cls)
VBEXT_CT_MSFORM = 3       # Form
VBEXT_CT_DOCUMENT = 100   # Document (Worksheet, Workbook, etc.)

# Dynamically discover all .bas and .cls files in the project root.
# This ensures all VBA modules are imported without needing to update
# the list manually when new modules are added.
FILES = sorted([
    f.name for f in (MODULES_PATH / "modules").iterdir()
    if f.suffix.lower() == ".bas" and f.is_file()
] + [
    f.name for f in (MODULES_PATH / "sheets").iterdir()
    if f.suffix.lower() == ".cls" and f.is_file()
])

# Sheet components that should be imported as Documents (not Class Modules)
# These go into "Microsoft Excel Objects" folder.
# VBComponents.Import() handles this automatically when the .cls file
# contains the correct Attribute VB_Base line with Worksheet CLSID.
# Dynamically discovered: any .cls file whose name starts with "Sheet"
# (English) or "Лист" (Russian) is treated as a worksheet document.
SHEET_COMPONENTS = {
    f.stem for f in (MODULES_PATH / "sheets").iterdir()
    if f.suffix.lower() == ".cls"
    and (
        f.stem.lower().startswith("sheet")
        or f.stem.lower().startswith("лист")
    )
    and f.is_file()
}


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
                    including event handler bindings and document type.
    """
    lines = text.split('\n')
    filtered = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('Attribute '):
            if is_cls:
                # For .cls files: keep ALL Attribute lines
                # (VB_Name, VB_PredeclaredId, VB_Exposed, VB_Base, etc.)
                # These are required by VBComponents.Import for proper
                # event handler binding (e.g. Worksheet_Change) and
                # document type detection (VB_Base with Worksheet CLSID).
                filtered.append(line)
            else:
                # For .bas files, keep ONLY VB_Name (required for named import)
                if 'VB_Name' in stripped:
                    filtered.append(line)
                # else: skip other Attribute lines
        else:
            filtered.append(line)
    return '\n'.join(filtered)


def extract_vb_name(text):
    """Extract VB_Name from VBA source text.

    Looks for 'Attribute VB_Name = "..."' line and returns the name.
    Returns None if not found.
    """
    match = re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', text)
    return match.group(1) if match else None


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
        # Read VB_Name from each source file to ensure correct component
        # identification (file name stem may differ from VB_Name, e.g.
        # Лист2_main.cls -> VB_Name = "Лист2")
        print("")
        print("--- Removing existing components ---")
        for file_name in FILES:
            if file_name.lower().endswith('.cls'):
                file_path = MODULES_PATH / "sheets" / file_name
            else:
                file_path = MODULES_PATH / "modules" / file_name

            if not file_path.exists():
                print(f"  Not found on disk: {file_name}")
                continue

            # Read VB_Name from source file
            try:
                text = read_vba_file(file_path)
            except (FileNotFoundError, ValueError) as e:
                print(f"  [!] {e}")
                continue

            vb_name = extract_vb_name(text)
            if not vb_name:
                # Fallback to file stem if VB_Name not found
                vb_name = Path(file_name).stem
                print(f"  VB_Name not found in {file_name}, using stem: {vb_name}")

            try:
                existing = vb_project.VBComponents.Item(vb_name)
                if existing:
                    vb_project.VBComponents.Remove(existing)
                    print(f"  Removed: {vb_name} (from {file_name})")
            except Exception:
                print(f"  Not found: {vb_name} (from {file_name})")

        # Second pass: import all components
        print("")
        print("--- Importing components ---")
        for file_name in FILES:
            if file_name.lower().endswith('.cls'):
                file_path = MODULES_PATH / "sheets" / file_name
            else:
                file_path = MODULES_PATH / "modules" / file_name
            print("")
            print(f"  Processing: {file_name}")

            if not file_path.exists():
                print(f"    [!] File not found: {file_path}")
                continue

            component_name = file_path.stem
            is_cls = file_name.lower().endswith('.cls')
            is_sheet = component_name in SHEET_COMPONENTS

            # Read file with auto-detection of encoding
            try:
                text = read_vba_file(file_path)
            except (FileNotFoundError, ValueError) as e:
                print(f"    [!] {e}")
                continue

            # Strip export-format header (VERSION, BEGIN...END)
            # For .bas files: remove the header (not valid VBA code)
            # For .cls files: remove the header too — VBComponents.Import
            #   does NOT need VERSION/BEGIN/END lines, it only needs
            #   the Attribute lines and the actual code.
            text = strip_export_header(text)
            print(f"    Stripped export header (VERSION, BEGIN...END)")

            # Strip/keep Attribute lines (different handling for .bas vs .cls)
            # .cls files keep ALL Attribute lines for VBComponents.Import
            # including VB_Base which determines document type
            text = strip_attribute_lines(text, is_cls=is_cls)
            print(f"    Stripped Attribute lines (is_cls={is_cls})")

            # Remove leading blank lines
            text = text.lstrip('\n\r')

            # Write temp file in Windows-1251 and import via VBComponents.Import
            # VBComponents.Import properly registers Public Subs as macros.
            # For .cls files with VB_Base attribute, it also correctly
            # places the component in "Microsoft Excel Objects" folder.
            # On Russian Windows, cp1251 is the correct encoding for VBA.
            # Use errors='replace' to handle characters not representable
            # in cp1251 (e.g. em-dash, typographic quotes) — they become '?'.
            temp_file = TEMP_DIR / file_name
            try:
                with open(temp_file, "w", encoding="cp1251", errors="replace") as f:
                    f.write(text)
            except Exception as e:
                print(f"    [!] Failed to write temp file: {e}")
                continue
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