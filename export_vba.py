#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export all VBA modules from work.xlsm to UTF-8 source files.

Encoding strategy (two-phase model):
  1. Excel exports VBA modules in Windows-1251 encoding
  2. This script converts them to UTF-8 for storage on disk
  3. .gitattributes normalizes encoding for Git
  4. import_all_vba.py converts back to Windows-1251 before importing into Excel

Usage:
    python export_vba.py          # Export all modules
    python export_vba.py --dry    # Dry run: show what would be exported
"""
import os
import sys
import shutil
import win32com.client
from win32com.client import gencache
from pathlib import Path

EXCEL_PATH = Path(r"L:\PROject\SysW\work.xlsm")
PROJECT_DIR = Path(r"L:\PROject\SysW")
TEMP_DIR = Path(r"L:\PROject\SysW\_temp_export")

# Mapping: VBA component name -> output filename
# .bas = standard module, .cls = class module
COMPONENTS = {
    "Mod_Utils": "Mod_Utils.bas",
    "Mod_OrderHeader": "Mod_OrderHeader.bas",
    "Mod_Import": "Mod_Import.bas",
    "Mod_ButtonDispatcher": "Mod_ButtonDispatcher.bas",
    "Mod_FullTestRunner": "Mod_FullTestRunner.bas",
    "Sheet1": "Sheet1_main.cls",
}


def ensure_temp_dir():
    """Create or recreate temporary export directory."""
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(parents=True)
    print(f"Created temp directory: {TEMP_DIR}")


def export_components(excel, vb_project):
    """Export all VBA components from the workbook to temp directory.

    Excel exports in Windows-1251 encoding by default.

    Returns:
        list of (component_name, temp_file_path) tuples
    """
    exported = []
    for comp_name in COMPONENTS:
        try:
            component = vb_project.VBComponents.Item(comp_name)
            temp_file = TEMP_DIR / COMPONENTS[comp_name]
            component.Export(str(temp_file))
            exported.append((comp_name, temp_file))
            print(f"  Exported: {comp_name} -> {temp_file.name}")
        except Exception as e:
            print(f"  [!] Failed to export {comp_name}: {e}")
    return exported


def convert_to_utf8(exported):
    """Convert exported files from Windows-1251 to UTF-8.

    Each exported file is read as CP1251 and written back as UTF-8
    in the project directory.

    Args:
        exported: list of (component_name, temp_file_path) tuples
    """
    for comp_name, temp_file in exported:
        # Read as Windows-1251 (Excel export encoding)
        with open(temp_file, "r", encoding="cp1251") as f:
            text = f.read()

        # Write as UTF-8 to project directory
        output_path = PROJECT_DIR / COMPONENTS[comp_name]
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text)

        print(f"  Converted: {temp_file.name} (CP1251 -> UTF-8) -> {output_path.name}")


def dry_run():
    """Show what would be exported without actually doing it."""
    print("=== DRY RUN ===")
    print(f"Excel file: {EXCEL_PATH}")
    print(f"Project directory: {PROJECT_DIR}")
    print("\nComponents to export:")
    for comp_name, filename in COMPONENTS.items():
        output = PROJECT_DIR / filename
        status = "EXISTS" if output.exists() else "NEW"
        print(f"  {comp_name:25s} -> {filename:30s} [{status}]")
    print("\nNo files were modified.")


def main():
    dry = "--dry" in sys.argv or "--dry-run" in sys.argv
    if dry:
        dry_run()
        return

    if not EXCEL_PATH.exists():
        print(f"Error: Excel file not found: {EXCEL_PATH}")
        sys.exit(1)

    print("Creating Excel COM object...")
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    success = False

    try:
        print(f"Opening workbook: {EXCEL_PATH}")
        workbook = excel.Workbooks.Open(str(EXCEL_PATH))

        print("Accessing VBA project...")
        vb_project = workbook.VBProject
        comp_count = vb_project.VBComponents.Count
        print(f"VBA project accessible. Components count: {comp_count}")

        # Step 1: Create temp directory
        ensure_temp_dir()

        # Step 2: Export all components from Excel to temp (CP1251)
        print("\n--- Exporting components from Excel ---")
        exported = export_components(excel, vb_project)

        if not exported:
            print("No components were exported. Aborting.")
            return

        # Step 3: Convert exported files from CP1251 to UTF-8
        print("\n--- Converting to UTF-8 ---")
        convert_to_utf8(exported)

        print("\n=== EXPORT COMPLETED SUCCESSFULLY ===")
        print(f"Exported {len(exported)} components to {PROJECT_DIR}")
        success = True

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        print("\nPossible causes:")
        print("  1. 'Trust access to the VBA project object model' is not enabled")
        print("     (File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings)")
        print("  2. Make sure work.xlsm is not open in Excel")
        print("  3. Install pywin32: pip install pywin32")
    finally:
        # Clean up temp directory
        if TEMP_DIR.exists():
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
            print(f"Cleaned up: {TEMP_DIR}")

        if workbook is not None:
            print("Closing workbook...")
            workbook.Close()

        print("Closing Excel...")
        excel.Quit()

    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()