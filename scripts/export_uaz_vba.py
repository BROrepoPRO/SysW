#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export all VBA modules from UAZ.xlsm to see what macros are inside.
Based on the existing export_vba.py pattern.
"""
import os
import sys
import shutil
import win32com.client
from win32com.client import gencache
from pathlib import Path

UAZ_PATH = Path(r"L:\PROject\SysW\base\models\UAZ.xlsm")
OUTPUT_DIR = Path(r"L:\PROject\SysW\scripts\_uaz_vba_export")


def main():
    if not UAZ_PATH.exists():
        print(f"Error: File not found: {UAZ_PATH}")
        sys.exit(1)

    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)
    print(f"Created output directory: {OUTPUT_DIR}")

    print("Creating Excel COM object...")
    excel = gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    success = False

    try:
        print(f"Opening workbook: {UAZ_PATH}")
        workbook = excel.Workbooks.Open(str(UAZ_PATH))

        print("Accessing VBA project...")
        vb_project = workbook.VBProject
        comp_count = vb_project.VBComponents.Count
        print(f"VBA components found: {comp_count}")

        for comp in vb_project.VBComponents:
            comp_name = comp.Name
            comp_type = comp.Type
            type_names = {1: "StdModule (.bas)", 2: "ClassModule (.cls)", 3: "MSForm (.frm)", 100: "Document"}
            print(f"\n  Component: {comp_name} ({type_names.get(comp_type, f'type={comp_type}')})")

            ext = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(comp_type, ".bas")
            out_file = OUTPUT_DIR / f"{comp_name}{ext}"
            comp.Export(str(out_file))
            print(f"    Exported -> {out_file.name}")

            # Read and print first 20 lines to understand the code
            with open(out_file, "r", encoding="cp1251") as f:
                lines = f.readlines()
            print(f"    Lines of code: {len(lines)}")
            print(f"    --- First 30 lines ---")
            for line in lines[:30]:
                print(f"    {line.rstrip()}")

        print(f"\n=== All components exported to {OUTPUT_DIR} ===")
        success = True

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
    finally:
        if workbook is not None:
            workbook.Close()
        excel.Quit()

    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()