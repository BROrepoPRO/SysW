#!/usr/bin/env python3
"""
Run RunMinimalTests macro in work.xlsm via Excel COM.
Shows Excel visible so user can see the MsgBox result.
"""
import sys
import time
import win32com.client
from win32com.client import gencache

EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"


def main():
    print("=" * 60)
    print("ZAPUSK TESTOV")
    print("=" * 60)
    print()
    print("[1/4] Sozdanie COM-obekta Excel...")
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    workbook = None

    try:
        print("[2/4] Otkrytie work.xlsm...")
        workbook = excel.Workbooks.Open(EXCEL_PATH)

        print("[3/4] Zapusk makrosa RunMinimalTests...")
        print()
        print("    Ozhidanie zaversheniya testov...")
        print("    Poyavitsya okno s rezultatami (MsgBox).")
        print("    Zakrojte okno MsgBox dlya prodolzheniya.")
        print()
        excel.Run("RunMinimalTests")

        print("[4/4] Makros vypolnen!")
        print()
        print("=" * 60)
        print("REZULTAT: Makros RunMinimalTests zapushchen!")
        print("=" * 60)
        print()
        print("Posle prosmotra rezultatov zakrojte Excel vruchnuyu.")

    except Exception as e:
        print(f"Oshibka: {e}", file=sys.stderr)
    finally:
        print()
        print("Excel ostavlen otkrytym dlya prosmotra rezultatov.")
        print("Zakrojte Excel kogda budete gotovy.")


if __name__ == "__main__":
    main()