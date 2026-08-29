#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Прогон бизнес-теста П1 на реальных данных в work.xlsm через COM-автоматизацию Excel.

Сценарий (3 макроса):
  1. ImportFromB2_UI  — перенос данных с листа {номер}M на main (L:N работы, X:AA ЗЧ)
  2. AutoMatchWorks    — автоподбор работ (E:K), ключ — наименование L
  3. AutoMatchParts    — автоподбор ЗЧ (Q:V), ключ — № кат. X

Никакой код/данные юзера не модифицируются; только выполнение макросов,
сбор результатов и сохранение книги для просмотра юзером.
"""
import sys
import time
import os
import gc
import win32com.client
import win32process

from config import WORKBOOK_PATH, LOGS_DIR

WORKBOOK = str(WORKBOOK_PATH)
REPORT_FILE = str(LOGS_DIR / "p1_test_report.log")

# Константы колонок листа main (совпадают с Mod_Constants / Mod_AutoMatch)
C = {
    "W_ARTICLE": 5,     # E  — Артикул (результат)
    "W_NAME": 6,        # F  — Наименование (результат)
    "W_NORMHOURS": 7,   # G  — Кол-во н/ч (результат)
    "W_QTY": 8,         # H  — Кол-во оп (результат)
    "W_PRICE": 9,       # I  — Цена н/ч (результат)
    "W_SUM": 10,        # J  — Сумма (формула)
    "W_IN_NAME": 12,    # L  — Наименование (входящее)
    "W_IN_QTY": 13,     # M  — Кол. оп.
    "W_IN_TOTAL": 14,   # N  — Всего
    "P_ARTICLE": 17,    # Q  — Артикул (результат)
    "P_NAME": 18,       # R  — Наименование (результат)
    "P_QTY": 20,        # T  — Кол-во (результат)
    "P_PRICE": 21,      # U  — Цена (результат)
    "P_SUM": 22,        # V  — Сумма (формула)
    "P_IN_CATNUM": 24,  # X  — № кат. (входящее)
    "P_IN_NAME": 25,    # Y  — Наименование (входящее)
    "P_IN_QTY": 26,     # Z  — Кол-во
    "P_IN_TOTAL": 27,   # AA — Всего
    "P_MODEL_AB": 28,   # AB — модельный артикул (подстановка при импорте)
    "W_MODEL_O": 15,    # O  — модельный артикул работы
}


def write_log(message: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    print(line, flush=True)
    with open(REPORT_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def val(cell):
    """Безопасное чтение значения ячейки в строку."""
    try:
        v = cell.Value
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def snapshot_main(ws, label):
    """Снимает состояние main по работам (L:N + E:K) и запчастям (X:AA + Q:V)."""
    def rows(incol, resultcols):
        last = ws.Cells(ws.Rows.Count, incol).End(-4162).Row  # xlUp
        if last < 4:
            return []
        out = []
        for i in range(4, last + 1):
            key = val(ws.Cells(i, incol))
            if key == "":
                continue
            rec = {"row": i, "in_key": key}
            for name, col in resultcols.items():
                rec[name] = val(ws.Cells(i, col))
            out.append(rec)
        return out

    works = rows(C["W_IN_NAME"], {
        "in_qty": C["W_IN_QTY"], "in_total": C["W_IN_TOTAL"],
        "art": C["W_ARTICLE"], "name": C["W_NAME"],
        "normhours": C["W_NORMHOURS"], "qty": C["W_QTY"],
        "price": C["W_PRICE"], "sum": C["W_SUM"],
        "model_art": C["W_MODEL_O"],
    })
    parts = rows(C["P_IN_CATNUM"], {
        "in_name": C["P_IN_NAME"], "in_qty": C["P_IN_QTY"], "in_total": C["P_IN_TOTAL"],
        "art": C["P_ARTICLE"], "name": C["P_NAME"],
        "qty": C["P_QTY"], "price": C["P_PRICE"], "sum": C["P_SUM"],
        "model_ab": C["P_MODEL_AB"],
    })
    write_log(f"[СНИМОК:{label}] Работ: {len(works)}, ЗЧ: {len(parts)}")
    w_found = sum(1 for r in works if r["art"] != "" and r["art"] != "НЕ НАЙДЕНО")
    w_nf = sum(1 for r in works if r["art"] == "НЕ НАЙДЕНО" or r["in_key"] == "НЕ НАЙДЕНО")
    p_found = sum(1 for r in parts if r["art"] != "" and r["art"] != "НЕ НАЙДЕНО")
    p_nf = sum(1 for r in parts if r["art"] == "НЕ НАЙДЕНО" or r["in_key"] == "НЕ НАЙДЕНО")
    write_log(f"[СНИМОК:{label}] Работы найдено={w_found}, не найдено={w_nf}; "
              f"ЗЧ найдено={p_found}, не найдено={p_nf}")
    for r in works:
        write_log(f"[СНИМОК:{label}] РАБОТА row={r['row']} | вх={r['in_key']!r} | "
                  f"арт={r['art']!r} | имя={r['name']!r} | нч={r['normhours']!r} | "
                  f"кол={r['qty']!r} | цена={r['price']!r} | сум={r['sum']!r} | O={r['model_art']!r}")
    for r in parts:
        write_log(f"[СНИМОК:{label}] ЗЧ row={r['row']} | №кат={r['in_key']!r} | "
                  f"имя_вх={r['in_name']!r} | арт={r['art']!r} | имя={r['name']!r} | "
                  f"кол={r['qty']!r} | цена={r['price']!r} | сум={r['sum']!r} | AB={r['model_ab']!r}")
    return works, parts


def inject_silence(wb, excel):
    """Вставляет временный VBA-модуль для подавления MsgBox в авто-прогоне.

    Устанавливает ВСЕ флаги подавления, используемые модулями:
      - Mod_Constants.SilenceMsgBox  — глобальный (Mod_AutoMatch, Mod_PickWork и др.)
      - Mod_Import.SilenceMsgBox     — локальный (собственный флаг Mod_Import)

    Mod_Import имеет собственный флаг SilenceMsgBox (строка 6 модуля), поэтому
    без его установки модальные диалоги импорта зависают headless-процесс.
    Возвращает компонент модуля или None, если доступ к VBProject запрещён."""
    try:
        vbproj = wb.VBProject
        comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = "zzTestHelper"
        comp.CodeModule.AddFromString(
            "Public Sub zzSetSilence()\n"
            "    Mod_Constants.SilenceMsgBox = True\n"
            "    Mod_Import.SilenceMsgBox = True\n"
            "End Sub\n"
            "Public Sub zzUnsetSilence()\n"
            "    Mod_Constants.SilenceMsgBox = False\n"
            "    Mod_Import.SilenceMsgBox = False\n"
            "End Sub\n"
        )
        return comp
    except Exception as exc:
        write_log(f"[!] Не удалось внедрить помощник SilenceMsgBox: {exc}")
        return None


def run_macro(excel, name):
    """Запуск макроса через Application.Run с логированием и возвратом результата."""
    write_log(f"[>] Запуск макроса {name} ...")
    try:
        excel.Application.Run(name)
        write_log(f"[OK] Макрос {name} завершён без исключения.")
        return True, None
    except Exception as exc:
        write_log(f"[!] ОШИБКА макроса {name}: {exc}")
        return False, str(exc)


def main():
    os.makedirs(str(LOGS_DIR), exist_ok=True)
    # Чистим лог от предыдущего запуска
    if os.path.exists(REPORT_FILE):
        os.remove(REPORT_FILE)

    write_log("=" * 70)
    write_log("ПРОГОН БИЗНЕС-ТЕСТА П1 (реальные данные work.xlsm)")
    write_log("=" * 70)

    excel = None
    workbook = None
    helper = None
    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # PID для безопасного завершения зависшего процесса
        try:
            hwnd = excel.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            pid_file = LOGS_DIR / "excel_pid_p1.txt"
            pid_file.parent.mkdir(parents=True, exist_ok=True)
            pid_file.write_text(str(pid), encoding="utf-8")
            write_log(f"PID Excel={pid}")
        except Exception:
            pass

        # Открытие с ретраями (паттерн проекта)
        wb = None
        for attempt in range(1, 6):
            try:
                wb = excel.Workbooks.Open(WORKBOOK, ReadOnly=False, UpdateLinks=0)
            except Exception as exc:
                write_log(f"    Попытка {attempt}/5: ошибка открытия ({exc}); пауза 3с")
                wb = None
                time.sleep(3)
                gc.collect()
                continue
            if wb is not None:
                break
            time.sleep(3)
            gc.collect()
        if wb is None:
            raise RuntimeError("Не удалось открыть work.xlsm")
        workbook = wb
        write_log("work.xlsm открыт.")

        # Автоматический пересчёт формул
        excel.Calculation = -4105  # xlCalculationAutomatic

        ws_main = workbook.Sheets("main")

        # ---------- ШАГ 1. Проверка значений ----------
        write_log("-" * 70)
        write_log("[ШАГ 1] Значения на листе main:")
        b4 = val(ws_main.Range("B4"))
        b14 = val(ws_main.Range("B14"))
        b13 = val(ws_main.Range("B13"))
        write_log(f"  B4 (номер заказа)  = {b4!r}")
        write_log(f"  B14 (группа модели)= {b14!r}")
        write_log(f"  B13 (цена н/ч)     = {b13!r}")

        sheet_names = [s.Name for s in workbook.Sheets]
        write_log(f"  Листы книги ({len(sheet_names)}): {sheet_names}")
        src_name = str(b4).strip() + "M"
        src_sheet = None
        for s in workbook.Sheets:
            if s.Name.strip().upper() == src_name.upper():
                src_sheet = s
                break
        if src_sheet is None:
            write_log(f"[!] Лист {src_name} НЕ найден в книге. "
                      "ImportFromB2_UI попробует взять его из report.xlsx.")
        else:
            write_log(f"  Лист-источник {src_sheet.Name} найден в книге.")

        # ---------- Подготовка тишины (SilenceMsgBox) ----------
        helper = inject_silence(workbook, excel)
        if helper is not None:
            try:
                excel.Application.Run("zzSetSilence")
                write_log("SilenceMsgBox=True установлен (тестовый режим без диалогов).")
            except Exception as exc:
                write_log(f"[!] Не удалось выставить SilenceMsgBox: {exc}")
        else:
            write_log("[!] ВНИМАНИЕ: SilenceMsgBox НЕ выставлен. "
                      "Макросы могут заблокироваться на MsgBox-диалогах.")

        # ---------- МАКРОС 1. Импорт ----------
        write_log("=" * 70)
        ok1, err1 = run_macro(excel, "ImportFromB2_UI")
        snapshot_main(ws_main, "ПОСЛЕ-ИМПОРТА")

        # ---------- МАКРОС 2. Автоподбор работ ----------
        write_log("=" * 70)
        ok2, err2 = run_macro(excel, "AutoMatchWorks")
        snapshot_main(ws_main, "ПОСЛЕ-AUTOMATCHWORKS")

        # ---------- МАКРОС 3. Автоподбор ЗЧ ----------
        write_log("=" * 70)
        ok3, err3 = run_macro(excel, "AutoMatchParts")
        snapshot_main(ws_main, "ПОСЛЕ-AUTOMATCHPARTS")

        # ---------- Итоговые показатели ----------
        write_log("=" * 70)
        write_log("[ИТОГ] Расчёт итоговых количеств по листу main:")
        works, parts = snapshot_main(ws_main, "ФИНАЛ")
        w_total = len(works)
        w_found = sum(1 for r in works if r["art"] != "" and r["art"] != "НЕ НАЙДЕНО")
        w_nf = w_total - w_found
        p_total = len(parts)
        p_found = sum(1 for r in parts if r["art"] != "" and r["art"] != "НЕ НАЙДЕНО")
        p_nf = p_total - p_found
        write_log(f"  Работы: всего={w_total}, найдено={w_found}, НЕ НАЙДЕНО={w_nf}")
        write_log(f"  ЗЧ:     всего={p_total}, найдено={p_found}, НЕ НАЙДЕНО={p_nf}")

        # Подставленные артикулы работ
        arts_w = sorted({r["art"] for r in works if r["art"] not in ("", "НЕ НАЙДЕНО")})
        arts_p = sorted({r["art"] for r in parts if r["art"] not in ("", "НЕ НАЙДЕНО")})
        write_log(f"  Артикулы работ: {arts_w}")
        write_log(f"  Артикулы ЗЧ:    {arts_p}")

        write_log("=" * 70)
        write_log(f"РЕЗУЛЬТАТЫ МАКРОСОВ: Import={ok1}({err1}), "
                  f"AutoMatchWorks={ok2}({err2}), AutoMatchParts={ok3}({err3})")

        # ---------- Сохранение книги для просмотра юзером ----------
        # Сначала убираем временный модуль, чтобы не оставлять изменений кода.
        if helper is not None:
            try:
                excel.Application.Run("zzUnsetSilence")
            except Exception:
                pass
            try:
                workbook.VBProject.VBComponents.Remove(helper)
                write_log("Временный VBA-модуль-помощник удалён (код книги не изменён).")
            except Exception as exc:
                write_log(f"[!] Не удалось удалить временный модуль: {exc}")

        workbook.Save()
        write_log("Книга сохранена (work.xlsm). Данные main оставлены для просмотра юзером.")

        workbook.Close(SaveChanges=False)
        excel.Quit()
        excel = None
        write_log("Excel аккуратно закрыт.")
        write_log("ПРОГОН П1 ЗАВЕРШЁН.")

    finally:
        # Аварийная очистка при сбое
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()