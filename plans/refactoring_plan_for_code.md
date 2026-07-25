# План рефакторинга SysW для режима Code

> **Дата:** 2026-07-25
> **Основание:** `plans/next_promt.md`, `plans/analysis_current_state.md`
> **Режим выполнения:** Code
> **Порядок выполнения:** строго последовательный, этап за этапом

---

## Этап 1: Создание конфигурационных файлов путей

**Цель:** Создать единую точку конфигурации путей для Python и PowerShell скриптов.

**Почему нет `config.bas`:** VBA работает внутри Excel и может использовать `ThisWorkbook.Name` / `ThisWorkbook.Path` для получения корректных значений. Вынос в отдельный конфиг создал бы лишнюю движущуюся часть без выгоды.

### 1.1 Создать `scripts/config.py`

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Общая конфигурация путей для Python-скриптов SysW.
Все скрипты должны импортировать пути из этого модуля,
чтобы избежать дублирования абсолютных путей.
"""
from pathlib import Path

# Корень проекта — родительская директория scripts/
PROJECT_DIR = Path(__file__).resolve().parent.parent

# Основной Excel-файл
WORKBOOK_PATH = PROJECT_DIR / "work.xlsm"

# Директория исходников VBA
SRC_DIR = PROJECT_DIR / "src"

# Временные директории
TEMP_EXPORT_DIR = PROJECT_DIR / "_temp_export"
TEMP_IMPORT_DIR = PROJECT_DIR / "_temp_import"

# Файл лога тестов
TEST_LOG_FILE = PROJECT_DIR / "test_results.log"
```

### 1.2 Создать `scripts/config.ps1`

```powershell
# scripts/config.ps1
# Общая конфигурация путей для PowerShell-скриптов SysW.
# Все скрипты должны dot-source этот файл:
#   . "$PSScriptRoot\config.ps1"

# Корень проекта — родительская директория scripts/
$Script:ProjectRoot = Resolve-Path "$PSScriptRoot\.."

# Основной Excel-файл
$Script:WorkbookPath = Join-Path $Script:ProjectRoot "work.xlsm"

# Директория исходников VBA
$Script:SrcDir = Join-Path $Script:ProjectRoot "src"
```

---

## Этап 2: Исправление жёстких привязок в Python-скриптах

**Цель:** Заменить все абсолютные пути `L:\PROject\SysW\...` на относительные через `config.py`.

### 2.1 Исправить `scripts/export_vba.py`

**Файл:** [`scripts/export_vba.py`](scripts/export_vba.py)

**Изменения:**
- Строки 23-25: заменить
  ```python
  EXCEL_PATH = Path(r"L:\PROject\SysW\work.xlsm")
  PROJECT_DIR = Path(r"L:\PROject\SysW\src")
  TEMP_DIR = Path(r"L:\PROject\SysW\_temp_export")
  ```
  на
  ```python
  from config import WORKBOOK_PATH, SRC_DIR, TEMP_EXPORT_DIR
  EXCEL_PATH = WORKBOOK_PATH
  PROJECT_DIR = SRC_DIR
  TEMP_DIR = TEMP_EXPORT_DIR
  ```

### 2.2 Исправить `scripts/impVBA.py`

**Файл:** [`scripts/impVBA.py`](scripts/impVBA.py)

**Изменения:**
- Строки 25-27: заменить
  ```python
  EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"
  MODULES_PATH = Path(r"L:\PROject\SysW\src")
  TEMP_DIR = Path(r"L:\PROject\SysW\_temp_import")
  ```
  на
  ```python
  from config import WORKBOOK_PATH, SRC_DIR, TEMP_IMPORT_DIR
  EXCEL_PATH = str(WORKBOOK_PATH)
  MODULES_PATH = SRC_DIR
  TEMP_DIR = TEMP_IMPORT_DIR
  ```

### 2.3 Исправить `scripts/run_tests.py`

**Файл:** [`scripts/run_tests.py`](scripts/run_tests.py)

**Изменения:**
- Строка 14: заменить
  ```python
  EXCEL_PATH = r"L:\PROject\SysW\work.xlsm"
  ```
  на
  ```python
  from config import WORKBOOK_PATH, TEST_LOG_FILE
  EXCEL_PATH = str(WORKBOOK_PATH)
  ```
- Строка 15: заменить
  ```python
  LOG_FILE = os.path.join(os.path.dirname(EXCEL_PATH), "test_results.log")
  ```
  на
  ```python
  LOG_FILE = str(TEST_LOG_FILE)
  ```

---

## Этап 3: Исправление жёстких привязок в PowerShell-скриптах

**Цель:** Заменить все абсолютные пути на относительные через `config.ps1`.

### 3.1 Исправить `scripts/Set-ExcelTrust.ps1`

**Файл:** [`scripts/Set-ExcelTrust.ps1`](scripts/Set-ExcelTrust.ps1)

**Изменения:**
- После строки 22 (`$ErrorActionPreference = "Stop"`) добавить:
  ```powershell
  . "$PSScriptRoot\config.ps1"
  ```
- Строки 25-26: заменить
  ```powershell
  $ProjectPath = "L:\PROject\SysW"
  $ExcelPath = "L:\PROject\SysW\work.xlsm"
  ```
  на
  ```powershell
  $ProjectPath = $Script:ProjectRoot
  $ExcelPath = $Script:WorkbookPath
  ```

---

## Этап 4: Исправление жёстких привязок в VBA-коде

**Цель:** Убрать жёсткую привязку к имени файла `work.xlsm`.

**Подход:** Вместо создания `config.bas` используем `ThisWorkbook.Name` — VBA работает внутри Excel и гарантированно знает собственное имя файла. Это безопаснее и не требует синхронизации конфига.

### 4.1 Исправить `Mod_MainButtons.bas`

**Файл:** [`src/modules/Mod_MainButtons.bas`](src/modules/Mod_MainButtons.bas)

**Изменения:**
- Строка 158: заменить
  ```vba
  mapping.Add "ИМПОРТ ВХ", "'work.xlsm'!Btn_main_ImportVH_Click"
  ```
  на
  ```vba
  mapping.Add "ИМПОРТ ВХ", "'" & ThisWorkbook.Name & "'!Btn_main_ImportVH_Click"
  ```

**Важно:** Функция `AssignMainButtons` (строка 155) не вызывается нигде в коде. Она помечена как мёртвый код (D8). Изменение вносится на случай, если функция будет использоваться в будущем.

---

## Этап 5: Чистка мёртвого кода в VBA-модулях

**Цель:** Удалить заглушки, дубликаты и неиспользуемые константы.

### 5.1 Удалить заглушки из `Mod_SheetButtons.bas`

**Файл:** [`src/modules/Mod_SheetButtons.bas`](src/modules/Mod_SheetButtons.bas)

**Удалить:**
- Строки 13-41: процедуры `Btn_z4_Action1`, `Btn_z4_Action2`, `Btn_z4_Action3` (заглушки)
- Строки 47-75: процедуры `Btn_work_Action1`, `Btn_work_Action2`, `Btn_work_Action3` (заглушки)

**Оставить:** процедуры поиска `Btn_UAZ_SearchByArticle`, `Btn_UAZ_SearchByName`, `Btn_UAZ_ClearFilter` и вспомогательные функции.

### 5.2 Удалить заглушку и дубликат из `Mod_MainButtons.bas`

**Файл:** [`src/modules/Mod_MainButtons.bas`](src/modules/Mod_MainButtons.bas)

**Удалить:**
- Строки 86-94: процедура `Btn_main_MANz4` (заглушка)
- Строки 121-145: процедура `Btn_main_ImportVH_Click` (дубликат `Mod_ButtonDispatcher.Btn_main_ImportVH_Click`)
- Строки 147-159: процедура `AssignMainButtons` (не вызывается)

**Оставить:** `Btn_main_Import`, `Btn_main_AUTOz4`, `Btn_main_AUTOw`, `Btn_main_MANWRK`.

### 5.3 Удалить неиспользуемые константы из `Mod_Constants.bas`

**Файл:** [`src/modules/Mod_Constants.bas`](src/modules/Mod_Constants.bas)

**Удалить строки 86-89:**
```vba
Public Const MAIN_HEADER_END_ROW As Long = 17
Public Const MAIN_CLEAR_START_ROW As Long = 4
Public Const MAIN_HEADER_RANGE As String = "B4:B17"
```

**Оставить:** `MAIN_HEADER_START_ROW` (используется), `MAIN_DATA_START_ROW` (используется).

---

## Этап 6: Удаление неиспользуемых скриптов

**Цель:** Убрать из системы все невостребованные скрипты.

**Порядок удаления (строго в указанном порядке):**

| № | Файл | Причина |
|---|------|---------|
| 6.1 | `scripts/export_uaz_vba.ps1` | Одноразовый экспорт VBA из UAZ.xlsm |
| 6.2 | `scripts/export_uaz_vba.py` | То же на Python |
| 6.3 | `scripts/parse_excel.ps1` | Одноразовый парсинг UAZ.xlsm |
| 6.4 | `scripts/parse_full.ps1` | Одноразовый полный парсинг |
| 6.5 | `scripts/parse_full.py` | Одноразовый парсинг (openpyxl) |
| 6.6 | `scripts/rewire_uaz_buttons.ps1` | После удаления work26 не нужен |
| 6.7 | `scripts/Import-VbaFromExcel.ps1` | Дубликат `export_vba.py` |

**Важно:** Соблюдать правило Z2 (.ycarules) — не удалять без подтверждения пользователя. Сначала показать список, получить подтверждение.

---

## Этап 7: Удаление work26 из системы

**Цель:** Полностью удалить встроенную книгу work26.xlsm из UAZ.xlsm.

### 7.1 Создать скрипт удаления work26

Создать `scripts/remove_work26.py`:

```python
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
```

### 7.2 Удалить директорию `scripts/_uaz_vba_export/`

После выполнения скрипта 7.1 удалить всю директорию `scripts/_uaz_vba_export/` вместе с `work26/`.

---

## Этап 8: Исправление виртуального окружения Python

**Цель:** Терминал VS Code должен автоматически активировать `.venv`.

### 8.1 Проверить наличие `.venv`

```powershell
# В терминале VS Code:
python -m venv .venv
.venv\Scripts\pip install pywin32 openpyxl
```

### 8.2 Проверить настройки VS Code

Файл: [`.vscode/settings.json`](.vscode/settings.json)

Текущие настройки (уже корректны):
```json
{
    "python.defaultInterpreterPath": "${workspaceFolder}/.venv/Scripts/python.exe",
    "python.terminal.activateEnvironment": true,
    "terminal.integrated.defaultProfile.windows": "SourceCraft",
    "terminal.integrated.profiles.windows": {
        "SourceCraft": {
            "path": "pwsh.exe",
            "args": [
                "-NoExit",
                "-Command",
                "$env:Path = '.venv\\Scripts;' + $env:Path"
            ]
        }
    }
}
```

Если `.venv` отсутствует — создать и установить зависимости.

---

## Этап 9: Тестирование

**Цель:** Убедиться, что после всех изменений система работает корректно.

### 9.1 Импортировать изменённые VBA-модули в work.xlsm

```powershell
python scripts/impVBA.py
```

### 9.2 Запустить тесты

```powershell
python scripts/run_tests.py
```

Ожидаемый результат: все тесты PASS (TC-01..TC-14).

### 9.3 Проверить экспорт

```powershell
python scripts/export_vba.py
```

Убедиться, что экспорт завершается без ошибок.

---

## Этап 10: Фиксация изменений

**Цель:** Зафиксировать все изменения в Git.

### 10.1 Обновить CHANGELOG.md

Добавить запись о новой версии с описанием всех изменений.

### 10.2 Обновить документацию

- `README.md` — актуализировать список скриптов
- `docs/DEVELOPER.md` — обновить при необходимости

### 10.3 Коммит

```powershell
git add -A
git commit -m "refactor: major cleanup and path centralization

- Removed work26 embedded workbook from UAZ.xlsm
- Created config.py/config.ps1 for centralized path management
- Fixed hardcoded absolute paths in all scripts
- Removed 7 unused scripts
- Removed dead code: 6 stub procedures, 1 duplicate handler, 3 unused constants
- Fixed VBA hardcoded workbook name reference
- Verified Python virtual environment configuration"
```

---

## Сводная таблица изменений

| № | Действие | Файлы | Риск |
|---|----------|-------|------|
| 1 | Создать `config.py` | 1 новый файл | Низкий |
| 2 | Создать `config.ps1` | 1 новый файл | Низкий |
| 3 | Исправить пути в `export_vba.py` | 1 файл | Средний |
| 4 | Исправить пути в `impVBA.py` | 1 файл | Средний |
| 5 | Исправить пути в `run_tests.py` | 1 файл | Средний |
| 6 | Исправить пути в `Set-ExcelTrust.ps1` | 1 файл | Низкий |
| 7 | Исправить `Mod_MainButtons.bas` (через `ThisWorkbook.Name`) | 1 файл | Низкий |
| 8 | Удалить мёртвый код из `Mod_SheetButtons.bas` | 1 файл | Низкий |
| 9 | Удалить мёртвый код из `Mod_MainButtons.bas` | 1 файл | Низкий |
| 10 | Удалить константы из `Mod_Constants.bas` | 1 файл | Низкий |
| 11 | Удалить 7 скриптов | 7 файлов | Низкий |
| 12 | Создать `remove_work26.py` и запустить | 1 новый файл | Высокий |
| 13 | Удалить `_uaz_vba_export/` | 1 директория | Низкий |
| 14 | Настроить `.venv` | — | Низкий |
| 15 | Запустить тесты | — | Критический |
| 16 | Обновить CHANGELOG и документацию | 2-3 файла | Низкий |
| 17 | Коммит | — | Низкий |

---

## Примечания для режима Code

1. **Порядок критичен:** Этапы 1-4 (конфиги + пути) должны быть выполнены до этапов 5-7 (чистка), чтобы скрипты оставались работоспособными после удаления.
2. **Тестирование после каждого этапа:** После этапов 4, 5, 7 рекомендуется запускать `python scripts/run_tests.py`.
3. **Work26 — высокий риск:** Удаление VBA-компонентов из UAZ.xlsm через COM может привести к повреждению файла. Рекомендуется сделать резервную копию UAZ.xlsm перед запуском `remove_work26.py`.
4. **Правило Z2 (.ycarules):** Перед удалением скриптов (этап 6) необходимо получить подтверждение пользователя.
5. **Правило Z5 (.ycarules):** Изменения в VBA-модулях (этапы 4, 5) должны быть утверждены пользователем перед внесением.
6. **VBA-конфиг не создаём:** В VBA используем `ThisWorkbook.Name` / `ThisWorkbook.Path` вместо отдельного `config.bas` — это безопаснее и не требует синхронизации.