# План миграции: Фаза 3 — Реструктуризация каталогов проекта

## 1. Целевая структура

```
SysW/
├── src/                        # Исходный код VBA
│   ├── modules/                # .bas модули
│   │   ├── Mod_Utils.bas
│   │   ├── Mod_OrderHeader.bas
│   │   ├── Mod_Import.bas
│   │   ├── Mod_ButtonDispatcher.bas
│   │   ├── Mod_FullTestRunner.bas
│   │   ├── Mod_MainButtons.bas
│   │   ├── Mod_SheetButtons.bas
│   │   ├── Mod_Logger.bas
│   │   ├── Mod_Constants.bas
│   │   └── Mod_SheetOps.bas
│   └── sheets/                 # .cls листы
│       ├── Sheet1_main.cls
│       ├── Sheet_work.cls
│       └── Sheet_z4.cls
│
├── scripts/                    # Скрипты автоматизации
│   ├── export_vba.py
│   ├── impVBA.py
│   ├── run_tests.py
│   └── Import-VbaFromExcel.ps1
│
├── docs/                       # Документация
│   ├── DEVELOPER.md
│   ├── git-workflow.md
│   ├── MODERNIZATION_CHECKLIST.md
│   └── sourcecraft-guide.md
│
├── plans/                      # Планы
│   └── architecture_modernization_proposal.md
│
├── _backup/                    # Бэкапы
│
├── .github/workflows/
│   └── vba-check.yml
├── .vscode/
│   └── settings.json
├── .ycarules
├── .gitattributes
├── CHANGELOG.md
├── README.md
├── work.xlsm                   (в .gitignore)
├── report.xlsx                 (в .gitignore)
└── base.xlsx                   (в .gitignore)

## 2. Маппинг перемещения файлов

| # | Файл | Откуда | Куда | Тип операции |
|---|------|--------|------|-------------|
| 1 | Mod_Utils.bas | `./Mod_Utils.bas` | `./src/modules/Mod_Utils.bas` | `git mv` |
| 2 | Mod_OrderHeader.bas | `./Mod_OrderHeader.bas` | `./src/modules/Mod_OrderHeader.bas` | `git mv` |
| 3 | Mod_Import.bas | `./Mod_Import.bas` | `./src/modules/Mod_Import.bas` | `git mv` |
| 4 | Mod_ButtonDispatcher.bas | `./Mod_ButtonDispatcher.bas` | `./src/modules/Mod_ButtonDispatcher.bas` | `git mv` |
| 5 | Mod_FullTestRunner.bas | `./Mod_FullTestRunner.bas` | `./src/modules/Mod_FullTestRunner.bas` | `git mv` |
| 6 | Mod_MainButtons.bas | `./Mod_MainButtons.bas` | `./src/modules/Mod_MainButtons.bas` | `git mv` |
| 7 | Mod_SheetButtons.bas | `./Mod_SheetButtons.bas` | `./src/modules/Mod_SheetButtons.bas` | `git mv` |
| 8 | Mod_Logger.bas | `./Mod_Logger.bas` | `./src/modules/Mod_Logger.bas` | `git mv` |
| 9 | Mod_Constants.bas | `./Mod_Constants.bas` | `./src/modules/Mod_Constants.bas` | `git mv` |
| 10 | Mod_SheetOps.bas | `./Mod_SheetOps.bas` | `./src/modules/Mod_SheetOps.bas` | `git mv` |
| 11 | Sheet1_main.cls | `./Sheet1_main.cls` | `./src/sheets/Sheet1_main.cls` | `git mv` |
| 12 | Sheet_work.cls | `./Sheet_work.cls` | `./src/sheets/Sheet_work.cls` | `git mv` |
| 13 | Sheet_z4.cls | `./Sheet_z4.cls` | `./src/sheets/Sheet_z4.cls` | `git mv` |
| 14 | export_vba.py | `./export_vba.py` | `./scripts/export_vba.py` | `git mv` |
| 15 | impVBA.py | `./impVBA.py` | `./scripts/impVBA.py` | `git mv` |
| 16 | run_tests.py | `./run_tests.py` | `./scripts/run_tests.py` | `git mv` |

## 3. Изменения в файлах (детально)

### 3.1 `scripts/impVBA.py` (бывший `impVBA.py`)

**Строка 26:** `MODULES_PATH`
```python
# Было:
MODULES_PATH = Path(r"L:\PROject\SysW")
# Стало:
MODULES_PATH = Path(r"L:\PROject\SysW\src")
```

**Строка 27:** `TEMP_DIR` (оставить без изменений — временная директория в корне)
```python
TEMP_DIR = Path(r"L:\PROject\SysW\_temp_import")
# Без изменений
```

**Строки 38-41:** `FILES` — динамическое обнаружение файлов
```python
# Было:
FILES = sorted([
    f.name for f in MODULES_PATH.iterdir()
    if f.suffix.lower() in (".bas", ".cls") and f.is_file()
])
# Стало: ищем .bas в src/modules/, .cls в src/sheets/
FILES = sorted([
    f.name for f in (MODULES_PATH / "modules").iterdir()
    if f.suffix.lower() == ".bas" and f.is_file()
] + [
    f.name for f in (MODULES_PATH / "sheets").iterdir()
    if f.suffix.lower() == ".cls" and f.is_file()
])
```

**Строки 49-52:** `SHEET_COMPONENTS` — динамическое обнаружение
```python
# Было:
SHEET_COMPONENTS = {
    f.stem for f in MODULES_PATH.iterdir()
    if f.suffix.lower() == ".cls" and f.stem.lower().startswith("sheet") and f.is_file()
}
# Стало:
SHEET_COMPONENTS = {
    f.stem for f in (MODULES_PATH / "sheets").iterdir()
    if f.suffix.lower() == ".cls" and f.stem.lower().startswith("sheet") and f.is_file()
}
```

**Строка 216:** Путь к файлу при импорте
```python
# Было:
file_path = MODULES_PATH / file_name
# Стало: определяем поддиректорию по расширению
if file_name.lower().endswith('.cls'):
    file_path = MODULES_PATH / "sheets" / file_name
else:
    file_path = MODULES_PATH / "modules" / file_name
```

### 3.2 `scripts/export_vba.py` (бывший `export_vba.py`)

**Строка 24:** `PROJECT_DIR`
```python
# Было:
PROJECT_DIR = Path(r"L:\PROject\SysW")
# Стало:
PROJECT_DIR = Path(r"L:\PROject\SysW\src")
```

**Строки 29-40:** `COMPONENTS` — обновить пути с учётом поддиректорий
```python
# Было:
COMPONENTS = {
    "Mod_Utils": "Mod_Utils.bas",
    "Mod_OrderHeader": "Mod_OrderHeader.bas",
    "Mod_Import": "Mod_Import.bas",
    "Mod_MainButtons": "Mod_MainButtons.bas",
    "Mod_SheetButtons": "Mod_SheetButtons.bas",
    "Mod_ButtonDispatcher": "Mod_ButtonDispatcher.bas",
    "Mod_FullTestRunner": "Mod_FullTestRunner.bas",
    "Sheet1_main": "Sheet1_main.cls",
    "Sheet_work": "Sheet_work.cls",
    "Sheet_z4": "Sheet_z4.cls",
}
# Стало:
COMPONENTS = {
    "Mod_Utils": "modules/Mod_Utils.bas",
    "Mod_OrderHeader": "modules/Mod_OrderHeader.bas",
    "Mod_Import": "modules/Mod_Import.bas",
    "Mod_MainButtons": "modules/Mod_MainButtons.bas",
    "Mod_SheetButtons": "modules/Mod_SheetButtons.bas",
    "Mod_ButtonDispatcher": "modules/Mod_ButtonDispatcher.bas",
    "Mod_FullTestRunner": "modules/Mod_FullTestRunner.bas",
    "Sheet1_main": "sheets/Sheet1_main.cls",
    "Sheet_work": "sheets/Sheet_work.cls",
    "Sheet_z4": "sheets/Sheet_z4.cls",
}
```

**Строка 87:** `output_path` — автоматически использует `COMPONENTS` с путями, изменений не требуется, т.к. `PROJECT_DIR / COMPONENTS[comp_name]` даст `src/modules/...` и `src/sheets/...`.

### 3.3 `.github/workflows/vba-check.yml`

**Строки 29, 41:** Список файлов для проверки существования и UTF-8
```yaml
# Было:
for f in Mod_Utils.bas Mod_OrderHeader.bas Mod_Import.bas Mod_ButtonDispatcher.bas Mod_FullTestRunner.bas Sheet1_main.cls; do
# Стало:
for f in src/modules/Mod_Utils.bas src/modules/Mod_OrderHeader.bas src/modules/Mod_Import.bas src/modules/Mod_ButtonDispatcher.bas src/modules/Mod_FullTestRunner.bas src/sheets/Sheet1_main.cls; do
```

**Строка 62:** Глобальный поиск `*.bas` и `*.cls`
```yaml
# Было:
for f in *.bas *.cls; do
# Стало:
for f in src/modules/*.bas src/sheets/*.cls; do
```

### 3.4 `scripts/Import-VbaFromExcel.ps1`

**Строка 24:** `OutputDir`
```powershell
# Было:
[string]$OutputDir = "L:\PROject\SysW",
# Стало:
[string]$OutputDir = "L:\PROject\SysW\src",
```

**Строки 29-36:** `$componentMap` — обновить пути
```powershell
# Было:
$componentMap = @{
    "Mod_Utils"           = "Mod_Utils.bas"
    "Mod_OrderHeader"     = "Mod_OrderHeader.bas"
    "Mod_Import"          = "Mod_Import.bas"
    "Mod_ButtonDispatcher" = "Mod_ButtonDispatcher.bas"
    "Mod_FullTestRunner"  = "Mod_FullTestRunner.bas"
    "Sheet1"              = "Sheet1_main.cls"
}
# Стало:
$componentMap = @{
    "Mod_Utils"           = "modules/Mod_Utils.bas"
    "Mod_OrderHeader"     = "modules/Mod_OrderHeader.bas"
    "Mod_Import"          = "modules/Mod_Import.bas"
    "Mod_ButtonDispatcher" = "modules/Mod_ButtonDispatcher.bas"
    "Mod_FullTestRunner"  = "modules/Mod_FullTestRunner.bas"
    "Sheet1"              = "sheets/Sheet1_main.cls"
}
```

### 3.5 `docs/DEVELOPER.md`

Обновить все относительные ссылки на VBA-файлы:
- `../Mod_Utils.bas` → `../src/modules/Mod_Utils.bas`
- `../Mod_OrderHeader.bas` → `../src/modules/Mod_OrderHeader.bas`
- `../Mod_Import.bas` → `../src/modules/Mod_Import.bas`
- `../Mod_ButtonDispatcher.bas` → `../src/modules/Mod_ButtonDispatcher.bas`
- `../Mod_FullTestRunner.bas` → `../src/modules/Mod_FullTestRunner.bas`
- `../Sheet1_main.cls` → `../src/sheets/Sheet1_main.cls`
- `../export_vba.py` → `../scripts/export_vba.py`
- `../import_all_vba.py` → `../scripts/impVBA.py`
- `../run_tests.py` → `../scripts/run_tests.py`

Обновить секцию "Маппинг файлов VBA" (Приложение B) — добавить колонку "Путь".

Обновить pre-commit процедуру (раздел 6.2) — путь к файлу для проверки кодировки:
```bash
# Было:
python -c "with open('Mod_Utils.bas', 'rb') as f: ..."
# Стало:
python -c "with open('src/modules/Mod_Utils.bas', 'rb') as f: ..."
```

### 3.6 `docs/sourcecraft-guide.md`

Обновить схему структуры проекта (строки 83-107):
- `Mod_*.bas` → `src/modules/Mod_*.bas`
- `Sheet1_main.cls` → `src/sheets/Sheet1_main.cls`
- `export_vba.py` → `scripts/export_vba.py`
- `import_all_vba.py` → `scripts/impVBA.py`
- `run_tests.py` → `scripts/run_tests.py`

Обновить все относительные ссылки на скрипты.

### 3.7 `.ycarules`

Обновить правило `[S1]` (строки 99-104):
```
[S1] **Организация папок**
- `src/modules/` — VBA-модули (.bas)
- `src/sheets/` — VBA-классы листов (.cls)
- `scripts/` — скрипты автоматизации (Python, PowerShell)
- `plans/` — планы изменений, архитектурные решения, отчёты
- `docs/` — документация проекта
- `.vscode/` — настройки VS Code
- Корень проекта — конфигурационные файлы (.gitattributes, .ycarules, README.md, CHANGELOG.md)
```

### 3.8 `CHANGELOG.md`

Добавить запись о новой версии с реструктуризацией.

## 4. Порядок выполнения (пошагово)

### Шаг 1: Создать новую структуру каталогов
```bash
mkdir -p src/modules src/sheets
```

### Шаг 2: Переместить VBA-файлы (git mv)
```bash
git mv Mod_Utils.bas src/modules/
git mv Mod_OrderHeader.bas src/modules/
git mv Mod_Import.bas src/modules/
git mv Mod_ButtonDispatcher.bas src/modules/
git mv Mod_FullTestRunner.bas src/modules/
git mv Mod_MainButtons.bas src/modules/
git mv Mod_SheetButtons.bas src/modules/
git mv Mod_Logger.bas src/modules/
git mv Mod_Constants.bas src/modules/
git mv Mod_SheetOps.bas src/modules/
git mv Sheet1_main.cls src/sheets/
git mv Sheet_work.cls src/sheets/
git mv Sheet_z4.cls src/sheets/
```

### Шаг 3: Переместить Python-скрипты (git mv)
```bash
git mv export_vba.py scripts/
git mv impVBA.py scripts/
git mv run_tests.py scripts/
```

### Шаг 4: Обновить `scripts/impVBA.py`
- Изменить `MODULES_PATH` на `src`
- Изменить логику поиска `.bas` в `modules/`, `.cls` в `sheets/`
- Изменить `SHEET_COMPONENTS` на поиск в `sheets/`
- Изменить `file_path` для импорта с учётом поддиректорий

### Шаг 5: Обновить `scripts/export_vba.py`
- Изменить `PROJECT_DIR` на `src`
- Обновить `COMPONENTS` с путями `modules/` и `sheets/`

### Шаг 6: Обновить `.github/workflows/vba-check.yml`
- Обновить пути к файлам во всех `for`-циклах

### Шаг 7: Обновить `scripts/Import-VbaFromExcel.ps1`
- Изменить `$OutputDir` на `src`
- Обновить `$componentMap` с путями `modules/` и `sheets/`

### Шаг 8: Обновить документацию
- `docs/DEVELOPER.md` — все ссылки
- `docs/sourcecraft-guide.md` — схема и ссылки

### Шаг 9: Обновить `.ycarules`
- Правило `[S1]` — новая структура

### Шаг 10: Проверить импорт
```bash
cd L:\PROject\SysW
python scripts/impVBA.py
```

### Шаг 11: Обновить CHANGELOG.md
- Добавить запись о версии 0.6.0

### Шаг 12: Финальная проверка
```bash
git status
# Проверить, что все файлы на месте
ls src/modules/*.bas src/sheets/*.cls scripts/*.py
```

## 5. Риски и митигация

| Риск | Вероятность | Влияние | Митигация |
|------|------------|---------|-----------|
| `git mv` не сохранит историю | Низкая | Среднее | `git mv` сохраняет историю, т.к. Git отслеживает переименования по content hash |
| `impVBA.py` не найдёт файлы после перемещения | Средняя | Высокое | Обновить `MODULES_PATH` и логику поиска; проверить после каждого изменения |
| `export_vba.py` запишет файлы не туда | Средняя | Высокое | Обновить `PROJECT_DIR` и `COMPONENTS`; проверить через `--dry` |
| PowerShell-скрипт не сработает | Низкая | Среднее | Обновить `$OutputDir` и `$componentMap` |
| CI/CD упадёт после пуша | Высокая | Среднее | Обновить пути в `vba-check.yml` до пуша |
| Документация содержит устаревшие ссылки | Средняя | Низкое | Обновить все ссылки в `DEVELOPER.md` и `sourcecraft-guide.md` |
| `.ycarules` правило `[S1]` устарело | Низкая | Низкое | Обновить описание структуры |

## 6. Критерии успешного завершения

- [ ] `git status` показывает только перемещённые файлы (rename), без delete+add
- [ ] `python scripts/impVBA.py` успешно импортирует все 13 VBA-компонентов
- [ ] `python scripts/export_vba.py --dry` показывает корректные пути
- [ ] Все ссылки в документации ведут на существующие файлы
- [ ] `.ycarules` описывает актуальную структуру
- [ ] CI/CD workflow проверяет файлы по новым путям