# Техническая документация разработчика — SysW

## 1. Архитектура проекта

### 1.1 Общая архитектура

Система SysW построена на модульной архитектуре VBA для Excel. Каждый модуль отвечает за свою область функциональности:

- **Утилиты** ([`Mod_Utils.bas`](../src/modules/Mod_Utils.bas)) — вспомогательные функции, общие типы данных
- **Шапка заказа** ([`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas)) — заполнение заголовка заказ-наряда
- **Импорт** ([`Mod_Import.bas`](../src/modules/Mod_Import.bas)) — импорт данных из Excel в SQLite и обратно
- **Диспетчер кнопок** ([`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas)) — прослойка между UI и бизнес-логикой
- **Тестовый раннер** ([`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)) — автоматическое тестирование
- **Лист main** ([`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls)) — обработчик событий листа

### 1.2 Схема взаимодействия модулей

```
Sheet1_main.cls (Worksheet_Change)
       │
       └── Mod_OrderHeader.FillHeaderFromOrder()
              │
              ├── Mod_Utils.GetSheetByName()
              ├── Mod_Utils.FileExists()
              └── Mod_Utils.FormatDateSQL()

Mod_ButtonDispatcher (обработчики кнопок)
       │
       ├── Mod_Import.ClearMainSheet_UI()
       ├── Mod_Import.ImportSheet_UI()
       ├── Mod_Import.ClearHeader_UI()
       ├── Mod_OrderHeader.FillHeaderFromOrder_UI()
       └── ...

Mod_FullTestRunner.RunAllTests()
       │
       ├── RunUtilsTests()         → Mod_Utils
       ├── RunOrderHeaderTests()   → Mod_OrderHeader
       ├── RunImportTests()        → Mod_Import
       └── RunButtonTests()        → Mod_ButtonDispatcher
```

### 1.3 Принципы модульной архитектуры

1. **Разделение ответственности** — каждый модуль решает только свою задачу
2. **Слабая связанность** — модули общаются через вызовы функций, без общих глобальных переменных
3. **UI-обёртки** — функции с суффиксом `_UI` содержат пользовательские диалоги (MsgBox), чистые функции — нет
4. **Диспетчер кнопок** — единственный модуль, который привязан к элементам управления на формах

---

## 2. Описание модулей VBA

### 2.1 Mod_Utils.bas — Утилитарные функции

**Файл:** [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas) (156 строк)

**Назначение:** Вспомогательные функции для работы с Excel.

**Типы данных:**

| Тип | Поле | Описание |
|-----|------|----------|
| `OrderHeader` | `OrderNumber As String` | № п/п (колонка A листа spisok) |
| | `ModelName As String` | Модель (колонка B) |
| | `GRZ As String` | ГРЗ/госномер (колонка C) |
| | `VIN As String` | VIN (колонка D) |
| | `GarageNumber As String` | Гаражный № (колонка E) |
| | `YearMade As Integer` | Год выпуска (колонка F) |
| | `MileageValue As Long` | Пробег (колонка G) |
| | `DateValue As Date` | Дата (колонка H) |

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `GetSheetByName(wb, SheetName)` | Получение листа по имени без ошибки, если лист не найден |
| `GetWorkbookPath()` | Путь к текущей книге |
| `FileExists(FilePath)` | Проверка существования файла через `Dir()` |
| `GetCurrentUser()` | Имя пользователя Windows (`Environ("USERNAME")`) |
| `FormatDateSQL(d)` | Форматирование даты в формат SQLite: `ГГГГ-ММ-ДД` |

### 2.2 Mod_OrderHeader.bas — Работа с шапкой заказ-наряда

**Файл:** [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas) (203 строки)

**Назначение:** Заполнение заголовка заказа-наряда (B3:B15) на листе main данными из листов spisok и model.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `FillHeaderFromOrder(OrderNum)` | Основная функция. Ищет номер заказа в листе spisok, заполняет B3:B15 на листе main. Возвращает `Boolean` — успех/неудача |
| `FillHeaderFromOrder_UI()` | UI-обёртка: вызывает `FillHeaderFromOrder`, показывает MsgBox с результатом |

**Константы столбцов листа spisok:**

| Константа | Значение | Назначение |
|-----------|----------|------------|
| `SPISOK_COL_MODEL` | 2 | Модель |
| `SPISOK_COL_GRZ` | 3 | ГРЗ |
| `SPISOK_COL_VIN` | 4 | VIN |
| `SPISOK_COL_GARAGE` | 5 | Гаражный № |
| `SPISOK_COL_YEAR` | 6 | Год выпуска |
| `SPISOK_COL_MILEAGE` | 7 | Пробег |
| `SPISOK_COL_DATE` | 8 | Дата |
| `SPISOK_COL_NOTE` | 10 | Примечание |

**Константы столбцов листа model:**

| Константа | Значение | Назначение |
|-----------|----------|------------|
| `MODEL_COL_GROUP` | 2 | Группа модели |
| `MODEL_COL_PRICE` | 3 | Цена |
| `MODEL_COL_WORKS_ORIG` | 4 | Оригинальные работы |
| `MODEL_COL_WORKS_MOD` | 5 | Модифицированные работы |

### 2.3 Mod_Import.bas — Импорт данных

**Файл:** [`Mod_Import.bas`](../src/modules/Mod_Import.bas) (370 строк)

**Назначение:** Импорт данных из Excel в SQLite и обратно. Поиск листов в report.xlsx по ГРЗ.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `ExtractNumberFromGRZ(GRZ)` | Извлекает цифровую группу длиной 3 или 4 цифры из строки ГРЗ (пример: "А123АН77" → "123") |
| `SearchSheetByGRZ(GRZ)` | Открывает report.xlsx (если не открыт), ищет лист по номеру ГРЗ. Возвращает найденный лист или Nothing |
| `ClearMainSheet_UI()` | Очищает все данные на листе main с подтверждением пользователя |
| `ImportSheet_UI()` | Запускает импорт из отчёта по ГРЗ из ячейки B4 |
| `ClearHeader_UI()` | Очищает только шапку заказа (B3:B15) на листе main |

### 2.4 Mod_ButtonDispatcher.bas — Диспетчер кнопок

**Файл:** [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) (121 строка)

**Назначение:** Содержит **только** однострочные вызовы UI-процедур. Является прослойкой между кнопками на формах и бизнес-логикой. Не содержит бизнес-логики.

**Обработчики:**

| Процедура | Вызов |
|-----------|-------|
| `Btn_main_Clear_Click()` | `Mod_Import.ClearMainSheet_UI` |
| `Btn_main_Import_Click()` | `Mod_Import.ImportSheet_UI` |
| `Btn_main_FillHeader_Click()` | `Mod_OrderHeader.FillHeaderFromOrder_UI` |
| `Btn_main_ClearHeader_Click()` | `Mod_Import.ClearHeader_UI` |
| `Btn_main_ImportByInput_Click()` | Импорт по ГРЗ, введённому пользователем через InputBox |

### 2.5 Mod_FullTestRunner.bas — Тестовый раннер

**Файл:** [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas) (564 строки)

**Назначение:** Автоматическое тестирование VBA-модулей. Содержит 20 тестовых сценариев (TC-01..TC-20), из которых 17 активных и 3 пропущенных (SKIP).

**Группы тестов:**

| Группа | Сценарии | Тестируемый модуль |
|--------|----------|-------------------|
| `RunUtilsTests()` | TC-01..TC-04, TC-06, TC-07, TC-19, TC-20 | Mod_Utils |
| `RunOrderHeaderTests()` | TC-08..TC-11 | Mod_OrderHeader |
| `RunImportTests()` | TC-12..TC-15 | Mod_Import |
| `RunButtonTests()` | TC-16..TC-18 | Mod_ButtonDispatcher |

**Механизмы:**
- **SKIP** — тесты, зависящие от отсутствующих данных, пропускаются
- **Сохранение/восстановление** состояния листов до и после тестов
- **Статистика** — подсчёт Total, Passed, Failed, Skipped

**Запуск:**
```bash
python scripts/run_tests.py
```

### 2.6 Sheet1_main.cls — Основной лист

**Файл:** [`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls) (40 строк)

**Назначение:** Класс листа main. Обрабатывает событие `Worksheet_Change`.

**Логика обработчика:**
1. Защита от рекурсии через `Static isProcessing As Boolean`
2. Проверка, что изменение произошло в ячейке B2
3. Очистка диапазона B3:B15
4. Вызов `Mod_OrderHeader.FillHeaderFromOrder(CStr(b2Value))`

---

## 3. Правила кодировки (двухфазная система)

### 3.1 Принцип

VBA-файлы (`.bas`, `.cls`) используют двухфазную модель кодировки:

| Фаза | Где | Кодировка | Инструмент |
|------|-----|-----------|------------|
| **Диск** | VS Code, Git, code review | **UTF-8** (без BOM) | [`export_vba.py`](../scripts/export_vba.py) |
| **Excel** | VBA Editor, COM-автоматизация | **Windows-1251** (CP1251) | [`impVBA.py`](../scripts/impVBA.py) |

### 3.2 Конвертация

**Экспорт из Excel (CP1251 → UTF-8):**
```bash
python scripts/export_vba.py
```
1. Запускает Excel через COM
2. Экспортирует модули во временную директорию `_temp_export/` (CP1251)
3. Конвертирует каждый файл из CP1251 в UTF-8
4. Копирует в корень проекта
5. Удаляет `_temp_export/`

**Импорт в Excel (UTF-8 → CP1251):**
```bash
python scripts/impVBA.py
```
1. Читает файлы с автоопределением кодировки (UTF-8 with BOM → UTF-8 → CP1251)
2. Удаляет `Attribute`-строки (недопустимы в VBA при программном импорте)
3. Сохраняет во временную директорию `_temp_import/` в CP1251
4. Запускает Excel через COM, удаляет старые компоненты, импортирует новые
5. Сохраняет и закрывает Excel
6. Удаляет `_temp_import/`

### 3.3 Git-нормализация

- [`.gitattributes`](../.gitattributes) — настройки нормализации UTF-8 для Git
- [`.vscode/settings.json`](../.vscode/settings.json) — `files.encoding: "utf8"` для всех VBA-файлов

### 3.4 Важные правила

- **Запрещено** вручную менять кодировку VBA-файлов
- **Запрещено** редактировать VBA-модули напрямую в Excel без последующего экспорта
- **Запрещено** редактировать VBA-модули на диске без последующего импорта в Excel
- PowerShell-скрипты должны быть в **UTF-8 with BOM** (требование PowerShell для кириллицы)

---

## 4. Процессы импорта/экспорта

### 4.1 export_vba.py

**Файл:** [`export_vba.py`](../scripts/export_vba.py) (174 строки)

**Назначение:** Выгрузка VBA-модулей из Excel на диск с конвертацией CP1251 → UTF-8.

**Маппинг компонентов:**

| VBA-компонент | Файл на диске | Тип |
|--------------|---------------|-----|
| `Mod_Utils` | `src/modules/Mod_Utils.bas` | Стандартный модуль |
| `Mod_OrderHeader` | `src/modules/Mod_OrderHeader.bas` | Стандартный модуль |
| `Mod_Import` | `src/modules/Mod_Import.bas` | Стандартный модуль |
| `Mod_ButtonDispatcher` | `src/modules/Mod_ButtonDispatcher.bas` | Стандартный модуль |
| `Mod_FullTestRunner` | `src/modules/Mod_FullTestRunner.bas` | Стандартный модуль |
| `Sheet1` | `src/sheets/Sheet1_main.cls` | Класс листа |

**Использование:**
```bash
python scripts/export_vba.py          # Экспорт всех модулей
python scripts/export_vba.py --dry    # Просмотр без реального экспорта
```

### 4.2 import_all_vba.py

**Файл:** [`impVBA.py`](../scripts/impVBA.py) (281 строка)

**Назначение:** Загрузка VBA-модулей с диска в Excel с конвертацией UTF-8 → CP1251.

**Особенности:**
- Автоопределение кодировки входных файлов (UTF-8 with BOM → UTF-8 → CP1251 fallback)
- Удаление `Attribute`-строк перед импортом
- Удаление существующих компонентов перед импортом новых

**Использование:**
```bash
python scripts/impVBA.py      # Импорт всех модулей
```

### 4.3 scripts/Import-VbaFromExcel.ps1

**Файл:** [`scripts/Import-VbaFromExcel.ps1`](../scripts/Import-VbaFromExcel.ps1)

**Назначение:** Альтернативный PowerShell-скрипт для импорта VBA-модулей из Excel.

**Использование:**
```powershell
.\scripts\Import-VbaFromExcel.ps1              # По умолчанию
.\scripts\Import-VbaFromExcel.ps1 -DryRun      # Режим просмотра
.\scripts\Import-VbaFromExcel.ps1 -ExcelPath "C:\path\to\work.xlsm" -OutputDir "C:\output"
```

**Параметры:**
- `-ExcelPath` — путь к `.xlsm` файлу (по умолчанию: `L:\PROject\SysW\work.xlsm`)
- `-OutputDir` — директория для сохранения `.bas`/`.cls` файлов (по умолчанию: `L:\PROject\SysW\src`)
- `-DryRun` — показать список модулей без реального экспорта

---

## 5. Настройка окружения

### 5.1 VS Code

1. Открыть `L:\PROject\SysW` как корневую директорию проекта
2. Установить расширение **VBA Language Server** (рекомендуется для подсветки синтаксиса)
3. Настройки в [`.vscode/settings.json`](../.vscode/settings.json):
   - `files.encoding: "utf8"` — кодировка UTF-8 для всех файлов
   - `terminal.integrated.defaultProfile.windows: "PowerShell"` — терминал PowerShell
   - Кастомный профиль `SourceCraft` с UTF-8 и автоматической активацией `.venv`

### 5.2 Python

- **Версия:** Python 3.12+
- **Виртуальное окружение:**
  ```bash
  python -m venv .venv
  .venv\Scripts\Activate.ps1
  ```
- **Зависимости:**
  ```bash
  pip install pywin32
  ```
  `pywin32` требуется для COM-автоматизации Excel (используется в `scripts/export_vba.py` и `scripts/impVBA.py`)

### 5.3 PowerShell

- **Версия:** PowerShell 5.1+ (поставляется с Windows 10)
- **Политика выполнения:**
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
- **Кодировка скриптов:** UTF-8 with BOM (требование PowerShell для корректной обработки кириллицы)

---

## 6. CI/CD

### 6.1 GitHub Actions Workflow

**Файл:** [`.github/workflows/vba-check.yml`](../.github/workflows/vba-check.yml) (109 строк)

**Триггеры:**
- `push` в ветки `main` и `dev`
- `pull_request` в ветки `main` и `dev`

**Проверки:**

| Шаг | Что проверяет | Действие при неудаче |
|-----|--------------|---------------------|
| 1. Check VBA files exist | Наличие всех 6 VBA-файлов | Fail |
| 2. Check UTF-8 encoding | Валидная UTF-8 кодировка каждого файла | Fail |
| 3. Check VBA syntax (basic) | Отсутствие недопустимых символов (коды < 32, кроме \n\r\t) | Fail |
| 4. Check CHANGELOG updated | Наличие записи за сегодняшнюю дату | Warning (non-blocking) |
| 5. Check .gitignore consistency | Наличие файла `.gitignore` | Fail |

### 6.2 Pre-commit процедура

Перед каждым коммитом необходимо выполнить:

```bash
# 1. Проверить статус
git status

# 2. Синхронизировать VBA-модули (если работали в Excel)
python scripts/export_vba.py

# 3. Обновить CHANGELOG.md
# Добавить запись в формате Keep a Changelog

# 4. Проверить кодировку UTF-8
python -c "
with open('src/modules/Mod_Utils.bas', 'rb') as f:
    raw = f.read()
raw.decode('utf-8')
print('OK: UTF-8')
"

# 5. Проиндексировать и закоммитить
git add <файлы>
git commit -m "тип(область): описание"

# 6. Отправить на GitHub
git push
```

---

## Связанные документы

- [`README.md`](../README.md) — общее описание проекта, быстрый старт
- [`docs/sourcecraft-guide.md`](sourcecraft-guide.md) — руководство по работе с SourceCraft Code Assistant
- [`docs/git-workflow.md`](git-workflow.md) — Git-инструкции и веточная стратегия
- [`CHANGELOG.md`](../CHANGELOG.md) — история версий проекта

---

## Приложение A: Схема зависимостей модулей

```
Sheet1_main.cls
  └── Mod_OrderHeader.FillHeaderFromOrder()
        └── Mod_Utils (GetSheetByName, FileExists, FormatDateSQL)

Mod_ButtonDispatcher
  ├── Mod_Import (ClearMainSheet_UI, ImportSheet_UI, ClearHeader_UI)
  └── Mod_OrderHeader (FillHeaderFromOrder_UI)

Mod_FullTestRunner
  ├── Mod_Utils (тесты утилит)
  ├── Mod_OrderHeader (тесты OrderHeader)
  ├── Mod_Import (тесты импорта)
  └── Mod_ButtonDispatcher (тесты кнопок)
```

## Приложение B: Маппинг файлов VBA

| VBA-компонент | Файл на диске | Тип модуля |
|--------------|---------------|------------|
| `Mod_Utils` | `src/modules/Mod_Utils.bas` | Стандартный модуль |
| `Mod_OrderHeader` | `src/modules/Mod_OrderHeader.bas` | Стандартный модуль |
| `Mod_Import` | `src/modules/Mod_Import.bas` | Стандартный модуль |
| `Mod_ButtonDispatcher` | `src/modules/Mod_ButtonDispatcher.bas` | Стандартный модуль |
| `Mod_FullTestRunner` | `src/modules/Mod_FullTestRunner.bas` | Стандартный модуль |
| `Sheet1` | `src/sheets/Sheet1_main.cls` | Класс листа |

## Приложение C: Скрипты автоматизации

| Скрипт | Назначение | Кодировка |
|--------|-----------|-----------|
| [`export_vba.py`](../scripts/export_vba.py) | Выгрузка VBA из Excel на диск (CP1251 → UTF-8) | UTF-8 |
| [`impVBA.py`](../scripts/impVBA.py) | Загрузка VBA с диска в Excel (UTF-8 → CP1251) | UTF-8 |
| [`run_tests.py`](../scripts/run_tests.py) | Запуск тестов VBA | UTF-8 |
| [`scripts/Import-VbaFromExcel.ps1`](../scripts/Import-VbaFromExcel.ps1) | Альтернативный импорт VBA из Excel (PowerShell) | UTF-8 with BOM |