# Техническая документация разработчика — SysW

## 1. Архитектура проекта

### 1.1 Общая архитектура

Система SysW построена на модульной архитектуре VBA для Excel. Каждый модуль отвечает за свою область функциональности:

- **Утилиты** ([`Mod_Utils.bas`](../src/modules/Mod_Utils.bas)) — вспомогательные функции, общие типы данных
- **Шапка заказа** ([`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas)) — заполнение заголовка заказ-наряда
- **Импорт** ([`Mod_Import.bas`](../src/modules/Mod_Import.bas)) — импорт данных из отчётов Excel
- **Диспетчер кнопок** ([`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas)) — прослойка между UI и бизнес-логикой
- **Тестовый раннер** ([`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)) — автоматическое тестирование
- **Логирование** ([`Mod_Logger.bas`](../src/modules/Mod_Logger.bas)) — логирование с ротацией
- **Операции с листами** ([`Mod_SheetOps.bas`](../src/modules/Mod_SheetOps.bas)) — операции с листами
- **Кнопки листа main** ([`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas)) — кнопки листа main
- **Кнопки листов z4/work** ([`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas)) — кнопки листов z4/work
- **Константы и реестр имён** ([`Mod_Constants.bas`](../src/modules/Mod_Constants.bas)) — константы столбцов и управление листом libname
- **Лист main** ([`Лист2_main.cls`](../src/sheets/Лист2_main.cls)) — обработчик событий листа
- **Лист work** ([`Sheet_work.cls`](../src/sheets/Sheet_work.cls)) — обработчик событий листа work
- **Лист z4** ([`Sheet_z4.cls`](../src/sheets/Sheet_z4.cls)) — обработчик событий листа z4

### 1.2 Схема взаимодействия модулей

```
Лист2_main.cls (Worksheet_Change)
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

Mod_MainButtons (кнопки листа main)
       │
       ├── Mod_Import (вызовы импорта)
       ├── Mod_OrderHeader (заполнение шапки)
       └── Mod_SheetOps (операции с листами)

Mod_SheetButtons (кнопки листов z4/work)
       │
       ├── Mod_Import (вызовы импорта)
       └── Mod_SheetOps (операции с листами)

Mod_FullTestRunner.RunAllTests()
       │
       ├── RunUtilsTests()         → Mod_Utils
       ├── RunLoggerTests()        → Mod_Logger
       ├── RunUtilsEdgeTests()     → Mod_Utils
       └── RunLibNameTests()       → Mod_Constants
```

### 1.3 Принципы модульной архитектуры

1. **Разделение ответственности** — каждый модуль решает только свою задачу
2. **Слабая связанность** — модули общаются через вызовы функций, без общих глобальных переменных
3. **UI-обёртки** — функции с суффиксом `_UI` содержат пользовательские диалоги (MsgBox), чистые функции — нет
4. **Диспетчер кнопок** — единственный модуль, который привязан к элементам управления на формах

---

## 2. Описание модулей VBA

### 2.1 Mod_Utils.bas — Утилитарные функции

**Файл:** [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas) (142 строки)

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
| `WriteLog(message)` | Обёртка для `Mod_Logger.WriteLog` (обратная совместимость) |

### 2.2 Mod_OrderHeader.bas — Работа с шапкой заказ-наряда

**Файл:** [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas) (265 строк)

**Назначение:** Заполнение заголовка заказа-наряда (B3:B15) на листе main данными из листов spisok и models.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `FillHeaderFromOrder(OrderNum)` | Основная функция. Ищет номер заказа в листе spisok, заполняет B3:B15 на листе main. Возвращает `Boolean` — успех/неудача |
| `FillHeaderFromOrder_UI()` | UI-обёртка: вызывает `FillHeaderFromOrder`, показывает MsgBox с результатом |
| `FindOrder(orderNum, Header)` | Поиск заказа по номеру, заполняет структуру `OrderHeader` |
| `FindOrder_UI()` | UI-обёртка: запрашивает номер через InputBox, показывает результат |

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

**Константы столбцов листа models:**

| Константа | Значение | Назначение |
|-----------|----------|------------|
| `MODELS_COL_MODEL` | 1 | Модель |
| `MODELS_COL_GROUP` | 2 | Группа модели |
| `MODELS_COL_PRICE` | 3 | Цена н/ч |

### 2.3 Mod_Import.bas — Импорт данных

**Файл:** [`Mod_Import.bas`](../src/modules/Mod_Import.bas) (229 строк)

**Назначение:** Импорт данных из отчётов Excel. Поиск листов в report.xlsx по ГРЗ.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `ImportSheet(grz)` | Импортирует лист из report.xlsx по ГРЗ в текущую книгу |
| `ImportDataToMain(wsSource)` | Переносит данные из листа-источника в лист main по столбцам |
| `ImportSheet_UI()` | Запускает импорт из отчёта по ГРЗ из ячейки B4 |
| `ImportByInput_UI()` | Запрашивает ГРЗ через InputBox, вызывает ImportSheet |
| `RenameSheets_UI()` | Переименовывает листы в report.xlsx по ГРЗ |
| `ImportDataToMain_UI()` | Переносит данные с активного листа в лист main |

### 2.4 Mod_ButtonDispatcher.bas — Диспетчер кнопок

**Файл:** [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) (121 строка)

**Назначение:** Содержит **только** однострочные вызовы UI-процедур. Является прослойкой между кнопками на формах и бизнес-логикой. Не содержит бизнес-логики.

**Обработчики:**

| Процедура | Вызов |
|-----------|-------|
| `Btn_main_Clear_Click()` | `Mod_SheetOps.ClearMainSheet_UI` |
| `Btn_main_Import_Click()` | `Mod_Import.ImportSheet_UI` |
| `Btn_main_FillHeader_Click()` | `Mod_OrderHeader.FillHeaderFromOrder_UI` |
| `Btn_main_ClearHeader_Click()` | `Mod_SheetOps.ClearHeader_UI` |
| `Btn_main_ImportByInput_Click()` | Импорт по ГРЗ, введённому пользователем через InputBox |
| `Btn_main_RunTests_Click()` | `Mod_FullTestRunner.RunAllTests_UI` |
| `Btn_main_WriteLog_Click()` | `Mod_Utils.WriteLog_UI` |
| `Btn_main_RenameSheets_Click()` | `Mod_Import.RenameSheets_UI` |
| `Btn_main_ImportDataToMain_Click()` | `Mod_Import.ImportDataToMain_UI` |
| `Btn_main_FindOrder_Click()` | `Mod_OrderHeader.FindOrder_UI` |
| `Btn_main_ShowWorkbookPath_Click()` | `Mod_Utils.ShowWorkbookPath_UI` |
| `Btn_main_ShowCurrentUser_Click()` | `Mod_Utils.ShowCurrentUser_UI` |
| `Btn_main_CheckFileExists_Click()` | `Mod_Utils.CheckFileExists_UI` |

### 2.5 Mod_FullTestRunner.bas — Тестовый раннер

**Файл:** [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas) (523 строки)

**Назначение:** Автоматическое тестирование VBA-модулей. Содержит 13 тестовых сценариев (TC-01..TC-13).

**Группы тестов:**

| Группа | Сценарии | Тестируемый модуль |
|--------|----------|-------------------|
| `RunUtilsTests()` | TC-01..TC-08 | Mod_Utils |
| `RunLoggerTests()` | TC-09..TC-11 | Mod_Logger |
| `RunUtilsEdgeTests()` | TC-12 | Mod_Utils (граничные случаи) |
| `RunLibNameTests()` | TC-13 | Mod_Constants |

**Механизмы:**
- **SKIP** — тесты, зависящие от отсутствующих данных, пропускаются
- **Сохранение/восстановление** состояния листов до и после тестов
- **Статистика** — подсчёт Total, Passed, Failed, Skipped
- **GetTestResults()** — функция для программного сбора результатов из Python
- **Silent mode** — параметр `silent` в `ClearMainSheet_UI` для автоматических тестов без MsgBox

**Запуск:**
```bash
python scripts/run_tests.py
```

### 2.6 Mod_Logger.bas — Логирование с ротацией

**Файл:** [`Mod_Logger.bas`](../src/modules/Mod_Logger.bas)

**Назначение:** Централизованное логирование с поддержкой ротации лог-файлов.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `WriteLog(ModuleName, Message)` | Запись сообщения в лог |
| `WriteLogE(ModuleName, Message)` | Запись сообщения об ошибке с префиксом [ERROR] |
| `RotateLogIfNeeded(MaxSizeKB)` | Ротация лог-файла при превышении указанного размера |
| `ClearLog()` | Очистка файла лога |
| `GetLogPath()` | Возвращает путь к файлу лога |

### 2.7 Mod_Constants.bas — Константы столбцов

**Файл:** [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas)

**Назначение:** Централизованное хранение констант столбцов для всех листов.

**Константы листа spisok:**

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

**Константы листа models:**

| Константа | Значение | Назначение |
|-----------|----------|------------|
| `MODELS_COL_MODEL` | 1 | Модель |
| `MODELS_COL_GROUP` | 2 | Группа модели |
| `MODELS_COL_PRICE` | 3 | Цена н/ч |

### 2.8 Mod_SheetOps.bas — Операции с листами

**Файл:** [`Mod_SheetOps.bas`](../src/modules/Mod_SheetOps.bas)

**Назначение:** Операции по управлению листами: поиск, переименование, очистка, работа с ГРЗ.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `ExtractNumberFromGRZ(GRZ)` | Извлекает цифровую группу длиной 3 или 4 цифры из строки ГРЗ |
| `SearchSheetByGRZ(GRZ)` | Открывает report.xlsx, ищет лист по номеру ГРЗ |
| `RenameSheetsByGRZ()` | Переименовывает листы в report.xlsx по номеру ГРЗ |
| `ClearMainSheet_UI([silent])` | Очищает все данные на листе main с подтверждением |
| `ClearHeader_UI()` | Очищает шапку заказа (B3:B15) на листе main |

### 2.9 Mod_MainButtons.bas — Кнопки листа main

**Файл:** [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas)

**Назначение:** Обработчики кнопок, расположенных на листе main. Содержит бизнес-логику, специфичную для листа main.

**Обработчики:**

| Процедура | Описание |
|-----------|----------|
| `Btn_main_Import()` | Импорт данных на лист main |
| `Btn_main_AUTOz4()` | Заглушка: автоподбор запчастей — в разработке |
| `Btn_main_AUTOw()` | Заглушка: автоподбор работ — в разработке |
| `Btn_main_MANz4()` | Заглушка: ручной подбор запчастей — в разработке |
| `Btn_main_MANw()` | Заглушка: ручной подбор работ — в разработке |

### 2.10 Mod_SheetButtons.bas — Кнопки листов z4/work

**Файл:** [`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas)

**Назначение:** Обработчики кнопок, расположенных на листах z4 и work. Содержит бизнес-логику, специфичную для этих листов.

**Обработчики:**

| Процедура | Описание |
|----------|----------|
| `Btn_z4_Action1()` | Заглушка: действие 1 для запчастей — в разработке |
| `Btn_z4_Action2()` | Заглушка: действие 2 для запчастей — в разработке |
| `Btn_z4_Action3()` | Заглушка: действие 3 для запчастей — в разработке |
| `Btn_work_Action1()` | Заглушка: действие 1 для работ — в разработке |
| `Btn_work_Action2()` | Заглушка: действие 2 для работ — в разработке |
| `Btn_work_Action3()` | Заглушка: действие 3 для работ — в разработке |

### 2.11 Mod_Constants.bas — Константы и реестр имён

**Файл:** [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas)

**Назначение:** Централизованное хранение констант проекта и управление листом libname (реестр имён). Объединяет числовые константы столбцов и строковые имена для реестра.

**Числовые константы (столбцы листов):**

| Константа | Значение | Описание |
|-----------|----------|----------|
| `SPISOK_COL_NUM` | 1 | № п/п (столбец A листа spisok) |
| `SPISOK_COL_MODEL` | 2 | Модель (столбец B листа spisok) |
| `SPISOK_COL_GRZ` | 3 | ГРЗ (столбец C листа spisok) |
| `SPISOK_COL_VIN` | 4 | VIN (столбец D листа spisok) |
| `SPISOK_COL_GARAGE` | 5 | Гараж. № (столбец E листа spisok) |
| `SPISOK_COL_YEAR` | 6 | Год выпуска (столбец F листа spisok) |
| `SPISOK_COL_MILEAGE` | 7 | Пробег (столбец G листа spisok) |
| `SPISOK_COL_DATE` | 8 | Дата (столбец H листа spisok) |
| `SPISOK_COL_GROUP` | 9 | Группа (столбец I листа spisok) |
| `SPISOK_COL_NOTE` | 10 | РЕЗЕРВ (столбец J листа spisok) |
| `MODELS_COL_MODEL` | 1 | Модель (столбец A листа models) |
| `MODELS_COL_GROUP` | 2 | Группа (столбец B листа models) |
| `MODELS_COL_PRICE` | 3 | Цена н/ч (столбец C листа models) |

**Строковые константы (для листа libname):**

| Константа | Значение |
|-----------|----------|
| `SPISOK_COL_NUM_NAME` | "spisok" |
| `SPISOK_COL_MODEL_NAME` | "model" |
| `SPISOK_COL_GRZ_NAME` | "grz" |
| `SPISOK_COL_VIN_NAME` | "vin" |
| `SPISOK_COL_GARAGE_NAME` | "garnum" |
| `SPISOK_COL_YEAR_NAME` | "year" |
| `SPISOK_COL_MILEAGE_NAME` | "mileage" |
| `SPISOK_COL_DATE_NAME` | "date" |
| `SPISOK_COL_GROUP_NAME` | "group" |
| `SPISOK_COL_NOTE_NAME` | "reserve" |
| `MODELS_COL_MODEL_NAME` | "model_name" |
| `MODELS_COL_GROUP_NAME` | "group" |
| `MODELS_COL_PRICE_NAME` | "hrpr" |
| `WORK_NAME` | "work" |
| `Z4_NAME` | "z4" |

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `InitLibName()` | Заполняет лист libname начальными данными реестра имён |
| `AddWorkEntry()` | Добавляет запись для work.xlsm в конец списка на листе libname |

> **Примечание:** Ранее функциональность реестра имён находилась в отдельном модуле `Mod_LibName.bas`, который был объединён с `Mod_Constants.bas` для централизованного управления константами.

### 2.12 Лист2_main.cls — Основной лист

**Файл:** [`Лист2_main.cls`](../src/sheets/Лист2_main.cls) (48 строк)

**Назначение:** Класс листа main. Обрабатывает событие `Worksheet_Change`.

**Логика обработчика:**
1. Защита от рекурсии через `Static isProcessing As Boolean`
2. Проверка, что изменение произошло в ячейке B2
3. Очистка диапазона B3:B15
4. Вызов `Mod_OrderHeader.FillHeaderFromOrder(CStr(b2Value))`

### 2.13 Sheet_work.cls — Лист work

**Файл:** [`Sheet_work.cls`](../src/sheets/Sheet_work.cls)

**Назначение:** Класс листа work. Обрабатывает события, специфичные для листа work.

**Обработчики:**
- `Worksheet_Activate` — закрепление первых двух строк при активации листа

### 2.14 Sheet_z4.cls — Лист z4

**Файл:** [`Sheet_z4.cls`](../src/sheets/Sheet_z4.cls)

**Назначение:** Класс листа z4. Обрабатывает события, специфичные для листа z4.

**Обработчики:**
- `Worksheet_Activate` — закрепление первых двух строк при активации листа

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
4. Копирует в `src/`
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
| `Mod_Logger` | `src/modules/Mod_Logger.bas` | Стандартный модуль |
| `Mod_Constants` | `src/modules/Mod_Constants.bas` | Стандартный модуль |
| `Mod_SheetOps` | `src/modules/Mod_SheetOps.bas` | Стандартный модуль |
| `Mod_MainButtons` | `src/modules/Mod_MainButtons.bas` | Стандартный модуль |
| `Mod_SheetButtons` | `src/modules/Mod_SheetButtons.bas` | Стандартный модуль |
| `Лист2` | `src/sheets/Лист2_main.cls` | Класс листа |
| `Sheet_work` | `src/sheets/Sheet_work.cls` | Класс листа |
| `Sheet_z4` | `src/sheets/Sheet_z4.cls` | Класс листа |

**Использование:**
```bash
python scripts/export_vba.py          # Экспорт всех модулей
python scripts/export_vba.py --dry    # Просмотр без реального экспорта
```

### 4.2 impVBA.py

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

## 6. Тестирование

### 6.1 Полный список тестов (TC-01..TC-13)

| ID | Название | Модуль | Тип | Статус |
|----|----------|--------|-----|--------|
| TC-01 | FileExists с существующим файлом | Mod_Utils | Модульный | ✅ |
| TC-02 | FileExists с несуществующим файлом | Mod_Utils | Модульный | ✅ |
| TC-03 | FormatDateSQL с корректной датой | Mod_Utils | Модульный | ✅ |
| TC-04 | FormatDateSQL с нулевой датой | Mod_Utils | Модульный | ✅ |
| TC-05 | GetSheetByName существующий лист | Mod_Utils | Модульный | ✅ |
| TC-06 | GetSheetByName несуществующий лист | Mod_Utils | Модульный | ✅ |
| TC-07 | WriteLog запись в лог | Mod_Utils | Модульный | ✅ |
| TC-08 | GetWorkbookPath / GetCurrentUser | Mod_Utils | Модульный | ✅ |
| TC-09 | WriteLog запись в лог-файл | Mod_Logger | Модульный | ✅ |
| TC-10 | RotateLogIfNeeded ротация лога | Mod_Logger | Модульный | ✅ |
| TC-11 | ClearLog очистка лога | Mod_Logger | Модульный | ✅ |
| TC-12 | FormatDateSQL граничные случаи | Mod_Utils | Модульный | ✅ |
| TC-13 | InitLibName заполнение libname | Mod_Constants | Модульный | ✅ |

**Легенда:**
- ✅ — тест проходит (PASS)
- ⚠️ — тест пропущен (SKIP) из-за отсутствия данных/окружения
- ❌ — тест падает (FAIL)

### 6.2 Таблица покрытия модулей

| Модуль | Всего тестов | PASS | SKIP | FAIL | Покрытие |
|--------|-------------|------|------|------|----------|
| Mod_Utils | 8 | 8 | 0 | 0 | 100% |
| Mod_Logger | 3 | 3 | 0 | 0 | 100% |
| Mod_Constants | 1 | 1 | 0 | 0 | 100% |
| **Итого** | **12** | **12** | **0** | **0** | **100%** |

### 6.3 Как добавить новый тест

1. Открой [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)
2. Выбери подходящую группу тестов или создай новую процедуру-группу
3. Добавь вызов новой группы в `RunAllTests()`
4. Используй шаблон:

```vba
' -------------------------------------------------------
' TC-XX: Название теста
' -------------------------------------------------------
On Error Resume Next
' ... код теста ...
If Err.number <> 0 Then
    AddResult "TC-XX", "Название теста", False, "Ошибка: " & Err.Description
    Err.Clear
Else
    AddResult "TC-XX", "Название теста", (условие), "пояснение при FAIL"
End If
On Error GoTo 0
```

5. Для SKIP используй: `AddResult "TC-XX", "...", True, "", True, "Причина пропуска"`
6. Запусти тесты: `python scripts/run_tests.py`

### 6.4 Запуск тестов в CI/CD

Тесты запускаются через Python-скрипт [`scripts/run_tests.py`](../scripts/run_tests.py):

```bash
python scripts/run_tests.py
```

**Что делает скрипт:**
1. Создаёт COM-объект Excel (невидимый режим)
2. Открывает `work.xlsm`
3. Запускает макрос `RunAllTests`
4. Собирает результаты через VBA-функцию `GetTestResults()`
5. Парсит статистику (Total, Passed, Failed, Skipped)
6. Возвращает exit code: `0` — все PASS, `1` — есть FAIL
7. Сохраняет результаты в `test_results.log`
8. Гарантированно закрывает Excel (в `finally` блоке)

**Интеграция с GitHub Actions:**

```yaml
- name: Run VBA tests
  run: python scripts/run_tests.py
  shell: powershell
```

> **Важно:** GitHub Actions не поддерживает COM-автоматизацию Excel (требуется Windows с установленным Excel). Для CI/CD на Linux/MacOS тесты будут пропущены.

---

## 7. CI/CD

### 7.1 GitHub Actions Workflow

**Файл:** [`.github/workflows/vba-check.yml`](../.github/workflows/vba-check.yml) (109 строк)

**Триггеры:**
- `push` в ветки `main` и `dev`
- `pull_request` в ветки `main` и `dev`

**Проверки:**

| Шаг | Что проверяет | Действие при неудаче |
|-----|--------------|---------------------|
| 1. Check VBA files exist | Наличие всех 13 VBA-файлов | Fail |
| 2. Check UTF-8 encoding | Валидная UTF-8 кодировка каждого файла | Fail |
| 3. Check VBA syntax (basic) | Отсутствие недопустимых символов (коды < 32, кроме \n\r\t) | Fail |
| 4. Check CHANGELOG updated | Наличие записи за сегодняшнюю дату | Warning (non-blocking) |
| 5. Check .gitignore consistency | Наличие файла `.gitignore` | Fail |

### 7.2 Pre-commit процедура

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
- [`docs/ARCHITECTURE_SQLITE.md`](ARCHITECTURE_SQLITE.md) — архитектура выноса данных в SQLite
- [`CHANGELOG.md`](../CHANGELOG.md) — история версий проекта

---

## Приложение A: Схема зависимостей модулей

```
Лист2_main.cls
  └── Mod_OrderHeader.FillHeaderFromOrder()
        └── Mod_Utils (GetSheetByName, FileExists, FormatDateSQL)

Sheet_work.cls ─── Mod_SheetButtons
Sheet_z4.cls   ─── Mod_SheetButtons

Mod_ButtonDispatcher
  ├── Mod_SheetOps (ClearMainSheet_UI, ClearHeader_UI)
  ├── Mod_Import (ImportSheet_UI, ImportByInput_UI, RenameSheets_UI, ImportDataToMain_UI)
  ├── Mod_OrderHeader (FillHeaderFromOrder_UI, FindOrder_UI)
  ├── Mod_FullTestRunner (RunAllTests_UI)
  └── Mod_Utils (WriteLog_UI, ShowWorkbookPath_UI, ShowCurrentUser_UI, CheckFileExists_UI)

Mod_MainButtons
  ├── Mod_Import (вызовы импорта)
  ├── Mod_OrderHeader (заполнение шапки)
  └── Mod_SheetOps (операции с листами)

Mod_SheetButtons
  ├── Mod_Import (вызовы импорта)
  └── Mod_SheetOps (операции с листами)

Mod_FullTestRunner
  ├── Mod_Utils (тесты утилит)
  ├── Mod_Logger (тесты логгера)
  └── Mod_Constants (тесты реестра имён)

Mod_Logger
  └── (используется всеми модулями для логирования)

Mod_Constants
  └── (используется Mod_OrderHeader, Mod_Import и другими модулями)

Mod_Constants
  └── Mod_Utils (GetSheetByName)
  └── Mod_Logger (WriteLog)
```

## Приложение B: Маппинг файлов VBA

| VBA-компонент | Файл на диске | Тип модуля |
|--------------|---------------|------------|
| `Mod_Utils` | `src/modules/Mod_Utils.bas` | Стандартный модуль |
| `Mod_OrderHeader` | `src/modules/Mod_OrderHeader.bas` | Стандартный модуль |
| `Mod_Import` | `src/modules/Mod_Import.bas` | Стандартный модуль |
| `Mod_ButtonDispatcher` | `src/modules/Mod_ButtonDispatcher.bas` | Стандартный модуль |
| `Mod_FullTestRunner` | `src/modules/Mod_FullTestRunner.bas` | Стандартный модуль |
| `Mod_Logger` | `src/modules/Mod_Logger.bas` | Стандартный модуль |
| `Mod_Constants` | `src/modules/Mod_Constants.bas` | Стандартный модуль |
| `Mod_SheetOps` | `src/modules/Mod_SheetOps.bas` | Стандартный модуль |
| `Mod_MainButtons` | `src/modules/Mod_MainButtons.bas` | Стандартный модуль |
| `Mod_SheetButtons` | `src/modules/Mod_SheetButtons.bas` | Стандартный модуль |
| `Лист2` | `src/sheets/Лист2_main.cls` | Класс листа |
| `Sheet_work` | `src/sheets/Sheet_work.cls` | Класс листа |
| `Sheet_z4` | `src/sheets/Sheet_z4.cls` | Класс листа |

## Приложение C: Скрипты автоматизации

| Скрипт | Назначение | Кодировка |
|--------|-----------|-----------|
| [`export_vba.py`](../scripts/export_vba.py) | Выгрузка VBA из Excel на диск (CP1251 → UTF-8) | UTF-8 |
| [`impVBA.py`](../scripts/impVBA.py) | Загрузка VBA с диска в Excel (UTF-8 → CP1251) | UTF-8 |
| [`run_tests.py`](../scripts/run_tests.py) | Запуск тестов VBA | UTF-8 |
| [`scripts/Import-VbaFromExcel.ps1`](../scripts/Import-VbaFromExcel.ps1) | Альтернативный импорт VBA из Excel (PowerShell) | UTF-8 with BOM |