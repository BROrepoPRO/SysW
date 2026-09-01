# Техническая документация разработчика — SysW (v1.1.6.2)

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
- **Кнопки листов z4/work** ([`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas)) — кнопки листов z4/work
- **Константы и реестр имён** ([`Mod_Constants.bas`](../src/modules/Mod_Constants.bas)) — константы столбцов и управление листом libname
- **Типы данных моделей** ([`Mod_ModelTypes.bas`](../src/modules/Mod_ModelTypes.bas)) — Type-структуры модельных данных
- **Доступ к файлам групп** ([`Mod_ModelDB.bas`](../src/modules/Mod_ModelDB.bas)) — открытие файлов групп, чтение работ/запчастей
- **Ручной подбор работ** ([`Mod_PickWork.bas`](../src/modules/Mod_PickWork.bas)) — ручной подбор работ из справочника группы
- **Автоматический подбор** ([`Mod_AutoMatch.bas`](../src/modules/Mod_AutoMatch.bas)) — автоматический подбор работ/запчастей
- **Объект детали** ([`PartIdentity.cls`](../src/classes/PartIdentity.cls)) — класс-объект детали
- **Объект работы** ([`WorkIdentity.cls`](../src/classes/WorkIdentity.cls)) — класс-объект работы
- **Объект записи работы** ([`WorkEntry.cls`](../src/classes/WorkEntry.cls)) — класс-объект записи работы
- **Интерфейс провайдера данных** ([`IModelDataProvider.cls`](../src/classes/IModelDataProvider.cls)) — контракт доступа к данным моделей
- **Провайдер доступа к БД** ([`Mod_ModelDBProvider.cls`](../src/classes/Mod_ModelDBProvider.cls)) — фабрика/провайдер данных
- **SQLite-провайдер** ([`Mod_SQLiteDB.cls`](../src/classes/Mod_SQLiteDB.cls)) — доступ к данным из единой базы `SysW.db`
- **Лист main** ([`Лист2.cls`](../src/sheets/Лист2.cls)) — обработчик событий листа (`Worksheet_Change` для B4, `Worksheet_Activate` → FreezePanes A4)
- **Листы libname/models/spisok** ([`Лист3.cls`](../src/sheets/Лист3.cls), [`Лист5.cls`](../src/sheets/Лист5.cls), [`Лист9.cls`](../src/sheets/Лист9.cls)) — классы листов с событием `Worksheet_Activate` → `ApplyFreezePanes`

> Фактический состав `src/sheets/`: **`Лист2.cls`, `Лист3.cls`, `Лист5.cls`, `Лист9.cls`**.
> Устаревшие упоминания `Лист2_main.cls` и `Лист4.cls` в старой документации фактическому составу не соответствуют.

### 1.2 Единый стандарт структуры листов (v1.0.9)

Рабочие листы `work.xlsm` (`main`, `spisok`, `models`, `libname`) и модельные листы
приведены к единому стандарту: строки 1–2 технические, строка 3 — заголовки,
данные с 4-й строки, закрепление первых трёх строк (`FreezePanes A4`).

Маппинг:
- `main` — **без сдвига** (ввод № заказа `B4`, шапка `B5:B17`); маппинг не изменялся.
- `spisok`, `libname` — заголовки стр.1→3, данные со 2-й→4-й (сдвиг +2).
- `models` — двухрядная шапка сведена к строке 3 (русские названия); ключи остаются
  только в константах `MODELS_COL_MODEL_NAME`/`MODELS_COL_GROUP_NAME`/`MODELS_COL_PRICE_NAME`;
  данные с 4-й строки.
- Модельные листы (`z4`, `{GroupName}*`) — без сдвига; только FreezePanes A4
  (через `apply_freeze_panes_to_models` — **Вариант C, XML-инъекция `<pane>`** в
  `sheet*.xml`, без пересохранения через openpyxl/`keep_vba` и без VBA-классов).

Единые константы в [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas):
`HEADER_ROW=3`, `DATA_START_ROW=4`, `FREEZE_START_CELL="A4"`; по листам
`MAIN_/SPISOK_/MODELS_/LIBNAME_{HEADER_ROW,DATA_START_ROW}`; `MAIN_INPUT_CELL="B4"`.
`MAIN_HEADER_START_ROW` (ошибочная, =4) удалена и заменена на `MAIN_HEADER_ROW=3`.

Закрепление выполняется в `Mod_SheetOps.ApplyFreezePanes` (точка `FREEZE_START_CELL`
→ `ActiveWindow.FreezePanes = True`) и вызывается из `Worksheet_Activate` классов
листов `Лист2`..`Лист5`. При добавлении/переименовании листов обновляйте список
`COMPONENTS` в [`scripts/export_vba.py`](../scripts/export_vba.py) (схема импорта
`impVBA.py` обнаруживает классы листов по префиксу `Лист`/`Sheet` автоматически).

### 1.2 Схема взаимодействия модулей

```
Лист2.cls (Worksheet_Change)
       │
       └── Mod_OrderHeader.FillHeaderFromOrder()
              │
              ├── Mod_Utils.GetSheetByName()
              ├── Mod_Utils.FileExists()
              └── Mod_Utils.FormatDateSQL()

Mod_ButtonDispatcher (обработчики кнопок)
       │
       ├── Mod_SheetOps.ClearMainSheet_UI()
       ├── Mod_Import.ImportDataToMain_UI()
       ├── Mod_SheetOps.ClearHeader_UI()
       ├── Mod_OrderHeader.FillHeaderFromOrder_UI()
       └── ...

Mod_SheetButtons (кнопки листов z4/work)
       │
       ├── Mod_Import (вызовы импорта)
       └── Mod_SheetOps (операции с листами)

Mod_FullTestRunner.RunAllTests()
       │
       ├── RunUtilsTests()         → Mod_Utils (TC-01..TC-08)
       ├── RunLoggerTests()        → Mod_Logger (TC-09..TC-11)
       ├── RunUtilsEdgeTests()     → Mod_Utils (TC-12)
       ├── RunLibNameTests()       → Mod_Constants (TC-13)
       ├── RunImportVHTests()      → Mod_Import (TC-14)
       ├── RunModelDBTests()       → Mod_ModelDB (TC-31..TC-35)     ★ NEW
       ├── RunPickWorkTests()      → Mod_PickWork (TC-36..TC-38)    ★ NEW
       └── RunAutoMatchTests()     → Mod_AutoMatch (TC-39..TC-44)   ★ NEW
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

**Назначение:** Заполнение заголовка заказа-наряда (B5:B17) на листе main данными из листов spisok и models.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `FillHeaderFromOrder(OrderNum)` | Основная функция. Ищет номер заказа в листе spisok, заполняет B5:B17 на листе main. Возвращает `Boolean` — успех/неудача |
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

**Файл:** [`Mod_Import.bas`](../src/modules/Mod_Import.bas) (447 строк)

**Назначение:** Импорт данных из отчётов Excel. Поиск листов в report.xlsx по ГРЗ.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `ImportSheet(grz)` | Импортирует лист из report.xlsx по ГРЗ в текущую книгу |
| `ImportDataToMain(wsSource)` | Переносит данные из листа-источника в лист main по столбцам |
| `ImportDataToMain_UI()` | Переносит данные с активного листа в лист main |
| `ImportFromB2_UI()` | Импорт данных на лист "мэйн" из листа {B4}M; если листа нет — копирует из report.xlsx |

### 2.4 Mod_ButtonDispatcher.bas — Диспетчер кнопок

**Файл:** [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) (153 строки)

**Назначение:** Содержит **только** однострочные вызовы UI-процедур. Является прослойкой между кнопками на формах и бизнес-логикой. Не содержит бизнес-логики.

**Обработчики:**

| Процедура | Вызов |
|-----------|-------|
| `Btn_main_Clear_Click()` | `Mod_SheetOps.ClearMainSheet_UI` |
| `Btn_main_ImportVH_Click()` | `Mod_Import.ImportFromB2_UI` — импорт + перенос данных по B4 |
| `Btn_main_FillHeader_Click()` | `Mod_OrderHeader.FillHeaderFromOrder_UI` |
| `Btn_main_ClearHeader_Click()` | `Mod_SheetOps.ClearHeader_UI` |
| `Btn_main_RunTests_Click()` | `Mod_FullTestRunner.RunAllTests_UI` |
| `Btn_main_WriteLog_Click()` | `Mod_Utils.WriteLog_UI` |
| `Btn_main_ImportDataToMain_Click()` | `Mod_Import.ImportDataToMain_UI` |
| `Btn_main_FindOrder_Click()` | `Mod_OrderHeader.FindOrder_UI` |
| `Btn_main_CheckFileExists_Click()` | `Mod_Utils.CheckFileExists_UI` |
| `Btn_main_AutoMatchWorks_Click()` | `Mod_AutoMatch.AutoMatchWorks` |
| `Btn_main_AutoMatchParts_Click()` | `Mod_AutoMatch.AutoMatchParts` |
| `Btn_main_PickWork_Click()` | `Mod_PickWork.PickWork_UI` |
| `Btn_main_PickParts_Click()` | `Mod_PickWork.PickParts_UI` — ручной подбор запчастей `{Group}z4`/`z4` |
| `Btn_Search_ByArticle_Click()` | `Mod_SheetButtons.Btn_Search_ByArticle` — поиск по артикулу (столбец B) |
| `Btn_Search_ByName_Click()` | `Mod_SheetButtons.Btn_Search_ByName` — поиск по наименованию (столбец C) |
| `Btn_Search_Clear_Click()` | `Mod_SheetButtons.Btn_ClearFilter` — сброс фильтра + очистка C1 |

### 2.5 Mod_FullTestRunner.bas — Тестовый раннер

**Файл:** [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)

**Назначение:** Автоматическое тестирование VBA-модулей. Набор сценариев **TC-01..TC-64** + **TC-S1..TC-S3** (полный перечень — см. [Приложение F](#приложение-f-реестр-тестов-mod_fulltestrunner)). Результаты каждого теста дублируются в расширенный тестовый лог `logs/test_results.log` (см. §2.6).

**Группы тестов:**

| Группа | Сценарии | Тестируемый модуль |
|--------|----------|-------------------|
| `RunUtilsTests()` | TC-01..TC-08 | Mod_Utils |
| `RunLoggerTests()` | TC-09..TC-11 | Mod_Logger |
| `RunUtilsEdgeTests()` | TC-12 | Mod_Utils (граничные случаи) |
| `RunLibNameTests()` | TC-13 | Mod_Constants |
| `RunImportVHTests()` | TC-14 | Mod_Import (ImportFromB2_UI) |
| `RunSheetOpsTests()` | TC-15..TC-18, TC-45 | Mod_SheetOps |
| `RunAggregateNameTests()` | TC-19..TC-21 | Mod_Constants |
| `RunModelDBReadTests()` | TC-22..TC-24 | Mod_ModelDB / провайдер |
| `RunOrderHeaderTests()` | TC-25..TC-28 | Mod_OrderHeader |
| `RunImportDataTests()` | TC-29, TC-30 | Mod_Import |
| `RunImportB2IntegrationTests()` | TC-65..TC-68 | Mod_Import (ImportFromB2_UI сквозной: импорт+шапка+работы+запчасти) |
| `RunModelDBTests()` | TC-31..TC-35, TC-64 | Mod_ModelDB |
| `RunPickWorkTests()` | TC-36..TC-38, TC-63 | Mod_PickWork / Mod_Utils |
| `RunAutoMatchTests()` | TC-39..TC-44 | Mod_AutoMatch |
| `RunConstantsTests()` | TC-46 | Mod_Constants |
| `RunSQLiteTests()` | TC-S1..TC-S3, TC-47..TC-50 | Mod_SQLiteDB / провайдер |
| `RunSearchTests()` | TC-51..TC-55 | Mod_SheetButtons / Mod_PickWork |
| `RunGlobalBaseTests()` | TC-60..TC-62 | Mod_ModelDB (глобальная база з/ч, fallback) |

**Механизмы:**
- **SKIP** — тесты, зависящие от отсутствующих данных, пропускаются
- **Сохранение/восстановление** состояния листов до и после тестов
- **Статистика** — подсчёт Total, Passed, Failed, Skipped
- **Детальный тестовый лог** — каждая строка результата пишется через `Mod_Logger.WriteTestLog` в `logs/test_results.log` (PASS→INFO, SKIP→WARN, FAIL→ERROR)
- **GetTestResults()** — **Public Sub**: записывает отчёт в ячейку `Z1` листа main для чтения из Python (не функция)
- **Silent mode** — параметр `silent` в `ClearMainSheet_UI` для автоматических тестов без MsgBox

**Запуск:**
```bash
python scripts/run_tests.py
```

### 2.6 Mod_Logger.bas — Логирование «2 лога» (Задача 1, v1.1.0)

**Файл:** [`Mod_Logger.bas`](../src/modules/Mod_Logger.bas)

**Назначение:** Централизованное логирование, разделённое на **два независимых лога**:

- **Системный лог** — `logs/log.txt`: общие системные события (текущее поведение).
- **Расширенный тестовый лог** — `logs/test_results.log`: детальный лог тестов
  с уровнями **INFO/WARN/ERROR**, включая ошибки VBA.

Пути задаются константами [`Mod_Constants.LOG_FILE`](../src/modules/Mod_Constants.bas)
(`log.txt`) и `Mod_Constants.TEST_LOG_FILE` (`test_results.log`) в директории
`Mod_Constants.LOGS_DIR`.

**Ключевые функции:**

| Функция | Лог | Описание |
|---------|-----|----------|
| `WriteLog(ModuleName, Message)` | системный | Запись сообщения |
| `WriteLogE(ModuleName, Message)` | системный | Запись ошибки с префиксом [ERROR] |
| `WriteTestLog(ModuleName, Level, Message)` | тестовый | Запись с уровнем INFO/WARN/ERROR |
| `LogTestInfo / LogTestWarn / LogTestError(...)` | тестовый | Обёртки фиксированного уровня |
| `RotateLogIfNeeded(MaxSizeKB)` | системный | Ротация при превышении размера |
| `ClearLog()` | системный | Очистка системного лога |
| `ClearTestLog()` | тестовый | Очистка тестового лога |
| `GetLogPath()` | системный | Путь к системному логу |
| `GetTestLogPath()` | тестовый | Путь к тестовому логу |

> Сопоставление статусов теста и уровней тестового лога: PASS→INFO, SKIP→WARN, FAIL→ERROR
> (реализуется в `Mod_FullTestRunner.AddResult`).

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
| `ClearHeader_UI()` | Очищает шапку заказа (B5:B17) на листе main |

### 2.9 Mod_ModelTypes.bas — Типы данных моделей

**Файл:** [`Mod_ModelTypes.bas`](../src/modules/Mod_ModelTypes.bas)

**Назначение:** Централизованное хранение Type-структур модельных данных, используемых модулями `Mod_ModelDB`, `Mod_PickWork` и `Mod_AutoMatch`.

**Типы данных:**

| Тип | Поле | Описание |
|-----|------|----------|
| `PartEntry` | `Code As String` | Код запчасти |
| | `Name As String` | Наименование запчасти |
| | `Unit As String` | Единица измерения |
| | `Price As Currency` | Цена запчасти |
| | `Note As String` | Примечание |
| `WorkEntry` | `Code As String` | Код работы |
| | `Name As String` | Наименование работы |
| | `Unit As String` | Единица измерения |
| | `NormHours As Double` | Норматив в нормо-часах |
| | `Price As Currency` | Цена работы |
| | `Note As String` | Примечание |
| `MatLibEntry` | `Code As String` | Код материала |
| | `Name As String` | Наименование материала |
| | `Unit As String` | Единица измерения |
| | `Price As Currency` | Цена материала |
| `ModelPartEntry` | `Part As PartEntry` | Запчасть модели |
| | `Model As String` | Модель |
| `ModelWorkEntry` | `Work As WorkEntry` | Работа модели |
| | `Model As String` | Модель |

### 2.10 Mod_SheetButtons.bas — Универсальный поиск на листах работ и запчастей

**Файл:** [`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas)

**Назначение:** Универсальные обработчики кнопок поиска/фильтрации на листах работ и запчастей любой группы. С v1.0.12 имя группы динамически читается из `main!$B$14` (с fallback на имя листа), поэтому имена листов в код **не зашиты** — поиск работает на `{Group}`, `{Group}w`, `z4`, `{Group}z4` любой группы. Модуль содержит **только** поисковые обработчики и приватные помощники. Заявленные в старой документации заглушки `Btn_z4_Action*`/`Btn_work_Action*` в коде **отсутствуют**.

**Классификация листов (`ClassifySheet`, enum `SheetKind`):**

| Тип | Значение | Лист |
|-----|----------|------|
| `skWorks` | 1 | `{Group}` (работы группы) |
| `skWorksModel` | 2 | `{Group}w` |
| `skParts` | 3 | `z4` (общие запчасти) |
| `skPartsModel` | 4 | `{Group}z4` (запчасти группы) |
| `skUnknown` | 0 | прочее (в поиске не участвует) |

Имя группы определяется через единый публичный хелпер `Mod_Utils.GetGroupName()`
(чтение `main!$B$14`) и `ResolveGroupName(ws)` — при пустом `B14` (например,
в модельной книге без листа main)
группа выводится из имени листа (`{Group}`, `{Group}w`, `{Group}z4`; для `z4` — пустая строка).
`ClassifySheet(ws, groupName)` — публичная функция классификации листа в `SheetKind`.

**Публичные обработчики:**

| Процедура | Описание |
|----------|----------|
| `Btn_Search_ByArticle()` | Поиск «содержит» по столбцу B (артикул) на активном листе работ/запчастей |
| `Btn_Search_ByName()` | Поиск «содержит» по столбцу C (наименование) на активном листе работ/запчастей |
| `Btn_ClearFilter()` | Сброс фильтра и очистка поля ввода C1 |

**Приватные помощники:**

| Процедура | Описание |
|----------|----------|
| `ExecuteSearch(searchColumn)` | Единый алгоритм поиска «содержит» (классификация листа, поле C1, данные с 4-й строки, AutoFilter) |
| `Mod_Utils.GetGroupName()` | Единый публичный хелпер чтения имени группы из `main!$B$14` (централизация дублей `Mod_SheetButtons`/`Mod_PickWork`/`Mod_AutoMatch`) |
| `ResolveGroupName(ws)` | Определяет группу: `B14` либо из имени листа |
| `IsSearchableSheet(ws)` | Проверка, что лист подходит для поиска (данные с 4-й строки) |

> Ранее существовавшие `Btn_UAZ_SearchByArticle`/`Btn_UAZ_SearchByName`/`Btn_UAZ_ClearFilter`,
> `Btn_Parts_SearchByArticle`/`Btn_Parts_SearchByName`/`Btn_Parts_ClearFilter` и
> `ExecuteUAZSearch`/`ExecutePartsSearch` переименованы в нейтральные имена выше (v1.0.12).

### 2.11 Mod_ModelDB.bas — Доступ к файлам модельных групп

**Файл:** [`Mod_ModelDB.bas`](../src/modules/Mod_ModelDB.bas) (167 строк)

**Назначение:** Базовый слой абстракции для работы с файлами модельных групп в `base/models/`. Обеспечивает открытие файлов групп и чтение данных работ/запчастей.

**Путь к каталогу моделей:**

| Функция | Описание |
|---------|----------|
| `GetModelDBBasePath()` | Возвращает каталог файлов групп `base\models\`. Константы `MODELDB_BASE_PATH` в коде нет — она заменена функцией |

**Типы данных:**

| Тип | Поле | Описание |
|-----|------|----------|
| `WorkEntry` | `Code As String` | Код работы |
| | `Name As String` | Наименование работы |
| | `Unit As String` | Единица измерения |
| | `NormHours As Double` | Норматив в нормо-часах |
| | `Price As Currency` | Цена работы |
| | `Note As String` | Примечание |

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `GetModelGroupFilePath(groupName)` | Возвращает полный путь к файлу группы `base/models/{groupName}.xlsx` |
| `ModelGroupFileExists(groupName)` | Проверяет существование файла группы через `Dir()` |
| `OpenModelGroupFile(groupName)` | Открывает файл группы (если ещё не открыт), возвращает `Workbook` или `Nothing` |
| `GetWorks(groupName, filters)` | Возвращает коллекцию `WorkEntry` из листа `{groupName}` файла группы |
| `GetModelDataProvider(ByRef provider)` | Фабрика провайдера данных (SQLite при наличии SysW.db, иначе Excel-fallback) |
| `GetParts/GetModelWorks/GetModelParts/GetMatLibEntries/GetWorkIdentities/GetPartIdentities/GetAllModelGroups/CreateModelGroupFile/FindModelGroupByModel` | Делегаты чтения/создания данных через выбранный провайдер |
| `GetGlobalPartsBasePath()` | Путь к глобальному файлу запчастей `base/models/z4.xlsx` (v1.0.13) |
| `ReadGlobalPartByKey(catNum, partName)` | Поиск з/ч в глобальной базе `z4.xlsx`: приоритет № кат. (B), fallback наименование (C); возвращает `PartIdentity` или `Nothing` |
| `AppendPartIdentity(groupName, identity)` | Запись тождества запчасти на лист `{GroupName}z4` модельной книги |
| `AppendWorkIdentity(groupName, identity)` | Запись тождества работы на лист `{GroupName}w` модельной книги |
| `ReadLocalWorkByName(groupName, workName)` | Поиск работы в локальном листе `{GroupName}` по наименованию (fallback для работ) |

### 2.12 Mod_PickWork.bas — Ручной подбор работ и запчастей

**Файл:** [`Mod_PickWork.bas`](../src/modules/Mod_PickWork.bas) (279 строк)

**Назначение:** Ручной подбор работ и запчастей из справочника группы. Открывает файл группы, активирует нужный лист, пользователь ищет через фильтр и копирует данные вручную в диапазон E4:H (работы) или P:W/X:AB (запчасти) на листе main.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `GetGroupNameFromMain()` | Читает название группы из ячейки B14 листа main |
| `GetWorkSheetName(groupName)` | Возвращает имя листа работ (совпадает с именем группы) |
| `GetPartsSheetName(wb, groupName)` | Имя модельного листа запчастей `{Group}z4`; при отсутствии — общий `z4` |
| `PickWork_UI()` | Ручной подбор работ (кнопка «РУЧ РАБ»): открывает файл группы, активирует лист работ, показывает инструкцию |
| `PickParts_UI()` | Ручной подбор запчастей (кнопка «РУЧ ЗЧ»): открывает файл группы, активирует `{Group}z4`/`z4`, показывает инструкцию |

**Процесс работы:**
1. Пользователь нажимает кнопку **РУЧ РАБ** на листе `main`
2. Макрос читает группу из **B14**
3. Открывает файл `base/models/{Group}.xlsx` через `Mod_ModelDB.OpenModelGroupFile`
4. Активирует лист `{GroupName}` (справочник работ)
5. Пользователь ищет работы через фильтр (скрипт поиска внутри файла группы)
6. Проставляет количество в столбце G (Кол-во ЗН)
7. Фильтрует по G — всё кроме 0
8. Копирует отфильтрованные строки вручную в **E4:H** на лист `main`

### 2.13 Mod_AutoMatch.bas — Автоматический подбор

**Файл:** [`Mod_AutoMatch.bas`](../src/modules/Mod_AutoMatch.bas)

**Назначение:** Автоматический подбор работ/запчастей. Выполняет сопоставление данных автомобиля со справочниками модельных групп.

**Ключевые функции:**

| Функция | Описание |
|---------|----------|
| `AutoMatchWorks()` | Автоподбор работ (L→E:I + формула J). Точка входа кнопки «АВТО РАБ» |
| `AutoMatchParts()` | Автоподбор запчастей (X→Q:U + формула V). Точка входа кнопки «АВТО ЗЧ» |

### 2.14 Mod_Constants.bas — Константы и реестр имён

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
| `MANWRK_COL_ARTICLE` | 5 | Артикул (столбец E листа main) |
| `MANWRK_COL_NAME` | 6 | Наименование (столбец F листа main) |
| `MANWRK_COL_NORMHOURS` | 7 | Кол-во н/ч (столбец G листа main) |
| `MANWRK_COL_QTY` | 8 | Кол-во оп (столбец H листа main) |
| `MANWRK_START_ROW` | 4 | Строка начала данных ручного подбора |

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

> **Примечание:** Ранее функциональность реестра имён находилась в отдельном модуле `Mod_LibName.bas` (удалён в версии 1.0.0), который был объединён с `Mod_Constants.bas` для централизованного управления константами.

### 2.15 Лист2.cls — Основной лист

**Файл:** [`Лист2.cls`](../src/sheets/Лист2.cls) (48 строк)

**Назначение:** Класс листа main. Обрабатывает события `Worksheet_Activate` (FreezePanes A4) и `Worksheet_Change`.

**Логика обработчика `Worksheet_Change`:**
1. Защита от рекурсии через `Static isProcessing As Boolean`
2. Проверка, что изменение произошло в ячейке B4
3. Очистка диапазона B5:B17
4. Вызов `Mod_OrderHeader.FillHeaderFromOrder(CStr(b4Value))`

### 2.16 PartIdentity.cls — Объект детали

**Файл:** [`PartIdentity.cls`](../src/classes/PartIdentity.cls)

**Назначение:** Класс-объект, представляющий деталь (запчасть). Используется для идентификации и передачи данных о детали между модулями.

### 2.17 WorkIdentity.cls — Объект работы

**Файл:** [`WorkIdentity.cls`](../src/classes/WorkIdentity.cls)

**Назначение:** Класс-объект, представляющий работу. Используется для идентификации и передачи данных о работе между модулями.

### 2.18 Структура листов work и z4

#### Лист work (работы)

Строка 4 — заголовки. Данные — с 5-й строки.

| Столбец | Назначение |
|---------|-----------|
| A | № |
| B | Артикул |
| C | Наименование |
| D | Кол-во н/ч |
| E | Кол-во оп |
| F | Цена н/ч |
| G | Сумма |

#### Лист z4 (запчасти)

Строка 4 — заголовки. Данные — с 5-й строки.

| Столбец | Назначение |
|---------|-----------|
| A | № |
| B | Артикул |
| C | Наименование |
| D | Ед.изм. |
| E | Кол-во |
| F | Цена |
| G | Сумма |

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
| `Mod_SheetButtons` | `src/modules/Mod_SheetButtons.bas` | Стандартный модуль |
| `Mod_ModelTypes` | `src/modules/Mod_ModelTypes.bas` | Стандартный модуль |
| `Mod_ModelDB` | `src/modules/Mod_ModelDB.bas` | Стандартный модуль |
| `Mod_PickWork` | `src/modules/Mod_PickWork.bas` | Стандартный модуль |
| `Mod_AutoMatch` | `src/modules/Mod_AutoMatch.bas` | Стандартный модуль |
| `IModelDataProvider` | `src/classes/IModelDataProvider.cls` | Класс |
| `Mod_ModelDBProvider` | `src/classes/Mod_ModelDBProvider.cls` | Класс |
| `Mod_SQLiteDB` | `src/classes/Mod_SQLiteDB.cls` | Класс |
| `PartIdentity` | `src/classes/PartIdentity.cls` | Класс |
| `WorkEntry` | `src/classes/WorkEntry.cls` | Класс |
| `WorkIdentity` | `src/classes/WorkIdentity.cls` | Класс |
| `Лист2` | `src/sheets/Лист2.cls` | Класс листа |

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


---

## 5. Настройка окружения

### 5.1 VS Code

1. Открыть корневую директорию проекта (репозиторий SysW)
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

### 6.1 Реестр тестов (TC-01..TC-62 + TC-S1..S3)

Полный перечень тестов с группами, статусами, механизмом записи результатов и замечаниями
приведён в [Приложении F](#приложение-f-реестр-тестов-mod_fulltestrunner) ниже. Здесь — сводка по группам и модулям.

| Группа | Сценарии | Модуль | Зависимость от окружения |
|--------|----------|--------|--------------------------|
| `RunUtilsTests()` | TC-01..TC-08 | Mod_Utils | нет |
| `RunLoggerTests()` | TC-09..TC-11 | Mod_Logger | нет |
| `RunUtilsEdgeTests()` | TC-12 | Mod_Utils | нет |
| `RunLibNameTests()` | TC-13 | Mod_Constants | лист libname |
| `RunImportVHTests()` | TC-14 | Mod_Import | лист main |
| `RunSheetOpsTests()` | TC-15..TC-18, TC-45 | Mod_SheetOps | нет |
| `RunAggregateNameTests()` | TC-19..TC-21 | Mod_Constants | нет |
| `RunModelDBReadTests()` | TC-22..TC-24 | Mod_ModelDB | SQLite-провайдер |
| `RunOrderHeaderTests()` | TC-25..TC-28 | Mod_OrderHeader | листы main/spisok/models |
| `RunImportDataTests()` | TC-29, TC-30 | Mod_Import | лист main |
| `RunImportB2IntegrationTests()` | TC-65..TC-68 | Mod_Import | лист main/spisok/models, временный лист {B4}M |
| `RunModelDBTests()` | TC-31..TC-35, TC-64 | Mod_ModelDB | файлы групп/провайдер (TC-64 — SQLite) |
| `RunPickWorkTests()` | TC-36..TC-38, TC-63 | Mod_PickWork / Mod_Utils | нет |
| `RunAutoMatchTests()` | TC-39..TC-44 | Mod_AutoMatch | нет (TC-41..44 — SKIP) |
| `RunConstantsTests()` | TC-46 | Mod_Constants | лист libname |
| `RunSQLiteTests()` | TC-S1..TC-S3, TC-47..TC-50 | Mod_SQLiteDB | SQLite-провайдер/SysW.db |
| `RunSearchTests()` | TC-51..TC-55 | Mod_SheetButtons / Mod_PickWork | лист main (B14) |
| `RunGlobalBaseTests()` | TC-60..TC-62 | Mod_ModelDB | глобальная база `z4.xlsx` / файлы групп |

**Легенда:** PASS — тест проходит; SKIP — пропущен (Private-процедура, небезопасность автотеста,
отсутствие листа/данных/SQLite-провайдера); FAIL — падение теста (останавливает `run_tests.py` с кодом 1).

### 6.2 Таблица покрытия модулей

| Модуль | Группы | Сценарии |
|--------|--------|----------|
| Mod_Utils | RunUtilsTests, RunUtilsEdgeTests | TC-01..08, TC-12, TC-63 |
| Mod_Logger | RunLoggerTests | TC-09..11 |
| Mod_Constants | RunLibNameTests, RunAggregateNameTests, RunConstantsTests | TC-13, TC-19..21, TC-46 |
| Mod_Import | RunImportVHTests, RunImportDataTests, RunImportB2IntegrationTests | TC-14, TC-29, TC-30, TC-65..TC-68 |
| Mod_SheetOps | RunSheetOpsTests | TC-15..18, TC-45 |
| Mod_OrderHeader | RunOrderHeaderTests | TC-25..28 |
| Mod_AutoMatch | RunAutoMatchTests | TC-39..44 (TC-41..44 — SKIP) |
| Mod_PickWork | RunPickWorkTests | TC-36..38, TC-63 |
| Mod_ModelDB | RunModelDBTests, RunModelDBReadTests, RunGlobalBaseTests | TC-22..24, TC-31..35, TC-64, TC-60..62 |
| Mod_SQLiteDB | RunSQLiteTests | TC-S1..S3, TC-47..50 |
| Mod_SheetButtons | RunSearchTests | TC-51..54 |
| Mod_PickWork | RunSearchTests | TC-55 |
| Mod_ModelDB | RunModelDBTests (TC-64), RunGlobalBaseTests | TC-64, TC-60..62 |

> Точная статистика PASS/SKIP/FAIL зависит от окружения (наличие `SysW.db`, листов, данных)
> и формируется по итогам прогона `python scripts/run_tests.py`.

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
4. Собирает результаты через VBA-процедуру `GetTestResults()` (запись в ячейку `Z1` листа main)
5. Парсит статистику (Total, Passed, Failed, Skipped)
6. Возвращает exit code: `0` — все PASS, `1` — есть FAIL
7. Дописывает служебные строки прогона в `logs/test_results.log` (тот же файл, куда VBA
   пишет детальный лог через `Mod_Logger.WriteTestLog` с уровнями INFO/WARN/ERROR;
   константа пути — `TEST_LOG_FILE` из `scripts/config.py`)
8. Гарантированно закрывает Excel (в `finally` блоке)

**CLI-опции (R-16, v1.1.1.1):**
- `python scripts/run_tests.py --module <имя>` — выборочный запуск тестов модуля;
- `python scripts/run_tests.py --verbose` — расширенный вывод прогона;
- `python scripts/run_tests.py --output <путь>` — генерация отчётов JSON/HTML
  (UTF-8, кириллица корректна);
- без аргументов — поведение прежнее, отчёты не создаются; exit code — по полной
  сводке Z1.

> **Обкатка и устойчивость (v1.1.1.1):** базовый прогон зелёный
> (Total=65/Passed=56/Failed=0/Skipped=9, exit 0); бизнес-прогон `run_p1_business_test.py`
> (сценарий ImportFromB2_UI → AutoMatchWorks → AutoMatchParts) прошёл без зависания после
> фикса `Mod_ModelDB.OpenModelGroupFile` (защитные параметры вокруг `Workbooks.Open`);
> защита от бесконечного ожидания при блокировке `SysW.db` — `ConnectionTimeout`/
> `CommandTimeout = 15` в `Mod_SQLiteDB.cls`. `report.xlsx` пользователя не изменён.

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
| 1. Check VBA files exist | Наличие всех 16 VBA-файлов (13 `.bas` + 2 класса + 1 класс листа) | Fail |
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

## 8. Шаблоны `base/templates/` (v1.0.4)

### 8.1 Состав каталога

| Шаблон | Назначение | Источник | Код VBA |
|--------|-----------|----------|---------|
| `base/templates/work.xlsm` | Шаблон рабочей книги (интерфейс) | корневой `work.xlsm` после `scripts/impVBA.py` | ✅ есть |
| `base/templates/work0.xlsm` | Шаблон work **пустой** | корневой `work.xlsm` | ❌ нет (VBA удалён) |
| `base/templates/model.xlsm` | Общий шаблон модельного файла | объединение структур `base/models/*.xlsm` | ✅ есть |
| `base/templates/model0.xlsm` | Шаблон модельного файла **пустой** | объединение структур `base/models/*.xlsm` | ❌ нет (VBA удалён) |
| `base/templates/report0.xlsx` | Шаблон файла-отчёта **пустой** | корневой `report.xlsx` | ❌ нет (xlsx) |

Состав листов (см. [`docs/table.md`](table.md), раздел 0.11):
- `work.xlsm` / `work0.xlsm`: `main`, `spisok`, `models`, `libname`;
- `model.xlsm` / `model0.xlsm`: `z4`, `{GroupName}`, `{GroupName}w`, `{GroupName}z4`;
- `report0.xlsx`: `report`, `spisok`.

Временные/архивные/дубликаты (`temp*`, `*OLD`, `ЗНW`, `СчетW`, `3M`, `_SETTINGS`)
в шаблоны **не включаются** (консервативно по docs/table.md 0.11).

### 8.2 Защита листов (`Protect` + `AllowEditRanges` + `FreezePanes`)

На листах шаблонов (`main`, листы ЗЧ, листы работ, прочие листы с заголовками)
применены:

1. **3 верхние строки (строки 1–3) блокируются** (`Locked=True`) и **сохраняют формат,
   заданный в шаблоне** — не изменяются при работе.
2. **`FreezePanes` на `A4`** — закреплены строки 1–3 (2 под кнопки + заголовки)
   на листах ЗЧ/работ/main и других листах с заголовками.
3. **Для `main`:** столбцы `A:B` — только **защита** (без закрепления столбцов);
   закрепление остаётся **только на `A4`** (закреплены только строки 1–3).
4. **`AllowEditRanges`** (зоны ввода/данных):
   - `main`: `B4`, `B5:B17`, `C1`, `Z1`, данные с 4-й строки (`D4:AB2000`);
   - листы работ/запчастей: `C1` (поле поиска), данные с 4-й строки (`A4:U2000`);
   - `spisok`/`models`/`libname`/`report`/`spisok`: данные с 2–3-й строки.

Реализовано **на уровне файлов-шаблонов** без изменения VBA-кода. Кнопки остаются
кликабельными при `Protect`; привязка к `Mod_ButtonDispatcher`/`Mod_SheetButtons` сохраняется.

### 8.3 Процедура пересоздания шаблона `work.xlsm`

1. `python scripts/impVBA.py` — импорт `src/` в корневой `work.xlsm`.
2. Копия корневого `work.xlsm` → `base/templates/work.xlsm`.
3. Повторно применить защиту листов (скрипт `scripts/apply_protection_templates.py`
   + внедрение `AllowEditRanges` на XML-уровне).

---

## 9. SQLite, единый конвейер, компиляция VBA и глубокая подстановка (v1.0.8)

### 9.1 SQLite-хранилище `SysW.db`

Единая база данных SQLite расположена в корне проекта (`SysW.db`), DDL-схема — `db/schema.sql`. Доступ из VBA реализован через классы `Mod_SQLiteDB.cls`, `Mod_ModelDBProvider.cls` и интерфейс `IModelDataProvider.cls`. Таблица `works` использует суррогатный первичный ключ `id INTEGER PRIMARY KEY AUTOINCREMENT` (сохраняет дубли наименований). Пересборка базы из `base/models/*` выполняется скриптом `scripts/migrate_models_to_sqlite.py`.

### 9.2 Единый конвейер `scripts/build_all.py`

Полная пересборка проекта запускается одной командой:

```bash
python scripts/build_all.py
```

Порядок этапов: резервное копирование (`work.xlsm`, `SysW.db`) → `impVBA.py` → **ранний контроль компиляции VBA (`check_vba_syntax.py`)** → `build_templates.py` → `migrate_models_to_sqlite.py` → контроль целостности БД (`PRAGMA integrity_check` + контрольные количества) → `run_tests.py`. При провале любого этапа конвейер останавливается с ненулевым exit-кодом. Прогресс пишется в `logs/build.log`.

Начиная с v1.0.8 каждому этапу задан таймаут (`STEP_TIMEOUT`), а зависшие процессы
`EXCEL.EXE` принудительно завершаются по PID (`taskkill /F /PID`, безопасно — только
«свои» процессы). Новому этапу проверки VBA соответствует exit-код конвейера
**`vbacompile` = 22**; сам чекер возвращает `0` (ошибок нет) или `1` (найдены ошибки).

### 9.3 Флаг `ApplyMatLibSubstitution` и бизнес-правило подстановки

Публичный флаг `Mod_Constants.ApplyMatLibSubstitution = True` включает глубокую подстановку модельных кодов при импорте:

- запчасти — по № кат. **X(24)** с fallback по наименованию **Y(25)** → **AB(28)** (`target_type='mod_part'`);
- работы — по наименованию **L(12)** → **O(15)** (`target_type='mod_work'`).

Подстановка выполняется только при точном совпадении; наименования сохраняются. В `GetMatLibEntries` используется детерминированный `ORDER BY target_type, target_code`. При выключенном флаге поведение импорта идентично прежнему.

### 9.4 Ранний контроль компиляции VBA (`scripts/check_vba_syntax.py`)

Самодостаточный Python-скрипт статического анализа исходников VBA (`src/`), запускается
без Excel. Выполняется этапом конвейера `build_all.py` сразу после `impVBA.py`, чтобы
ловить синтаксические/компиляционные ошибки (типа `Public ... As Boolean = True` из
v1.0.7) **до** прогона тестов.

```bash
python scripts/check_vba_syntax.py
```

Что проверяет:
- запрещённую inline-инициализацию модульных/глобальных переменных вида
  `Public/Private ... As Boolean = <литерал>` (на уровне модуля допустимо только `Const`);
- несбалансированные блоки `Sub/Function/End Sub/End Function`, `If/End If`,
  `For/Next`, `With/End With`;
- отсутствие `Attribute VB_Name` у `.bas`/`.cls` или несовпадение с именем файла;
- дубликаты имён процедур в пределах одного модуля;
- нечитаемый файл (битый `utf-8`/`cp1251`), пустой обязательный модуль,
  лишний файл в `src/modules/`.

Exit-коды: `0` — ошибок нет; `1` — найдены ошибки (вывод списка `файл:строка:сообщение`).
В конвейере этапу соответствует код **22**.

---

## 10. Система отслеживания задач в ROADMAP (v1.0.11)

Начиная с v1.0.11 файл [`docs/ROADMAP.md`](ROADMAP.md) является **единым источником задач и планов**. В нём ведётся актуальный статус каждой задачи через Markdown-чекбоксы, а дерево задач визуализируется mermaid-схемой.

### 10.1 Структура раздела «Актуальные задачи и подзадачи»

Раздел «## 11. Актуальные задачи и подзадачи» содержит:

- **Mermaid-схему дерева задач** (`flowchart TD`) — текущие P0/P1 и глобальные цели верхнего уровня с рёбрами зависимостей.
- **Текущие задачи P0** (R-07..R-13) и **P1** (R-14..R-17, R-WS, R-ZS, R-DS, R-IT).
- **Глобальные цели верхнего уровня** (G-RM, G-Q1, G-SQ, G-TC, G-DO).

Единый формат строки каждой задачи:

```markdown
- [ ] **ID** · тип · связь · условие — описание
```

Атрибуты:
- **тип** — `текущая` (P0/P1) либо `глобальная` (цель верхнего уровня);
- **связь** — `независимая` либо `зависит от <ID/фазы>`;
- **условие** — `без условий` либо конкретное условие (например, доступ к `work.xlsm` в Excel);
- **`[ ]`** — не выполнено, **`[x]`** — выполнено.

Таблицы фаз (§3–§9) остаются справочным описанием; актуальный статус задач ведётся только в разделе §11.

### 10.2 Как SourceCraft обновляет статусы

1. **Чтение:** SourceCraft (роль Code/Orchestrator) читает `docs/ROADMAP.md` и находит строки вида `- [ ] **ID**`.
2. **Обновление:** после фактического выполнения задачи меняет `- [ ]` → `- [x]` в соответствующей строке раздела (через `apply_diff`/`write_file`).
3. **Проверка предшественников:** перед отметкой выполненной задачи убеждается, что её предшественники помечены `[x]`. Примеры цепочек: R-08/R-09 — только после R-07; R-13 — после R-12; R-14 — после Фазы 1.2 (R-13); R-DS — после R-WS и R-ZS.
4. **Глобальные цели:** отметка глобальной цели (`G-*`) допустима только после выполнения всех релевантных текущих задач.
5. **Фиксация:** после правок обновляется `CHANGELOG.md` (правило `[U2]`) и выполняется коммит по Conventional Commits.

### 10.3 Отображение в Todo Tree

Расширение **Todo Tree** настроено в `.vscode/settings.json`:

- Теги: `[ ]`, `[x]`, `TODO`, `FIXME`.
- Regex: `(^\s*[-*]\s*\[[ xX]\]|$TAGS)` — захватывает Markdown-чекбоксы `- [ ]`/`- [x]` и классические теги TODO/FIXME.
- Кастомная подсветка `[ ]`/`[x]`/`TODO`/`FIXME`.
- Исключения сканирования: `**/node_modules/**`, `**/.venv/**`, `**/logs/*.log`.

При добавлении новой задачи в ROADMAP используйте формат `- [ ] **ID** …`, чтобы она автоматически отображалась в дереве Todo Tree и была доступна SourceCraft для обновления.

### 10.4 Правила рабочей области (.ycarules и .codeassistant/rules/)

Правила поведения рабочей области хранятся в **двух связанных источниках** (иерархия приоритета — правило `[E5]`):

1. **Единый справочник [`.ycarules`](../.ycarules)** — корневой файл правил, подключённый через `.sourcecraft` (`customInstructions.file`). Структура — **8 тематических модулей** `[O][G][Z][K][U][S][T][E]` (Роли, Глобальные правила, Запреты и ограничения, Кодировка и синхронизация, Управление изменениями, Структура проекта, Автодополнение текста, Исключения и критические файлы).
2. **Зеркальный каталог [`.codeassistant/rules/`](../.codeassistant/rules/)** — рабочие правила рабочей области: **8 файлов-зеркал** (`01-roles.md` … `08-exceptions.md`) с кодами, идентичными `.ycarules`.

> `.ycarules` и `.codeassistant/rules/*` — критические файлы ([E3]); редактируются только по согласованию. Подробное описание — в [`docs/sourcecraft-guide.md`](sourcecraft-guide.md), раздел «Правила SourceCraft».

---

## Связанные документы

- [`README.md`](../README.md) — общее описание проекта, быстрый старт
- [`docs/sourcecraft-guide.md`](sourcecraft-guide.md) — руководство по работе с SourceCraft Code Assistant
- [`docs/git-workflow.md`](git-workflow.md) — Git-инструкции и веточная стратегия
- [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) — архитектура выноса данных в SQLite
- [`CHANGELOG.md`](CHANGELOG.md) — история версий проекта
- [`ROADMAP.md`](ROADMAP.md) — единый источник задач и планов (раздел «Актуальные задачи и подзадачи»)

---

## Приложение A: Схема зависимостей модулей

```
Лист2.cls
  └── Mod_OrderHeader.FillHeaderFromOrder()
        └── Mod_Utils (GetSheetByName, FileExists, FormatDateSQL)

Mod_ButtonDispatcher
  ├── Mod_SheetOps (ClearMainSheet_UI, ClearHeader_UI)
  ├── Mod_Import (ImportDataToMain_UI, ImportFromB2_UI)
  ├── Mod_OrderHeader (FillHeaderFromOrder_UI, FindOrder_UI)
  ├── Mod_FullTestRunner (RunAllTests_UI)
  ├── Mod_Utils (WriteLog_UI, CheckFileExists_UI)
  ├── Mod_SheetButtons (Btn_Search_ByArticle, Btn_Search_ByName, Btn_ClearFilter)
  └── Mod_PickWork (PickWork_UI, PickParts_UI)

Mod_SheetButtons
  ├── Mod_Constants (константы структуры листов, SilenceMsgBox)
  └── Mod_Logger (логирование ошибок)

Mod_FullTestRunner
  ├── Mod_Utils (тесты утилит TC-01..TC-08, TC-12)
  ├── Mod_Logger (тесты логгера TC-09..TC-11)
  ├── Mod_Constants (тесты реестра имён TC-13)
  ├── Mod_Import (тесты импорта TC-14, TC-21, TC-22, TC-30)
  ├── Mod_SheetOps (тесты операций с листами TC-18..TC-20, TC-25..TC-29)
  └── Mod_ButtonDispatcher (интеграционный тест TC-24)

Mod_Logger
  └── (используется всеми модулями для логирования)

Mod_Constants
  └── (используется Mod_OrderHeader, Mod_Import и другими модулями)

Mod_Utils
  └── Mod_Logger (WriteLog — обёртка обратной совместимости)
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
| `Mod_SheetButtons` | `src/modules/Mod_SheetButtons.bas` | Стандартный модуль |
| `Mod_ModelTypes` | `src/modules/Mod_ModelTypes.bas` | Стандартный модуль |
| `Mod_ModelDB` | `src/modules/Mod_ModelDB.bas` | Стандартный модуль |
| `Mod_PickWork` | `src/modules/Mod_PickWork.bas` | Стандартный модуль |
| `Mod_AutoMatch` | `src/modules/Mod_AutoMatch.bas` | Стандартный модуль |
| `IModelDataProvider` | `src/classes/IModelDataProvider.cls` | Класс |
| `Mod_ModelDBProvider` | `src/classes/Mod_ModelDBProvider.cls` | Класс |
| `Mod_SQLiteDB` | `src/classes/Mod_SQLiteDB.cls` | Класс |
| `PartIdentity` | `src/classes/PartIdentity.cls` | Класс |
| `WorkEntry` | `src/classes/WorkEntry.cls` | Класс |
| `WorkIdentity` | `src/classes/WorkIdentity.cls` | Класс |
| `Лист2` | `src/sheets/Лист2.cls` | Класс листа |

## Приложение C: Скрипты автоматизации

| Скрипт | Назначение | Кодировка |
|--------|-----------|-----------|
| [`export_vba.py`](../scripts/export_vba.py) | Выгрузка VBA из Excel на диск (CP1251 → UTF-8) | UTF-8 |
| [`impVBA.py`](../scripts/impVBA.py) | Загрузка VBA с диска в Excel (UTF-8 → CP1251) | UTF-8 |
| [`check_vba_syntax.py`](../scripts/check_vba_syntax.py) | Статический контроль компиляции VBA (`src/`), этап после impVBA; exit 0/1, код конвейера 22 | UTF-8 |
| [`run_tests.py`](../scripts/run_tests.py) | Запуск тестов VBA | UTF-8 |
| [`build_all.py`](../scripts/build_all.py) | Единый конвейер сборки: бэкап → импорт VBA → шаблоны → миграция БД → контроль целостности → тесты | UTF-8 |
| [`migrate_models_to_sqlite.py`](../scripts/migrate_models_to_sqlite.py) | Пересборка `SysW.db` из `base/models/*` | UTF-8 |
| [`initiate_models.py`](../scripts/initiate_models.py) | Инициация пользовательских модельных файлов: детекция новых групп, валидация структуры, регистрация в `model_groups` и миграция данных в `SysW.db`; отчёт `logs/initiation_report.log` | UTF-8 |
| [`config.py`](../scripts/config.py) | Конфигурация проекта (пути, настройки) | UTF-8 |
| [`config.ps1`](../scripts/config.ps1) | Конфигурация окружения PowerShell | UTF-8 with BOM |
| [`Set-ExcelTrust.ps1`](../scripts/Set-ExcelTrust.ps1) | Настройка доверия Excel для работы VBA-макросов | UTF-8 with BOM |
| [`check_docs.py`](../scripts/check_docs.py) | Проверка актуальности документации | UTF-8 |

---

## Приложение D: Реестр макросов VBA

> **Версия:** v1.0.15 (историческая — реестр отражает фактическое состояние кодовой базы по
> результатам аудита `src/` и `scripts/`). Содержимое перенесено из `docs/reestr.md` без потерь.
> Служит перекрёстной ссылкой с текущим документом, [ARCHITECTURE.md](ARCHITECTURE.md),
> [table.md](table.md) и [README.md](../README.md).

### D.1 Кнопки листа main (обработчики `Mod_ButtonDispatcher`)

| Имя макроса               | Краткое описание функционала                   | Кнопка (если есть) | Модуль         |
| ----------------------------------- | ------------------------------------------------------------------------ | -------------------------------- | -------------------- |
| `Btn_main_Clear_Click`            | Очистка данных main B4:ZZ (с подтверждением) | «ОЧИСТ ВСЁ»            | Mod_ButtonDispatcher |
| `Btn_main_FillHeader_Click`       | Заполнение шапки B5:B17 по № из B4                   | «ЗАПОЛН ШАПКУ»      | Mod_ButtonDispatcher |
| `Btn_main_ClearHeader_Click`      | Очистка шапки B5:B17                                         | «ОЧИСТ ШАПКУ»        | Mod_ButtonDispatcher |
| `Btn_main_RunTests_Click`         | Запуск всех автотестов (TC-01..TC-64+TC-S*)          | «Тесты»                   | Mod_ButtonDispatcher |
| `Btn_main_WriteLog_Click`         | Запись сообщения в лог через InputBox            | «Лог»                       | Mod_ButtonDispatcher |
| `Btn_main_ImportDataToMain_Click` | Перенос данных с активного листа в main     | «Перенести в main»   | Mod_ButtonDispatcher |
| `Btn_main_FindOrder_Click`        | Поиск заказа по № (InputBox) + вывод                  | «Найти заказ»        | Mod_ButtonDispatcher |
| `Btn_main_CheckFileExists_Click`  | Проверка существования файла по пути     | «Проверить файл»  | Mod_ButtonDispatcher |
| `Btn_main_ImportVH_Click`         | Импорт ВХ: лист {B4}M + заполнение шапки      | «Импорт ВХ»            | Mod_ButtonDispatcher |

### D.2 Автоподбор / ручной подбор (`Mod_ButtonDispatcher` → логика)

| Имя макроса             | Краткое описание функционала                                                | Кнопка (если есть) | Модуль         |
| --------------------------------- | ----------------------------------------------------------------------------------------------------- | -------------------------------- | -------------------- |
| `Btn_main_AutoMatchWorks_Click` | Автоподбор работ (L→E:I + формула J)                                           | «АВТО РАБ»              | Mod_ButtonDispatcher |
| `Btn_main_AutoMatchParts_Click` | Автоподбор запчастей (X→Q:U + формула V)                                   | «АВТО ЗЧ»                | Mod_ButtonDispatcher |
| `Btn_main_PickWork_Click`       | Ручной подбор работ: открытие файла группы + инструкция | «РУЧ РАБ»                | Mod_ButtonDispatcher |
| `Btn_main_PickParts_Click`      | Ручной подбор запчастей: открытие `{Group}z4`/`z4` + инструкция | «РУЧ ЗЧ»    | Mod_ButtonDispatcher |

### D.3 Универсальный поиск листов работ и запчастей (`Mod_SheetButtons`)

| Имя макроса                                        | Краткое описание функционала | Кнопка (если есть)             | Модуль     |
| ------------------------------------------------------------ | ------------------------------------------------------ | -------------------------------------------- | ---------------- |
| `Btn_Search_ByArticle_Click` → `Btn_Search_ByArticle`   | Поиск «содержит» по ст. B (артикул) | «Поиск по артикулу»         | Mod_SheetButtons |
| `Btn_Search_ByName_Click` → `Btn_Search_ByName`         | Поиск по ст. C (наименование)          | «Поиск по наименованию» | Mod_SheetButtons |
| `Btn_Search_Clear_Click` → `Btn_ClearFilter`            | Сброс фильтра + очистка поля ввода C1 | «Сброс»                            | Mod_SheetButtons |

> Универсальный поиск (v1.0.12): имя группы читается из `main!$B$14` (fallback — имя листа);
> классификация листов `{Group}`/`{Group}w`/`z4`/`{Group}z4` через `ClassifySheet` (enum `SheetKind`).
> Прежние имена `Btn_UAZ_*`/`Btn_Parts_*` и `ExecuteUAZSearch`/`ExecutePartsSearch`
> переименованы в нейтральные (`Btn_Search_*`/`Btn_ClearFilter`/`ExecuteSearch`).

> Приватные помощники модуля `Mod_SheetButtons` (в состав реестра включаются для полноты):
> `ExecuteSearch(searchColumn)`, `IsSearchableSheet(ws)`, `ResolveGroupName(ws)`,
> `GetGroupName()`. Других обработчиков в модуле нет.

### D.4 Служебная бизнес-логика по модулям

| Имя                                                | Краткое описание функционала                                                | Кнопка (если есть)  | Модуль    |
| ----------------------------------------------------- | ----------------------------------------------------------------------------------------------------- | --------------------------------- | --------------- |
| `FillHeaderFromOrder(orderNum)`                     | Заполнение шапки B5:B17 данными из spisok/models                              | —                                | Mod_OrderHeader |
| `FindOrder(orderNum, Header)`                       | Поиск заказа по №, заполнение структуры`OrderHeader`               | —                                | Mod_OrderHeader |
| `FillHeaderFromOrder_UI`                            | UI-обёртка заполнения шапки                                                     | «Заполнить шапку» | Mod_OrderHeader |
| `FindOrder_UI`                                      | UI-обёртка поиска заказа                                                           | «Найти заказ»         | Mod_OrderHeader |
| `ImportSheet(grz)`                                  | Импорт листа из report.xlsx по ГРЗ                                                  | —                                | Mod_Import      |
| `ImportDataToMain(wsSource)`                        | Перенос данных с листа-источника в main                                  | —                                | Mod_Import      |
| `ImportDataToMain_UI`                               | Перенос данных с активного листа в main                                  | «Перенести в main»    | Mod_Import      |
| `ImportFromB2_UI`                                   | Импорт ВХ: фактически читает B4/{B4}M                                         | «Импорт ВХ»             | Mod_Import      |
| `SubstituteWorkArticle` / `SubstitutePartArticle` | Приватная подстановка тождеств (работы/запчасти)            | —                                | Mod_Import      |
| `AutoMatchWorks()`                                  | Автоподбор работ (L→E:I + формула J)                                           | «АВТО РАБ»               | Mod_AutoMatch   |
| `AutoMatchParts()`                                  | Автоподбор запчастей (X→Q:U + формула V)                                   | «АВТО ЗЧ»                 | Mod_AutoMatch   |
| `PickWork_UI`                                       | Ручной подбор работ: открытие файла группы + инструкция | «РУЧ РАБ»                 | Mod_PickWork    |
| `PickParts_UI`                                      | Ручной подбор запчастей: открытие `{Group}z4`/`z4` + инструкция | «РУЧ ЗЧ»    | Mod_PickWork    |
| `GetGroupNameFromMain`                              | Чтение группы из B14 листа main                                                    | —                                | Mod_PickWork    |
| `GetWorkSheetName(groupName)`                       | Имя листа работ (совпадает с группой)                                   | —                                | Mod_PickWork    |
| `GetPartsSheetName(wb, groupName)`                  | Имя модельного листа запчастей `{Group}z4`; при отсутствии — `z4` | —                 | Mod_PickWork    |
| `ExtractNumberFromGRZ(grz)`                         | Извлечение цифровой группы 3/4 из ГРЗ                                    | —                                | Mod_SheetOps    |
| `SearchSheetByGRZ(grz)`                             | Поиск листа в report.xlsx по № ГРЗ                                                   | —                                | Mod_SheetOps    |
| `RenameSheetsByGRZ`                                 | Переименование листов report.xlsx по ГРЗ                                     | —                                | Mod_SheetOps    |
| `ApplyFreezePanes(ws)`                              | Закрепление строк (FreezePanes A4)                                                    | —                                | Mod_SheetOps    |
| `ClearMainSheet_UI([silent])`                       | Очистка данных main B4:ZZ                                                                | «Очистить»              | Mod_SheetOps    |
| `ClearHeader_UI`                                    | Очистка шапки B5:B17                                                                      | «Очистить шапку»   | Mod_SheetOps    |
| `GetSheetByName(wb, name)`                          | Получение листа по имени без ошибки                                     | —                                | Mod_Utils       |
| `GetWorkbookPath`                                   | Путь к книге                                                                                | —                                | Mod_Utils       |
| `FileExists(filePath)`                              | Проверка существования файла                                                | —                                | Mod_Utils       |
| `GetCurrentUser`                                    | Имя пользователя Windows                                                               | —                                | Mod_Utils       |
| `FormatDateSQL(d)`                                  | Форматирование даты ГГГГ-ММ-ДД                                              | —                                | Mod_Utils       |
| `WriteLog(message)`                                 | Обёртка обратной совместимости для`Mod_Logger`                       | —                                | Mod_Utils       |
| `WriteLog_UI`                                       | Запись сообщения в лог через InputBox                                         | «Лог»                        | Mod_Utils       |
| `ShowWorkbookPath_UI`                               | Показ пути к книге                                                                     | «Путь к книге»        | Mod_Utils       |
| `ShowCurrentUser_UI`                                | Показ имени пользователя Windows                                                | «Пользователь»      | Mod_Utils       |
| `CheckFileExists_UI`                                | Проверка существования файла по пути                                  | «Проверить файл»   | Mod_Utils       |
| `InitLibName`                                       | Заполнение листа libname реестром имён                                     | —                                | Mod_Constants   |
| `AddWorkEntry`                                      | Добавление записи work.xlsm на лист libname                                     | —                                | Mod_Constants   |
| `GetAggregateName(code)`                            | Код агрегата → русское название                                            | —                                | Mod_Constants   |
| `SqliteProviderEnabled`                             | Флаг доступности SQLite-провайдера                                           | —                                | Mod_Constants   |
| `GetLogPath`                                        | Путь к файлу лога                                                                       | —                                | Mod_Logger      |
| `WriteLog(moduleName, message)`                     | Запись сообщения в лог                                                             | —                                | Mod_Logger      |
| `WriteLogE(moduleName, message)`                    | Запись ошибки с префиксом [ERROR]                                               | —                                | Mod_Logger      |
| `RotateLogIfNeeded(maxSizeKB)`                      | Ротация лог-файла при превышении размера                           | —                                | Mod_Logger      |
| `ClearLog`                                          | Очистка файла лога                                                                    | —                                | Mod_Logger      |

**Провайдеры данных (классы):**

| Имя                                                            | Краткое описание функционала                             | Кнопка (если есть) | Модуль           |
| ----------------------------------------------------------------- | ---------------------------------------------------------------------------------- | -------------------------------- | ---------------------- |
| `GetModelDataProvider(ByRef provider)`                          | Фабрика провайдера данных (SQLite / Excel-fallback)         | —                               | Mod_ModelDB            |
| `GetModelDBBasePath`                                            | Функция: базовый путь`base\models\`                            | —                               | Mod_ModelDB            |
| `GetModelGroupFilePath(groupName)`                              | Полный путь к файлу группы                                   | —                               | Mod_ModelDB            |
| `ModelGroupFileExists(groupName)`                               | Проверка существования файла группы                | —                               | Mod_ModelDB            |
| `OpenModelGroupFile(groupName)`                                 | Открытие файла группы                                           | —                               | Mod_ModelDB            |
| `GetWorks/GetParts/GetModelWorks/GetModelParts`                 | Чтение работ/запчастей (модельных)                    | —                               | Mod_ModelDB            |
| `GetMatLibEntries(groupName, entryCode)`                        | Тождества соответствий                                        | —                               | Mod_ModelDB            |
| `GetWorkIdentities/GetPartIdentities`                           | Коллекции тождеств работ/запчастей                  | —                               | Mod_ModelDB            |
| `GetAllModelGroups`                                             | Список всех модельных групп                                | —                               | Mod_ModelDB            |
| `CreateModelGroupFile(groupName)`                               | Создание файла группы                                           | —                               | Mod_ModelDB            |
| `FindModelGroupByModel(modelName)`                              | Группа по названию модели                                    | —                               | Mod_ModelDB            |
| `OpenConnection/CloseConnection/IsConnected`                    | Жизненный цикл соединения SQLite                            | —                               | Mod_SQLiteDB           |
| `ExecuteScalar/ExecuteQuery/ExecuteNonQuery`                    | Выполнение SQL-запросов                                          | —                               | Mod_SQLiteDB           |
| `IModelDataProvider_*`                                          | Реализация интерфейса (чтение из SQLite)               | —                               | Mod_SQLiteDB           |
| `IModelDataProvider_*`                                          | Реализация интерфейса (Excel-fallback)                         | —                               | Mod_ModelDBProvider    |
| `ReadWorkIdentitiesFromSheet` / `ReadPartIdentitiesFromSheet` | Приватное чтение тождеств с листа                     | —                               | Mod_ModelDBProvider    |
| `IModelDataProvider` (интерфейс)                       | Контракт доступа к данным моделей (10 методов) | —                               | IModelDataProvider.cls |
| `PartIdentity` / `WorkIdentity` / `WorkEntry`               | Объекты-контейнеры данных                                   | —                               | *.cls                  |

**Листовые события (`src/sheets/`):**

| Имя                           | Краткое описание функционала               | Кнопка (если есть) | Модуль  |
| -------------------------------- | -------------------------------------------------------------------- | -------------------------------- | ------------- |
| `Лист2.Worksheet_Activate` | FreezePanes A4 при активации листа main             | —                               | Лист2.cls |
| `Лист2.Worksheet_Change`   | Автозаполнение шапки при изменении B4 | —                               | Лист2.cls |
| `Лист3.Worksheet_Activate` | FreezePanes A4 (лист libname)                                    | —                               | Лист3.cls |
| `Лист5.Worksheet_Activate` | FreezePanes A4 (лист models)                                     | —                               | Лист5.cls |
| `Лист9.Worksheet_Activate` | FreezePanes A4 (лист spisok)                                     | —                               | Лист9.cls |

> Фактический состав `src/sheets/`: `Лист2.cls`, `Лист3.cls`, `Лист5.cls`, `Лист9.cls`.
> Устаревшие упоминания `Лист2_main.cls` и `Лист4.cls` в документации не соответствуют
> текущему составу каталога.

### D.5 Заглушки и отсутствующие процедуры

> Включены в реестр для полноты, но явно помечены как заглушки/несуществующие.

| Имя                                           | Статус                                             | Примечание                                                                                                                        |
| ------------------------------------------------ | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `IModelDataProvider_CreateModelGroupFile`      | **ЗАГЛУШКА**                               | Excel-ветка (`Mod_ModelDBProvider`) возвращает `True` без фактического создания файла        |
| `IModelDataProvider.*` (10 методов)     | **ЗАГЛУШКИ-контракт**              | В`IModelDataProvider.cls` все методы бросают `Err.Raise ... "Not implemented"` (контракт интерфейса) |
| `Mod_ModelTypes.bas`                           | **Устаревший пустой модуль** | Хранится для обратной совместимости ссылок; UDT вынесены в классы/др.                |
| `Btn_z4_Action1/2/3`, `Btn_work_Action1/2/3` | **Отсутствуют в коде**             | Упоминались в старой документации; в`Mod_SheetButtons` не реализованы                         |
| `AutoMatch_UI()`                               | **Отсутствует в коде**             | Реальные точки входа —`AutoMatchWorks` / `AutoMatchParts`                                                            |
| `MODELDB_BASE_PATH` (константа)       | **Отсутствует в коде**             | Реализована функция`GetModelDBBasePath`                                                                                 |

---

## Приложение E: Реестр скриптов

> Столбцы: **Скрипт | Назначение | Входные параметры | Выходные артефакты | Зависимости**.

### E.1 Python — ядро конвейера

| Скрипт                    | Назначение                                                                                                                                          | Входные параметры                                                   | Выходные артефакты                               | Зависимости                                                                                                          |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- | ----------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `build_all.py`                | Единый конвейер: бэкап → impVBA → check_vba_syntax → build_templates → migrate → integrity → run_tests; exit-коды 1/2/22/3/4/5/6 | нет                                                                              | `_backup/*`, `logs/build.log`                                 | config, sqlite3, subprocess; вызывает impVBA/check_vba_syntax/build_templates/migrate/run_tests                         |
| `impVBA.py`                   | Импорт VBA из src/ в work.xlsm (UTF-8→CP1251) через COM; удаление компонентов; PID Excel                                    | нет                                                                              | `work.xlsm`, `logs/excel_pid_impvba.txt`                      | config, pywin32, src/{modules,sheets,classes}                                                                                   |
| `export_vba.py`               | Экспорт VBA из work.xlsm в src/                                                                                                                     | нет (все компоненты по маппингу`COMPONENTS`); `--dry` | `src/**`                                                        | pywin32; маппинг`COMPONENTS` (содержит ошибочные `Лист4`, отсутствует `Лист9`) |
| `check_vba_syntax.py`         | Статическая проверка синтаксиса VBA src/                                                                                         | опц. путь (по умолч.`src/`)                                         | отчёт stdout; exit 0/1                                       | stdlib                                                                                                                          |
| `build_templates.py`          | Пересборка base/templates/ + защита листов + FreezePanes A4 для base/models/*.xlsm                                                   | нет                                                                              | `base/templates/*`, изменённые `base/models/*.xlsm` | template_protection, pywin32                                                                                                    |
| `migrate_models_to_sqlite.py` | Конвертация base/models/*.xlsm → SysW.db (WAL); идемпотентно                                                                          | `--force`                                                                         | `SysW.db`, `logs/migration_report.log`, `_backup/SysW_*.db` | config, sqlite_schema, openpyxl                                                                                                 |
| `run_tests.py`                | Запуск макроса RunAllTests через COM, чтение GetTestResults из Z1                                                                   | нет                                                                              | `logs/test_results.log`                                         | config, pywin32, work.xlsm                                                                                                      |

### E.2 Python — библиотеки/служебные

| Скрипт               | Назначение                                                                      | Входные параметры                                                              | Выходные артефакты                                                                        | Зависимости           |
| -------------------------- | ----------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- | -------------------------------- |
| `config.py`              | Единый источник версии и путей (APP_VERSION)                    | импортируется                                                                     | константы                                                                                         | stdlib                           |
| `sqlite_schema.py`       | DDL-схема SysW.db, init_db/set/get_user_version                                      | импортируется; прямое исполнение = демо-создание БД | `SysW.db` (при прямом запуске)                                                           | stdlib sqlite3                   |
| `template_protection.py` | Библиотека защиты листов (Protect/AllowEditRanges/FreezePanes, XML) | импортируется                                                                     | функции                                                                                             | openpyxl, zipfile/re/ElementTree |
| `check_docs.py`          | Проверка согласованности документации                  | `--check` / `--fix` (обязательны)                                               | отчёт; exit 0/1                                                                                       | config                           |
| `update_version.py`      | Автообновление версии SemVer во всех местах + CHANGELOG   | `update_version.py <X.Y.Z>`                                                                  | правит Mod_Constants.bas, config.py, config.ps1, README, DEVELOPER, ROADMAP, ARCHITECTURE, CHANGELOG | stdlib                           |

### E.3 PowerShell

| Скрипт              | Назначение                                                                                                           | Входные параметры | Выходные артефакты | Зависимости         |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------ | --------------------------------- | ----------------------------------- | ------------------------------ |
| `config.ps1`            | PS-конфигурация (версия, пути), dot-source                                                               | dot-source                        | переменные`$Script:*`   | stdlib                         |
| `Set-ExcelTrust.ps1`    | Настройка доверия Excel через реестр (AccessVBOM=1, VBAWarnings=1, доверенная папка) | нет (админ)               | правки HKCU Excel             | config.ps1                     |
| `fix_vbom_and_venv.ps1` | Диагностика/исправление AccessVBOM и .venv                                                              | нет                            | правки реестра         | `.venv/Scripts/Activate.ps1` |

### E.4 Вспомогательные / служебные (пометки)

| Скрипт                         | Статус / примечание                                                                                                        |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `apply_protection_templates.py`    | Вспомогательный; дублирует защиту из`build_templates.py` — кандидат на упразднение |
| `.codeassistant/rules/`            | Каталог рабочих правил рабочей области (8 файлов-зеркал `.ycarules`, коды `[O][G][Z][K][U][S][T][E]`) — см. §10.4 |
| `.github/workflows/docs-check.yml` | CI:`python scripts/check_docs.py --check`                                                                                                |
| `.github/workflows/vba-check.yml`  | CI: проверка VBA-файлов и актуальности`CHANGELOG.md`                                                          |

---

## Приложение F: Реестр тестов (Mod_FullTestRunner)

> Набор автоматических тестов покрывает **TC-01..TC-64** + **TC-S1..TC-S3**
> (включая TC-60..TC-62 из группы `RunGlobalBaseTests` — глобальная база запчастей
> `z4.xlsx` и fallback-поиск; TC-63 — единый хелпер `Mod_Utils.GetGroupName`;
> TC-64 — регистрация группы `CreateModelGroupFile` в `model_groups`, SQLite-ветка).
> Управляется процедурами-группами в [`src/modules/Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)
> и запускается через `RunAllTests()` / `RunAllTests_UI()` либо `python scripts/run_tests.py`.
> Запуск подавляет MsgBox (`Mod_Constants.SilenceMsgBox = True`), результаты пишутся
> в ячейку `Z1` листа main (для чтения COM-клиентом) и дублируются в расширенный
> тестовый лог `logs/test_results.log` (уровни INFO/WARN/ERROR, см. F.3).
> Ожидаемый прогон (v1.0.15): **Total=65, Passed=56, Failed=0, Skipped=9**.

### F.1 Состав групп и сценариев

| Группа (процедура) | Сценарии           | Тестируемый модуль                   |
| --------------------------------- | -------------------------- | ----------------------------------------------------- |
| `RunUtilsTests`                 | TC-01..TC-08               | Mod_Utils                                             |
| `RunLoggerTests`                | TC-09..TC-11               | Mod_Logger                                            |
| `RunUtilsEdgeTests`             | TC-12                      | Mod_Utils (граничные случаи)           |
| `RunLibNameTests`               | TC-13                      | Mod_Constants (libname)                               |
| `RunImportVHTests`              | TC-14                      | Mod_Import (ImportFromB2_UI)                          |
| `RunSheetOpsTests`              | TC-15..TC-18, TC-45        | Mod_SheetOps                                          |
| `RunAggregateNameTests`         | TC-19..TC-21               | Mod_Constants (GetAggregateName)                      |
| `RunModelDBReadTests`           | TC-22..TC-24               | Mod_ModelDB / провайдер (тождества) |
| `RunOrderHeaderTests`           | TC-25..TC-28               | Mod_OrderHeader                                       |
| `RunImportDataTests`            | TC-29, TC-30               | Mod_Import                                            |
| `RunImportB2IntegrationTests`   | TC-65..TC-68               | Mod_Import (сквозной импорт: шапка+работы+запчасти)   |
| `RunModelDBTests`               | TC-31..TC-35, TC-64        | Mod_ModelDB                                           |
| `RunPickWorkTests`              | TC-36..TC-38, TC-63        | Mod_PickWork / Mod_Utils                              |
| `RunAutoMatchTests`             | TC-39..TC-44               | Mod_AutoMatch                                         |
| `RunConstantsTests`             | TC-46                      | Mod_Constants (AddWorkEntry)                          |
| `RunSQLiteTests`                | TC-S1..TC-S3, TC-47..TC-50 | Mod_SQLiteDB / провайдер                     |
| `RunSearchTests`                | TC-51..TC-55               | Mod_SheetButtons / Mod_PickWork              |
| `RunGlobalBaseTests`            | TC-60..TC-62               | Mod_ModelDB (глобальная база з/ч, fallback) |

### F.2 Полный перечень тестов

| ID    | Название                                                                     | Группа          | Модуль    | Тип                       | Статус по умолчанию                                                                                   |
| ----- | ------------------------------------------------------------------------------------ | --------------------- | --------------- | ---------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| TC-01 | FileExists с существующим файлом                                  | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-02 | FileExists с несуществующим файлом                              | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-03 | FormatDateSQL с корректной датой                                     | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-04 | FormatDateSQL с нулевой датой                                           | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-05 | GetSheetByName существующий лист                                     | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-06 | GetSheetByName несуществующий лист                                 | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-07 | WriteLog запись в лог                                                      | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-08 | GetWorkbookPath / GetCurrentUser                                                     | RunUtilsTests         | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-09 | WriteLog запись в лог-файл                                             | RunLoggerTests        | Mod_Logger      | Модульный           | PASS                                                                                                                   |
| TC-10 | RotateLogIfNeeded ротация лога                                            | RunLoggerTests        | Mod_Logger      | Модульный           | PASS                                                                                                                   |
| TC-11 | ClearLog очистка лога                                                     | RunLoggerTests        | Mod_Logger      | Модульный           | PASS                                                                                                                   |
| TC-12 | FormatDateSQL граничные случаи                                        | RunUtilsEdgeTests     | Mod_Utils       | Модульный           | PASS; подпроверка «пустая строка» — SKIP (невалидный аргумент для Date) |
| TC-13 | InitLibName заполнение libname                                             | RunLibNameTests       | Mod_Constants   | Модульный           | PASS                                                                                                                   |
| TC-14 | ImportFromB2_UI с пустым B4                                                   | RunImportVHTests      | Mod_Import      | Интеграционный | PASS (silent)                                                                                                          |
| TC-15 | ExtractNumberFromGRZ 'А123АН77' → '123'                                          | RunSheetOpsTests      | Mod_SheetOps    | Модульный           | PASS                                                                                                                   |
| TC-16 | ExtractNumberFromGRZ 'А12АН34' → ''                                              | RunSheetOpsTests      | Mod_SheetOps    | Модульный           | PASS                                                                                                                   |
| TC-17 | ExtractNumberFromGRZ 'А1234АН77' → '1234'                                        | RunSheetOpsTests      | Mod_SheetOps    | Модульный           | PASS                                                                                                                   |
| TC-18 | ExtractNumberFromGRZ '' → ''                                                        | RunSheetOpsTests      | Mod_SheetOps    | Модульный           | PASS                                                                                                                   |
| TC-19 | GetAggregateName 'DIAG' → 'Диагностика'                                  | RunAggregateNameTests | Mod_Constants   | Модульный           | PASS                                                                                                                   |
| TC-20 | GetAggregateName 'TO' → 'ТО'                                                      | RunAggregateNameTests | Mod_Constants   | Модульный           | PASS                                                                                                                   |
| TC-21 | GetAggregateName 'XXX' → ''                                                         | RunAggregateNameTests | Mod_Constants   | Модульный           | PASS                                                                                                                   |
| TC-22 | GetWorkIdentities UAZ (из SysW.db)                                                 | RunModelDBReadTests   | Mod_ModelDB     | Модульный           | PASS; SKIP, если провайдер/SysW.db недоступен                                                   |
| TC-23 | GetPartIdentities UAZ (из SysW.db)                                                 | RunModelDBReadTests   | Mod_ModelDB     | Модульный           | PASS; SKIP, если провайдер/SysW.db недоступен                                                   |
| TC-24 | GetWorks UAZ (из SysW.db)                                                          | RunModelDBReadTests   | Mod_ModelDB     | Модульный           | PASS; SKIP, если провайдер/SysW.db недоступен                                                   |
| TC-25 | FillHeaderFromOrder существующий заказ                              | RunOrderHeaderTests   | Mod_OrderHeader | Модульный           | PASS; SKIP, если нет данных spisok                                                                        |
| TC-26 | FillHeaderFromOrder несуществующий заказ                          | RunOrderHeaderTests   | Mod_OrderHeader | Модульный           | PASS; SKIP, если листы не найдены                                                                    |
| TC-27 | FindOrder существующий заказ                                        | RunOrderHeaderTests   | Mod_OrderHeader | Модульный           | PASS; SKIP, если нет данных spisok                                                                        |
| TC-28 | FindOrder несуществующий заказ                                    | RunOrderHeaderTests   | Mod_OrderHeader | Модульный           | PASS; SKIP, если листы не найдены                                                                    |
| TC-29 | ImportDataToMain перенос данных                                         | RunImportDataTests    | Mod_Import      | Интеграционный | PASS; SKIP, если лист main не найден                                                                   |
| TC-30 | ImportSheet несуществующий ГРЗ                                      | RunImportDataTests    | Mod_Import      | Интеграционный | PASS; SKIP, если лист main не найден                                                                   |
| TC-31 | GetModelDBBasePath возвращает путь                                     | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-32 | GetModelGroupFilePath формирует путь                                    | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-33 | ModelGroupFileExists существующий файл                               | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-34 | ModelGroupFileExists несуществующий файл                           | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-35 | GetModelDataProvider возвращает провайдера                       | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS; SKIP, если провайдер недоступен                                                           |
| TC-36 | GetGroupNameFromMain читает B14                                                | RunPickWorkTests      | Mod_PickWork    | Модульный           | PASS                                                                                                                   |
| TC-37 | GetWorkSheetName возвращает имя листа                              | RunPickWorkTests      | Mod_PickWork    | Модульный           | PASS                                                                                                                   |
| TC-38 | PickWork_UI вызов без ошибки                                           | RunPickWorkTests      | Mod_PickWork    | Модульный           | PASS                                                                                                                   |
| TC-39 | AutoMatchWorks выполняется без ошибки                            | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | PASS                                                                                                                   |
| TC-40 | AutoMatchParts выполняется без ошибки                            | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | PASS                                                                                                                   |
| TC-41 | HighlightNotFound (Private)                                                          | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | SKIP (Private, недоступен прямому вызову)                                                       |
| TC-42 | ClearHighlight (Private)                                                             | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | SKIP (Private, недоступен прямому вызову)                                                       |
| TC-43 | IsAllFound (Private)                                                                 | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | SKIP (Private, недоступен прямому вызову)                                                       |
| TC-44 | AutoMatchWorks без изменения данных                                | RunAutoMatchTests     | Mod_AutoMatch   | Модульный           | SKIP (небезопасно автоматизировать)                                                         |
| TC-45 | SearchSheetByGRZ несуществующий ГРЗ → Nothing                      | RunSheetOpsTests      | Mod_SheetOps    | Модульный           | PASS                                                                                                                   |
| TC-46 | AddWorkEntry добавление work.xlsm                                          | RunConstantsTests     | Mod_Constants   | Модульный           | PASS; SKIP, если лист libname не найден                                                                |
| TC-47 | GetParts через SQLite (JOIN parts_catalog)                                      | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если SQLite-провайдер недоступен                                                    |
| TC-48 | works с дублями наименований без схлопывания       | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если нет групп с дублями                                                               |
| TC-49 | Чтение parts_catalog                                                           | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если SQLite-провайдер недоступен                                                    |
| TC-50 | GetMatLibEntries: детерминированный порядок                  | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если нет записей matlib                                                                      |
| TC-51 | ClassifySheet классификация листов (`{Group}`/`w`/`z4`/`{Group}z4`) | RunSearchTests | Mod_SheetButtons | Модульный | PASS                                                                      |
| TC-52 | Btn_Search_ByArticle без ошибки                                     | RunSearchTests        | Mod_SheetButtons | Модульный | PASS                                                                      |
| TC-53 | Btn_Search_ByName без ошибки                                        | RunSearchTests        | Mod_SheetButtons | Модульный | PASS                                                                      |
| TC-54 | Btn_ClearFilter (сброс фильтра + очистка C1)                        | RunSearchTests        | Mod_SheetButtons | Модульный | PASS                                                                      |
| TC-55 | PickParts_UI вызов без ошибки                                       | RunSearchTests        | Mod_PickWork     | Модульный | PASS                                                                      |
| TC-S1 | Фабрика GetModelDataProvider                                                  | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если SysW.db недоступен                                                                      |
| TC-S2 | GetWorks через провайдер (SQLite эквивалент Excel)           | RunSQLiteTests        | Mod_SQLiteDB    | Модульный           | PASS; SKIP, если провайдер недоступен                                                           |
| TC-S3 | Данные мигрированы в SysW.db (контрольные объёмы) | RunSQLiteTests        | Mod_SQLiteDB    | Интеграционный | PASS; SKIP, если провайдер недоступен                                                           |
| TC-60 | GetGlobalPartsBasePath возвращает путь к `z4.xlsx`             | RunGlobalBaseTests    | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-61 | ReadGlobalPartByKey отсутствие совпадения                        | RunGlobalBaseTests    | Mod_ModelDB     | Модульный           | PASS; SKIP, если файл `z4.xlsx` недоступен                                                                |
| TC-62 | ReadLocalWorkByName отсутствие группы                              | RunGlobalBaseTests    | Mod_ModelDB     | Модульный           | PASS                                                                                                                   |
| TC-63 | GetGroupName единый хелпер чтения B14                     | RunPickWorkTests      | Mod_Utils       | Модульный           | PASS                                                                                                                   |
| TC-64 | CreateModelGroupFile регистрация в model_groups (SQLite) | RunModelDBTests       | Mod_ModelDB     | Модульный           | PASS; SKIP, если SQLite-провайдер недоступен                                                     |
| TC-65 | Импорт B2: работы в L:N (сквозной ImportFromB2_UI)          | RunImportB2IntegrationTests | Mod_Import | Интеграционный | PASS; SKIP, если лист main/spisok/models не найден |
| TC-66 | Импорт B2: запчасти в X:AA (сквозной ImportFromB2_UI)       | RunImportB2IntegrationTests | Mod_Import | Интеграционный | PASS; SKIP, если лист main/spisok/models не найден |
| TC-67 | Импорт B2: шапка B5:B17 (FillHeaderFromOrder)                | RunImportB2IntegrationTests | Mod_Import/Mod_OrderHeader | Интеграционный | PASS; SKIP, если лист main/spisok/models не найден |
| TC-68 | Импорт B2: без исключений/зависаний                          | RunImportB2IntegrationTests | Mod_Import | Интеграционный | PASS; SKIP, если лист main/spisok/models не найден |

**Легенда статусов:** PASS — тест проходит при наличии данных/окружения; SKIP — тест пропускается по условию (Private-процедура, небезопасность автотеста, отсутствие листа/данных/SQLite-провайдера); FAIL — падение теста (останавливает конвейер `run_tests.py` с кодом 1).

### F.3 Механизм записи результатов («2 лога», Задача 1, v1.1.0)

Логирование разделено на два независимых лога (см. [`src/modules/Mod_Logger.bas`](../src/modules/Mod_Logger.bas)):

- **Системный лог** — `logs/log.txt`: общие системные события (текущее поведение,
  методы `WriteLog` / `WriteLogE`, ротация через `RotateLogIfNeeded`, очистка `ClearLog`).
- **Расширенный тестовый лог** — `logs/test_results.log`: детальный лог тестов с уровнями
  **INFO/WARN/ERROR**, включая ошибки VBA. Пишется методом `Mod_Logger.WriteTestLog`
  (обёртки `LogTestInfo` / `LogTestWarn` / `LogTestError`).

Сопоставление статусов теста и уровней тестового лога:

| Статус теста | Уровень в `test_results.log` |
| ------------ | ---------------------------- |
| PASS         | `INFO`                       |
| SKIP         | `WARN`                       |
| FAIL (в т.ч. ошибки VBA) | `ERROR`      |

Механика:

- `RunAllTests()` — запускает все группы, подавляет MsgBox; перед стартом очищает тестовый лог
  (`Mod_Logger.ClearTestLog`) и пишет маркеры `START` / `END` (INFO). Вызывает `WriteResultsToSheet`.
- `AddResult()` — помимо статистики и Immediate Window, пишет каждую строку результата
  в `logs/test_results.log` через `Mod_Logger.WriteTestLog` (PASS→INFO, SKIP→WARN, FAIL→ERROR).
- `WriteResultsToSheet()` — приватная процедура записи отчёта (`Total/Passed/Failed/Skipped`) в ячейку `Z1` листа main.
- `GetTestResults()` — **Public Sub**, дублирует запись отчёта в `Z1` (вызывается из Python-клиента). Не является функцией.
- `scripts/run_tests.py` — дополнительно дописывает в тот же `logs/test_results.log` служебные
  строки прогона (этапы, итог) и читает итог из `Z1` листа main; константа пути —
  `TEST_LOG_FILE` из [`scripts/config.py`](../scripts/config.py).

---

## Приложение G: Ошибки, расхождения, актуальность, рекомендации

> Раздел фиксирует результаты аудита кодовой базы. Здесь собраны фактические расхождения
> между кодом и документацией, статус актуальности и рекомендации. Изменения бизнес-логики
> в рамках этой задачи **не вносятся** — только фиксация.

### G.1 Макросы (VBA)

| Объект                                     | Тип замечания                          | Описание                                                                                                                                                | Актуальность / Рекомендация                                                                                                              |
| ------------------------------------------------ | -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Btn_main_ImportVH_Click`                      | Расхождение в комментарии   | Комментарий в`Mod_ButtonDispatcher.bas` ссылался на «{B2}M», фактически `ImportFromB2_UI` читает `B4`/`{B4}M` | Исправлено в v1.0.15: комментарий актуализирован на `{B4}M` |
| `Btn_z4_Action1/2/3`, `Btn_work_Action1/2/3` | Отсутствуют в коде                 | Упоминались в старой документации (`DEVELOPER.md §2.10`), удалены из `Mod_SheetButtons` (v0.16)                     | Актуальны нейтральные поисковые обработчики `Btn_Search_*`/`Btn_ClearFilter` (v1.0.12); из доков исключены |
| `AutoMatch_UI()`                               | Отсутствует в коде                 | Упоминалась в`DEVELOPER.md §2.13`; реальные точки входа — `AutoMatchWorks` / `AutoMatchParts`                             | Актуализировать документацию                                                                                                          |
| `MODELDB_BASE_PATH` (константа)       | Отсутствует в коде                 | В доке указана как константа; в коде реализована функция`GetModelDBBasePath`                                   | Актуализировать`DEVELOPER.md §2.11` и `ARCHITECTURE.md §3.2`                                                                               |
| `IModelDataProvider_CreateModelGroupFile`      | ЗАГЛУШКА                                   | Excel-ветка (`Mod_ModelDBProvider`) возвращает `True` без фактического создания файла                            | Зафиксировано; реализовать создание или явно пометить заглушкой                                          |
| `IModelDataProvider.*` (10 методов)     | ЗАГЛУШКИ-контракт                  | Все методы бросают`Err.Raise ... "Not implemented"`                                                                                           | Контракт интерфейса; допустимо, реализации —`Mod_SQLiteDB`/`Mod_ModelDBProvider`                                       |
| `Mod_ModelTypes.bas`                           | Устаревший пустой модуль     | Хранится для обратной совместимости ссылок; UDT вынесены                                                          | Кандидат на упразднение после полного перехода на классы                                                        |
| `RunAllTests_UI`                               | Расхождение в комментарии   | Комментарий «TC-01..TC-46» устарел; фактически набор TC-01..TC-50 + TC-S*                                                    | Поправить комментарий при следующей правке                                                                                 |
| `GetTestResults()`                             | Расхождение в документации | В`DEVELOPER.md §6.4` описан как функция; фактически **Public Sub** (запись в Z1)                                     | Актуализировать описание                                                                                                                  |

### G.2 Тесты (Mod_FullTestRunner)

| Объект                           | Тип замечания                          | Описание                                                                                                                                                                                                                              | Актуальность / Рекомендация                                                                                                             |
| -------------------------------------- | -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| TC-41..TC-44                           | SKIP по дизайну                           | Private-процедуры (`HighlightNotFound`, `ClearHighlight`, `IsAllFound`) недоступны прямому вызову; TC-44 небезопасно автоматизировать (меняет данные листа) | Покрытие этих сценариев возможно только интеграционно; рекомендуются ручные проверки |
| TC-12                                  | Частичный SKIP                            | Подпроверка «пустая строка» пропущена (невалидный аргумент для типа`Date`)                                                                                                       | Поведение корректное; замечание к полноте покрытия                                                                  |
| TC-22..24, TC-35, TC-S1..S3, TC-47..50 | Зависимость от окружения     | Пропускаются (SKIP), если SQLite-провайдер/`SysW.db` недоступен или нет данных                                                                                                               | При CI без SQLite покрытие неполное; рекомендуется наличие`SysW.db` в прогоне                               |
| TC-25..30, TC-46                       | Зависимость от данных           | Пропускаются, если отсутствуют листы/данные (`spisok`, `main`, `libname`)                                                                                                                         | Требуются контрольные данные для полного прогона                                                                     |
| Состав набора              | Расхождение в документации | `DEVELOPER.md §6.1` перечисляет TC-01..TC-44; фактически TC-01..TC-50 + TC-S1..S3                                                                                                                                     | Актуализировать перечень и таблицу покрытия                                                                              |

### G.3 Скрипты (Python / PowerShell)

| Объект                                                 | Тип замечания                                | Описание                                                                                                                                                                                                | Актуальность / Рекомендация                                                                                             |
| ------------------------------------------------------------ | -------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `scripts/export_vba.py`                                    | Расхождение в маппинге`COMPONENTS` | Содержит ошибочные`"Лист4": sheets/Лист4.cls` (файла нет); **отсутствует** `"Лист9": sheets/Лист9.cls` (фактический лист)          | Исправить маппинг`COMPONENTS` (актуальный состав `src/sheets/`: Лист2, Лист3, Лист5, Лист9) |
| `scripts/apply_protection_templates.py`                    | Дублирование                                 | Дублирует защиту листов из`build_templates.py`                                                                                                                                         | Кандидат на упразднение (объединить в`build_templates.py`)                                                    |
| `.codeassistant/mcp.json`                                  | Устарел (удалён)                               | Файл`mcp.json` удалён; заменён корневым`.sourcecraft`; `.codeassistant/` теперь содержит каталог `rules/`                                                                                            | Устранить упоминание при следующей правке; используется`.codeassistant/rules/`               |
| `docs/sourcecraft-guide.md` (список scripts/)        | Устранено (v1.1.6.2)                         | Перечень скриптов в`sourcecraft-guide.md` дополнен до полного состава (`apply_protection_templates`, `apply_sheet_format`, `build_global_parts`, `check_z4_fallback`, `clean_system`, `initiate_models`, бизнес-прогоны, `monitor_long.ps1` и др.) | Закрыто в v1.1.6.2                                                                                              |
| `README.md` / `DEVELOPER.md` (состав листов) | Расхождение                                   | Упоминаются`Лист2_main.cls`, `Лист4.cls`; фактически `Лист2.cls`, `Лист3.cls`, `Лист5.cls`, `Лист9.cls`                                                | Актуализировать упоминания состава листов                                                                 |
| `docs/ARCHITECTURE.md §4.2`                               | Расхождение диапазона                | Указано «B3:B15»; фактически шапка`B5:B17` (согласовано с `DEVELOPER.md §2.2`)                                                                                         | Исправить на`B5:B17`                                                                                                               |