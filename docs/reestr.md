# Реестр макросов и скриптов системы SysW

> **Версия:** v1.0.14
>
> **Назначение:** единый реестр процедур/функций VBA основной бизнес-логики и скриптов
> автоматизации (Python, PowerShell). Документ отражает фактическое состояние кодовой
> базы по результатам аудита `src/` и `scripts/`. Служит перекрёстной ссылкой для
> [`docs/DEVELOPER.md`](DEVELOPER.md), [`docs/ARCHITECTURE.md`](ARCHITECTURE.md),
> [`docs/table.md`](table.md) и [`README.md`](../README.md).

---

## 1. Реестр макросов основной бизнес-логики

### 1.1 Кнопки листа main (обработчики `Mod_ButtonDispatcher`)

| Имя макроса               | Краткое описание функционала                   | Кнопка (если есть) | Модуль         |
| ----------------------------------- | ------------------------------------------------------------------------ | -------------------------------- | -------------------- |
| `Btn_main_Clear_Click`            | Очистка данных main B4:ZZ (с подтверждением) | «ОЧИСТ ВСЁ»            | Mod_ButtonDispatcher |
| `Btn_main_FillHeader_Click`       | Заполнение шапки B5:B17 по № из B4                   | «ЗАПОЛН ШАПКУ»      | Mod_ButtonDispatcher |
| `Btn_main_ClearHeader_Click`      | Очистка шапки B5:B17                                         | «ОЧИСТ ШАПКУ»        | Mod_ButtonDispatcher |
| `Btn_main_RunTests_Click`         | Запуск всех автотестов (TC-01..TC-62+TC-S*)          | «Тесты»                   | Mod_ButtonDispatcher |
| `Btn_main_WriteLog_Click`         | Запись сообщения в лог через InputBox            | «Лог»                       | Mod_ButtonDispatcher |
| `Btn_main_ImportDataToMain_Click` | Перенос данных с активного листа в main     | «Перенести в main»   | Mod_ButtonDispatcher |
| `Btn_main_FindOrder_Click`        | Поиск заказа по № (InputBox) + вывод                  | «Найти заказ»        | Mod_ButtonDispatcher |
| `Btn_main_CheckFileExists_Click`  | Проверка существования файла по пути     | «Проверить файл»  | Mod_ButtonDispatcher |
| `Btn_main_ImportVH_Click`         | Импорт ВХ: лист {B4}M + заполнение шапки      | «Импорт ВХ»            | Mod_ButtonDispatcher |

### 1.2 Автоподбор / ручной подбор (`Mod_ButtonDispatcher` → логика)

| Имя макроса             | Краткое описание функционала                                                | Кнопка (если есть) | Модуль         |
| --------------------------------- | ----------------------------------------------------------------------------------------------------- | -------------------------------- | -------------------- |
| `Btn_main_AutoMatchWorks_Click` | Автоподбор работ (L→E:I + формула J)                                           | «АВТО РАБ»              | Mod_ButtonDispatcher |
| `Btn_main_AutoMatchParts_Click` | Автоподбор запчастей (X→Q:U + формула V)                                   | «АВТО ЗЧ»                | Mod_ButtonDispatcher |
| `Btn_main_PickWork_Click`       | Ручной подбор работ: открытие файла группы + инструкция | «РУЧ РАБ»                | Mod_ButtonDispatcher |
| `Btn_main_PickParts_Click`      | Ручной подбор запчастей: открытие `{Group}z4`/`z4` + инструкция | «РУЧ ЗЧ»    | Mod_ButtonDispatcher |

### 1.3 Универсальный поиск листов работ и запчастей (`Mod_SheetButtons`)

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

### 1.4 Служебная бизнес-логика по модулям

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

### 1.5 Заглушки и отсутствующие процедуры

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

## 2. Реестр скриптов системы

> Столбцы: **Скрипт | Назначение | Входные параметры | Выходные артефакты | Зависимости**.

### 2.1 Python — ядро конвейера

| Скрипт                    | Назначение                                                                                                                                          | Входные параметры                                                   | Выходные артефакты                               | Зависимости                                                                                                          |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- | ----------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `build_all.py`                | Единый конвейер: бэкап → impVBA → check_vba_syntax → build_templates → migrate → integrity → run_tests; exit-коды 1/2/22/3/4/5/6 | нет                                                                              | `_backup/*`, `logs/build.log`                                 | config, sqlite3, subprocess; вызывает impVBA/check_vba_syntax/build_templates/migrate/run_tests                         |
| `impVBA.py`                   | Импорт VBA из src/ в work.xlsm (UTF-8→CP1251) через COM; удаление компонентов; PID Excel                                    | нет                                                                              | `work.xlsm`, `logs/excel_pid_impvba.txt`                      | config, pywin32, src/{modules,sheets,classes}                                                                                   |
| `export_vba.py`               | Экспорт VBA из work.xlsm в src/                                                                                                                     | нет (все компоненты по маппингу`COMPONENTS`); `--dry` | `src/**`                                                        | pywin32; маппинг`COMPONENTS` (содержит ошибочные `Лист4`, отсутствует `Лист9`) |
| `check_vba_syntax.py`         | Статическая проверка синтаксиса VBA src/                                                                                         | опц. путь (по умолч.`src/`)                                         | отчёт stdout; exit 0/1                                       | stdlib                                                                                                                          |
| `build_templates.py`          | Пересборка base/templates/ + защита листов + FreezePanes A4 для base/models/*.xlsm                                                   | нет                                                                              | `base/templates/*`, изменённые `base/models/*.xlsm` | template_protection, pywin32                                                                                                    |
| `migrate_models_to_sqlite.py` | Конвертация base/models/*.xlsm → SysW.db (WAL); идемпотентно                                                                          | `--force`                                                                         | `SysW.db`, `logs/migration_report.log`, `_backup/SysW_*.db` | config, sqlite_schema, openpyxl                                                                                                 |
| `run_tests.py`                | Запуск макроса RunAllTests через COM, чтение GetTestResults из Z1                                                                   | нет                                                                              | `logs/test_results.log`                                         | config, pywin32, work.xlsm                                                                                                      |

### 2.2 Python — библиотеки/служебные

| Скрипт               | Назначение                                                                      | Входные параметры                                                              | Выходные артефакты                                                                        | Зависимости           |
| -------------------------- | ----------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- | -------------------------------- |
| `config.py`              | Единый источник версии и путей (APP_VERSION)                    | импортируется                                                                     | константы                                                                                         | stdlib                           |
| `sqlite_schema.py`       | DDL-схема SysW.db, init_db/set/get_user_version                                      | импортируется; прямое исполнение = демо-создание БД | `SysW.db` (при прямом запуске)                                                           | stdlib sqlite3                   |
| `template_protection.py` | Библиотека защиты листов (Protect/AllowEditRanges/FreezePanes, XML) | импортируется                                                                     | функции                                                                                             | openpyxl, zipfile/re/ElementTree |
| `check_docs.py`          | Проверка согласованности документации                  | `--check` / `--fix` (обязательны)                                               | отчёт; exit 0/1                                                                                       | config                           |
| `update_version.py`      | Автообновление версии SemVer во всех местах + CHANGELOG   | `update_version.py <X.Y.Z>`                                                                  | правит Mod_Constants.bas, config.py, config.ps1, README, DEVELOPER, ROADMAP, ARCHITECTURE, CHANGELOG | stdlib                           |

### 2.3 PowerShell

| Скрипт              | Назначение                                                                                                           | Входные параметры | Выходные артефакты | Зависимости         |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------------ | --------------------------------- | ----------------------------------- | ------------------------------ |
| `config.ps1`            | PS-конфигурация (версия, пути), dot-source                                                               | dot-source                        | переменные`$Script:*`   | stdlib                         |
| `Set-ExcelTrust.ps1`    | Настройка доверия Excel через реестр (AccessVBOM=1, VBAWarnings=1, доверенная папка) | нет (админ)               | правки HKCU Excel             | config.ps1                     |
| `fix_vbom_and_venv.ps1` | Диагностика/исправление AccessVBOM и .venv                                                              | нет                            | правки реестра         | `.venv/Scripts/Activate.ps1` |

### 2.4 Вспомогательные / служебные (пометки)

| Скрипт                         | Статус / примечание                                                                                                        |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `apply_protection_templates.py`    | Вспомогательный; дублирует защиту из`build_templates.py` — кандидат на упразднение |
| `.codeassistant/mcp.json`          | Устарел; заменён на`.sourcecraft`                                                                                        |
| `.github/workflows/docs-check.yml` | CI:`python scripts/check_docs.py --check`                                                                                                |
| `.github/workflows/vba-check.yml`  | CI: проверка VBA-файлов и актуальности`CHANGELOG.md`                                                          |

---

## 3. Реестр тестов (Mod_FullTestRunner)

> Набор автоматических тестов покрывает **TC-01..TC-62** + **TC-S1..TC-S3**
> (включая новые TC-60..TC-62 из группы `RunGlobalBaseTests` — глобальная база
> запчастей `z4.xlsx` и fallback-поиск `ReadGlobalPartByKey`/`ReadLocalWorkByName`).
> Управляется процедурами-группами в [`src/modules/Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)
> и запускается через `RunAllTests()` / `RunAllTests_UI()` либо `python scripts/run_tests.py`.
> Запуск подавляет MsgBox (`Mod_Constants.SilenceMsgBox = True`), результаты пишутся
> в ячейку `Z1` листа main (для чтения COM-клиентом).

### 3.1 Состав групп и сценариев

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
| `RunModelDBTests`               | TC-31..TC-35               | Mod_ModelDB                                           |
| `RunPickWorkTests`              | TC-36..TC-38               | Mod_PickWork                                          |
| `RunAutoMatchTests`             | TC-39..TC-44               | Mod_AutoMatch                                         |
| `RunConstantsTests`             | TC-46                      | Mod_Constants (AddWorkEntry)                          |
| `RunSQLiteTests`                | TC-S1..TC-S3, TC-47..TC-50 | Mod_SQLiteDB / провайдер                     |
| `RunSearchTests`                | TC-51..TC-55               | Mod_SheetButtons / Mod_PickWork              |
| `RunGlobalBaseTests`            | TC-60..TC-62               | Mod_ModelDB (глобальная база з/ч, fallback) |

### 3.2 Полный перечень тестов

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

**Легенда статусов:** PASS — тест проходит при наличии данных/окружения; SKIP — тест пропускается по условию (Private-процедура, небезопасность автотеста, отсутствие листа/данных/SQLite-провайдера); FAIL — падение теста (останавливает конвейер `run_tests.py` с кодом 1).

### 3.3 Механизм записи результатов

- `RunAllTests()` — запускает все группы, подавляет MsgBox, вызывает `WriteResultsToSheet`.
- `WriteResultsToSheet()` — приватная процедура записи отчёта (`Total/Passed/Failed/Skipped`) в ячейку `Z1` листа main.
- `GetTestResults()` — **Public Sub**, дублирует запись отчёта в `Z1` (вызывается из Python-клиента). Не является функцией.

---

## 4. Ошибки, расхождения, актуальность, рекомендации

> Раздел фиксирует результаты аудита кодовой базы. Здесь собраны фактические расхождения
> между кодом и документацией, статус актуальности и рекомендации. Изменения бизнес-логики
> в рамках этой задачи **не вносятся** — только фиксация.

### 4.1 Макросы (VBA)

| Объект                                     | Тип замечания                          | Описание                                                                                                                                                | Актуальность / Рекомендация                                                                                                              |
| ------------------------------------------------ | -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Btn_main_ImportVH_Click`                      | Расхождение в комментарии   | Комментарий в`Mod_ButtonDispatcher.bas` ссылается на «{B2}M», фактически `ImportFromB2_UI` читает `B4`/`{B4}M` | Комментарий устарел; код не меняем — поправить комментарий при следующей правке модуля |
| `Btn_z4_Action1/2/3`, `Btn_work_Action1/2/3` | Отсутствуют в коде                 | Упоминались в старой документации (`DEVELOPER.md §2.10`), удалены из `Mod_SheetButtons` (v0.16)                     | Актуальны только поисковые обработчики UAZ/Parts; исключить из доков как реальные                  |
| `AutoMatch_UI()`                               | Отсутствует в коде                 | Упоминалась в`DEVELOPER.md §2.13`; реальные точки входа — `AutoMatchWorks` / `AutoMatchParts`                             | Актуализировать документацию                                                                                                          |
| `MODELDB_BASE_PATH` (константа)       | Отсутствует в коде                 | В доке указана как константа; в коде реализована функция`GetModelDBBasePath`                                   | Актуализировать`DEVELOPER.md §2.11` и `ARCHITECTURE.md §3.2`                                                                               |
| `IModelDataProvider_CreateModelGroupFile`      | ЗАГЛУШКА                                   | Excel-ветка (`Mod_ModelDBProvider`) возвращает `True` без фактического создания файла                            | Зафиксировано; реализовать создание или явно пометить заглушкой                                          |
| `IModelDataProvider.*` (10 методов)     | ЗАГЛУШКИ-контракт                  | Все методы бросают`Err.Raise ... "Not implemented"`                                                                                           | Контракт интерфейса; допустимо, реализации —`Mod_SQLiteDB`/`Mod_ModelDBProvider`                                       |
| `Mod_ModelTypes.bas`                           | Устаревший пустой модуль     | Хранится для обратной совместимости ссылок; UDT вынесены                                                          | Кандидат на упразднение после полного перехода на классы                                                        |
| `RunAllTests_UI`                               | Расхождение в комментарии   | Комментарий «TC-01..TC-46» устарел; фактически набор TC-01..TC-50 + TC-S*                                                    | Поправить комментарий при следующей правке                                                                                 |
| `GetTestResults()`                             | Расхождение в документации | В`DEVELOPER.md §6.4` описан как функция; фактически **Public Sub** (запись в Z1)                                     | Актуализировать описание                                                                                                                  |

### 4.2 Тесты (Mod_FullTestRunner)

| Объект                           | Тип замечания                          | Описание                                                                                                                                                                                                                              | Актуальность / Рекомендация                                                                                                             |
| -------------------------------------- | -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| TC-41..TC-44                           | SKIP по дизайну                           | Private-процедуры (`HighlightNotFound`, `ClearHighlight`, `IsAllFound`) недоступны прямому вызову; TC-44 небезопасно автоматизировать (меняет данные листа) | Покрытие этих сценариев возможно только интеграционно; рекомендуются ручные проверки |
| TC-12                                  | Частичный SKIP                            | Подпроверка «пустая строка» пропущена (невалидный аргумент для типа`Date`)                                                                                                       | Поведение корректное; замечание к полноте покрытия                                                                  |
| TC-22..24, TC-35, TC-S1..S3, TC-47..50 | Зависимость от окружения     | Пропускаются (SKIP), если SQLite-провайдер/`SysW.db` недоступен или нет данных                                                                                                               | При CI без SQLite покрытие неполное; рекомендуется наличие`SysW.db` в прогоне                               |
| TC-25..30, TC-46                       | Зависимость от данных           | Пропускаются, если отсутствуют листы/данные (`spisok`, `main`, `libname`)                                                                                                                         | Требуются контрольные данные для полного прогона                                                                     |
| Состав набора              | Расхождение в документации | `DEVELOPER.md §6.1` перечисляет TC-01..TC-44; фактически TC-01..TC-50 + TC-S1..S3                                                                                                                                     | Актуализировать перечень и таблицу покрытия                                                                              |

### 4.3 Скрипты (Python / PowerShell)

| Объект                                                 | Тип замечания                                | Описание                                                                                                                                                                                                | Актуальность / Рекомендация                                                                                             |
| ------------------------------------------------------------ | -------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `scripts/export_vba.py`                                    | Расхождение в маппинге`COMPONENTS` | Содержит ошибочные`"Лист4": sheets/Лист4.cls` (файла нет); **отсутствует** `"Лист9": sheets/Лист9.cls` (фактический лист)          | Исправить маппинг`COMPONENTS` (актуальный состав `src/sheets/`: Лист2, Лист3, Лист5, Лист9) |
| `scripts/apply_protection_templates.py`                    | Дублирование                                 | Дублирует защиту листов из`build_templates.py`                                                                                                                                         | Кандидат на упразднение (объединить в`build_templates.py`)                                                    |
| `.codeassistant/mcp.json`                                  | Устарел                                           | Заменён на корневой`.sourcecraft`                                                                                                                                                            | Оставить только для обратной совместимости; не использовать                                 |
| `docs/sourcecraft-guide.md` (список scripts/)        | Неполный список                            | Не перечислены`build_templates.py`, `migrate_models_to_sqlite.py`, `check_vba_syntax.py`, `update_version.py`, `fix_vbom_and_venv.ps1`, `template_protection.py`, `sqlite_schema.py` | Актуализировать перечень скриптов                                                                                |
| `README.md` / `DEVELOPER.md` (состав листов) | Расхождение                                   | Упоминаются`Лист2_main.cls`, `Лист4.cls`; фактически `Лист2.cls`, `Лист3.cls`, `Лист5.cls`, `Лист9.cls`                                                | Актуализировать упоминания состава листов                                                                 |
| `docs/ARCHITECTURE.md §4.2`                               | Расхождение диапазона                | Указано «B3:B15»; фактически шапка`B5:B17` (согласовано с `DEVELOPER.md §2.2`)                                                                                         | Исправить на`B5:B17`                                                                                                               |

---

## Связанные документы

- [`docs/DEVELOPER.md`](DEVELOPER.md) — техническая документация разработчика
- [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) — архитектура выноса данных в SQLite
- [`docs/table.md`](table.md) — маппинг столбцов и структура листов
- [`README.md`](../README.md) — общее описание проекта
