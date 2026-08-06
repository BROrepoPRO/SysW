# История изменений

Все заметные изменения в проекте будут документироваться в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/),
версионирование следует [Semantic Versioning](https://semver.org/lang/ru/).

## [v1.0.2] — 2026-08-06

### Fixed
- **Активация виртуального окружения Python в терминале SourceCraft:** в `.vscode/settings.json` путь к активатору `.venv\Scripts\activate.bat` заменён на относительный через переменную `${workspaceFolder}`. Теперь терминал корректно активирует `.venv` независимо от фактического расположения проекта на диске.

### Removed
- **Синхронизация через Syncthing:** удалены `scripts/sync_fix.ps1` и корневой `.stignore`. Из документации (`docs/CHANGELOG.md`, `docs/ARCHITECTURE.md`, `docs/DEVELOPER.md`, `docs/ROADMAP.md`, `docs/git-workflow.md`, `README.md`) убраны все упоминания Syncthing, `sync_fix`, `SysW-syncthing`, `.stfolder`, `.stignore` и файловой синхронизации через Syncthing. Из `scripts/check_docs.py` удалена запись о `sync_fix.ps1`.

### Added
- **Интеграция MCP-серверов File System и Git Tools** (пакеты `@modelcontextprotocol/server-filesystem` и `@modelcontextprotocol/server-git`) в `.sourcecraft` — доступ к файловой системе и Git-репозиторию проекта через `${workspaceFolder}`. В `docs/sourcecraft-guide.md` добавлен соответствующий раздел документации.

## [v1.0.0] — 2026-08-06

### Fixed
- **`docs/ARCHITECTURE.md`:** исправлен раздел 1.3 «Структура файла модельной группы» — удалён несуществующий лист `matlib{GroupName}`, структура файла `base/models/{GroupName}.xlsm` приведена к фактическим 4 листам (`z4`, `{GroupName}z4`, `{GroupName}`, `{GroupName}w`). Соответствия хранятся внутри листов `{GroupName}z4` и `{GroupName}w`.
- **`docs/ARCHITECTURE.md`:** устранены все упоминания несуществующего листа `matlib{GroupName}` (разделы 1.4, 2.5, 3.2, 3.3, 5.2, 5.4). Раздел 2.5 переработан в описание хранения соответствий внутри листов `{GroupName}z4`/`{GroupName}w`; функция `GetMatLibEntries` переработана на чтение соответствий из этих листов; из SQLite-схемы удалена таблица `matlib`.

### Changed
- **Версия проекта обновлена до v1.0.0** — `README.md`, `docs/DEVELOPER.md`, `docs/ROADMAP.md`, `docs/ARCHITECTURE.md`, `scripts/check_docs.py`, `plans/update_docs_plan.md`.

## [Unreleased]

### Removed
- **Дублирующие обработчики кнопок из `Mod_MainButtons.bas`:** удалены `Btn_main_Import`, `Btn_main_AUTOz4`, `Btn_main_AUTOw`, `Btn_main_MANWRK`, `Btn_main_ImportVH_Click` (обёртки-делегирования, дублирующие канонические обработчики в `Mod_ButtonDispatcher.bas`)
- **Модуль `Mod_MainButtons.bas`:** удалён целиком (после удаления всех процедур модуль стал пустым и не вызывался нигде в кодовой базе)
- **`scripts/export_vba.py`:** запись `Mod_MainButtons` удалена из `COMPONENTS`

### Added
- **`scripts/check_docs.py`** — инструмент проверки актуальности документации: проверяет версию v1.0.0, отсутствие упоминаний удалённого `Mod_MainButtons`, наличие новых элементов (`Mod_ModelTypes`, `PartIdentity`, `WorkIdentity`), корректность путей к скриптам (`scripts/`) и модулям (`src/modules/`), отсутствие ссылок на `ARCHITECTURE_SQLITE.md`
- **`.github/workflows/docs-check.yml`** — GitHub Actions workflow для автоматической проверки документации при `push` и `pull_request` (установка Python, запуск `python scripts/check_docs.py`, публикация отчёта)
- **Единый источник версии** — версия проекта централизована в трёх точках: `Mod_Constants.APP_VERSION` (VBA), `config.APP_VERSION` (Python), `$Script:AppVersion` (PowerShell). `scripts/check_docs.py` теперь импортирует версию из `config.APP_VERSION` вместо дублирования константы
- **`scripts/update_version.py`** — скрипт автоматического обновления версии при релизе (`python scripts/update_version.py 1.1.0`): обновляет `Mod_Constants.bas`, `config.py`, `config.ps1`, `README.md`, `docs/DEVELOPER.md`, `docs/ROADMAP.md`, `docs/ARCHITECTURE.md` и добавляет новый раздел в `docs/CHANGELOG.md`

### Changed
- **`docs/ARCHITECTURE_SQLITE.md`** переименован в **`docs/ARCHITECTURE.md`** — актуализация имени документа архитектуры; обновлены все внутренние ссылки в README, DEVELOPER, ROADMAP, sourcecraft-guide и `.ycarules`
- **Документация актуализирована до v0.18.0** — README.md, docs/ROADMAP.md, docs/DEVELOPER.md, docs/sourcecraft-guide.md, docs/git-workflow.md приведены в соответствие с фактической структурой проекта (13 модулей, классы `PartIdentity.cls`/`WorkIdentity.cls`, архивные классы листов `.bak`)

## [v0.17.0] — 2026-08-03

### Added
- **`src/modules/Mod_ModelTypes.bas`** — стандартный BAS-модуль с UDT (`WorkIdentity`, `PartIdentity`, `WorkEntry`) для устранения ошибки компиляции "Only public user defined types..."
- **`src/classes/WorkIdentity.cls`** и **`src/classes/PartIdentity.cls`** — классы с корректным заголовком `VERSION 1.0 CLASS`
- **Флаг `SilenceMsgBox`** в `Mod_Constants.bas` — устранение deadlock при COM-автоматизации
- **`plans/analysis_and_plan.md`** — сводный анализ тестового покрытия и план исправлений

### Changed
- **`src/modules/Mod_ModelDB.bas`** — ссылки `ModelTypes.*` заменены на `Mod_ModelTypes.*`
- **`scripts/impVBA.py`** — логика `strip_export_header()` теперь сохраняет заголовок для `.cls` классов
- **`Mod_AutoMatch.bas`** — `.Formula` заменён на `.FormulaLocal` для корректной работы с русским разделителем

### Fixed
- Ошибка компиляции "Only public user defined types..." устранена
- Ранее падавшие тесты TC-38, TC-39, TC-40 теперь проходят

### Test
- Результаты тестов: 25 PASS, 0 FAIL, 5 SKIP
- Выявлены пробелы в тестовом покрытии (6 модулей без тестов, пропуск TC-15..TC-30)

## [v0.16.0] — 2026-08-03

### Added
- **`.sourcecraft`** — корневой файл конфигурации SourceCraft с MCP-серверами (filesystem, git) на относительных путях `${workspaceFolder}`
- **Интеграция `.ycarules`** — подключение через `customInstructions.file` в `.sourcecraft`
- **Правила exclude/include/critical** — дублирование ключевых секций из `.ycarules` в `.sourcecraft` для работы на уровне платформы

### Changed
- **`docs/sourcecraft-guide.md`** — добавлен раздел "Конфигурация SourceCraft (`.sourcecraft`)" с описанием структуры и назначения файла; обновлена схема структуры проекта

### Removed
- **`.vscode/mcp.json`** — удалён, функциональность перенесена в `.sourcecraft` с заменой абсолютных путей на `${workspaceFolder}`

## [0.15.0] — 2026-07-25

### Added
- **Mod_ModelDB:** тесты TC-31..TC-35 (GetModelDBBasePath, GetModelGroupFilePath, ModelGroupFileExists, OpenModelGroupFile)
- **Mod_PickWork:** тесты TC-36..TC-38 (GetGroupNameFromMain, GetWorkSheetName, PickWork_UI)
- **Mod_AutoMatch:** тесты TC-39..TC-44 (AutoMatchWorks, AutoMatchParts, HighlightNotFound, ClearHighlight, IsAllFound)
- **Покрытие модулей тестами:** 46% → 69% (9 из 13 модулей)

### Changed
- **Mod_ModelDB:** константа `MODELDB_BASE_PATH` помечена как deprecated (заменена на `GetModelDBBasePath`)
- **Mod_MainButtons:** `Btn_main_ImportVH_Click` делегирован `Mod_ButtonDispatcher`
- **`scripts/export_vba.py`:** COMPONENTS расширен до 13 модулей + 3 листа
- **`scripts/Set-ExcelTrust.ps1`:** хардкодный путь `L:\PROject\SysW` заменён на `$ProjectPath`
- **`docs/DEVELOPER.md`:** синхронизирован с актуальной архитектурой (13 модулей, 38 тестов)
- **`docs/sourcecraft-guide.md`:** обновлён до v0.14.0

## [0.14.0] — 2026-07-25

### Added
- **Константа `LOGS_DIR`:** добавлена в `Mod_Constants.bas` — централизованное значение директории логов (`"logs"`)
- **Директория `logs/`:** создана для хранения логов и результатов тестов вместо корня проекта
- **`scripts/config.py`:** добавлена константа `LOGS_DIR` и `TEST_LOG_FILE` (путь к `logs/test_results.log`)
- **`scripts/run_tests.py`:** автоматическое создание `logs/` при запуске (`ensure_logs_dir`)

### Changed
- **`Mod_Logger.bas`:** путь к лог-файлу изменён с `log.txt` на `logs/log.txt` через `Mod_Constants.LOGS_DIR`
- **`Mod_FullTestRunner.bas`:** хардкодные пути к логам заменены на константы из `Mod_Constants`
- **`scripts/run_tests.py`:** результаты тестов сохраняются в `logs/test_results.log` вместо `test_results.log` в корне

### Removed
- **Мусорные копии из `base/models/`:** удалены неиспользуемые файлы, оставлен только `UAZ.xlsm` и `.gitkeep`
- **Старые лог-файлы из корня:** `log.txt` и `log_old.txt` удалены из корня проекта (перенесены в `logs/`)

## [0.13.0] — 2026-07-25

### Added
- **Централизация путей:** созданы `scripts/config.py` и `scripts/config.ps1` — единая точка конфигурации путей для Python и PowerShell скриптов
- **Скрипт удаления work26:** `scripts/remove_work26.py` — удаление встроенной книги work26.xlsm из UAZ.xlsm через COM

### Changed
- **Пути в Python-скриптах:** `export_vba.py`, `impVBA.py`, `run_tests.py` — жёсткие абсолютные пути `L:\PROject\SysW\...` заменены на относительные через `config.py`
- **Пути в PowerShell:** `Set-ExcelTrust.ps1` — жёсткие пути заменены на переменные из `config.ps1`
- **VBA-привязка:** `Mod_MainButtons.bas:158` — `'work.xlsm'` заменён на `ThisWorkbook.Name`

### Removed
- **Мёртвый код VBA (6 процедур-заглушек):** `Btn_z4_Action1`, `Btn_z4_Action2`, `Btn_z4_Action3`, `Btn_work_Action1`, `Btn_work_Action2`, `Btn_work_Action3` из `Mod_SheetButtons.bas`
- **Мёртвый код VBA (дубликат + заглушка):** `Btn_main_ImportVH_Click` (дубликат `Mod_ButtonDispatcher`) и `Btn_main_MANz4` (заглушка) из `Mod_MainButtons.bas`
- **Мёртвый код VBA (невызываемая функция):** `AssignMainButtons` из `Mod_MainButtons.bas`
- **Неиспользуемые константы:** `MAIN_HEADER_END_ROW`, `MAIN_CLEAR_START_ROW`, `MAIN_HEADER_RANGE` из `Mod_Constants.bas`
- **7 неиспользуемых скриптов:** `export_uaz_vba.ps1`, `export_uaz_vba.py`, `parse_excel.ps1`, `parse_full.ps1`, `parse_full.py`, `rewire_uaz_buttons.ps1`, `Import-VbaFromExcel.ps1`
- **Директория `scripts/_uaz_vba_export/`** (включая `work26/`)

### Security
- Установлен `AccessVBOM=1` и `VBAWarnings=1` в реестре Excel для доступа к VBA Project Object Model

## [0.12.0] — 2026-07-21

### Added
- **Ручной подбор работ (кнопка РУЧ РАБ):** новый модуль `Mod_PickWork.bas` с процедурой `PickWork_UI` — открывает файл группы из `base/models/` по значению B14, активирует лист работ, пользователь ищет через фильтр и копирует данные вручную в диапазон E4:H
- **Модуль доступа к файлам групп:** `Mod_ModelDB.bas` с функциями `OpenModelGroupFile` (открытие файла группы) и `GetWorks` (получение списка работ) — базовый слой для работы с `base/models/`
- Константы колонок ручного подбора работ: `MANWRK_COL_ARTICLE` (E), `MANWRK_COL_NAME` (F), `MANWRK_COL_NORMHOURS` (G), `MANWRK_COL_QTY` (H), `MANWRK_START_ROW` (4)

### Changed
- Заглушка `Btn_main_MANw` заменена на рабочую `Btn_main_MANWRK` с вызовом `Mod_PickWork.PickWork_UI`

## [0.11.0] — 2026-07-21

### Changed
- Сдвиг структуры листа main на 2 строки вниз для резервирования места под будущую логику
- Все жёсткие ссылки на строки обновлены (+2)
- Добавлены константы `MAIN_HEADER_START_ROW`, `MAIN_DATA_START_ROW` и др.

## [0.10.0] — 2026-07-21

### Added
- **Кнопка "ИМПОРТ ВХ":** новая процедура `ImportFromB2_UI` в `Mod_Import.bas` — импорт данных на лист "мэйн" из листа `{B2}M`; если листа нет — копирует из `report.xlsx`
- Обработчик `Btn_main_ImportVH_Click` в `Mod_MainButtons.bas` и `Mod_ButtonDispatcher.bas`
- Процедура `AssignMainButtons` в `Mod_MainButtons.bas` с mapping для новой кнопки

### Исправлено
- Ошибка импорта в `Mod_Import.bas`: циклы пропуска заголовка таблиц с `IsNumeric()` останавливались на строке подзаголовка (нумерация столбцов), а не на данных. Заменены на безусловный пропуск 2 строк заголовка. Импорт данных услуг и материалов теперь выполняется корректно, без смещения.

## [0.9.0] — 2026-07-21

### Changed
- **Рефакторинг:** модуль `Mod_LibName` объединён с `Mod_Constants` — централизованное управление константами и реестром имён
- `Mod_LibName.bas` удалён; функциональность (`InitLibName`, `AddWorkEntry`, `BuildEntryArray`) перенесена в `Mod_Constants.bas`
- Добавлены строковые константы `*_NAME` для листа libname (15 констант)
- Добавлены числовые константы `SPISOK_COL_NUM` (столбец A) и `SPISOK_COL_GROUP` (столбец I) для листа spisok
- Добавлены записи в реестр libname: `models_col_group`, `z4`
- Исправлено несоответствие: `spisok_col_reserve` → `spisok_col_note` (синхронизация с `SPISOK_COL_NOTE`)
- Обновлён тест TC-13: вызов `Mod_LibName.InitLibName` → `Mod_Constants.InitLibName`, добавлены проверки `models_col_group` и `z4`
- Обновлена документация: DEVELOPER.md, sourcecraft-guide.md

## [0.8.0] — 2026-07-19

### Added
- Новый лист `models` со структурой (русские заголовки + латинские имена)
- Модуль `Mod_Models` для инициализации листа models
- Расширение реестра имён `Mod_LibName` записями для models (`col_model_name`, `col_hrpr`)
- Константы столбцов `MODELS_COL_MODEL`, `MODELS_COL_GROUP`, `MODELS_COL_PRICE` в `Mod_Constants`

## [0.7.1] — 2026-07-16

### Добавлено
- Защитное программирование и обработка ошибок в модулях VBA (аудит стабильности)

### Исправлено
- **OH-01** (`Mod_OrderHeader.bas`): добавлены проверки `Nothing` для листов `main`, `spisok`, `model` перед обращением в `FillHeaderFromOrder`
- **OH-02** (`Mod_OrderHeader.bas`): добавлен обработчик ошибок с восстановлением `Application.EnableEvents` в `FillHeaderFromOrder`
- **IM-01** (`Mod_Import.bas`): добавлен блок `On Error GoTo ErrHandler` в `ImportSheet` с восстановлением `EnableEvents`, `ScreenUpdating`, `DisplayAlerts`
- **IM-02** (`Mod_Import.bas`): добавлен обработчик ошибок с восстановлением состояния приложения в `ImportDataToMain`
- **SO-01** (`Mod_SheetOps.bas`): исправлена утечка объекта Workbook в `SearchSheetByGRZ` — закрытие книги в `ErrHandler`
- **SO-02** (`Mod_SheetOps.bas`): удаление листов в `RenameSheetsByGRZ` теперь с проверкой существования через цикл по `Sheets`
- **SO-03** (`Mod_SheetOps.bas`): добавлено восстановление `DisplayAlerts` в обработчиках ошибок `RenameSheetsByGRZ`, `ClearMainSheet_UI`, `ClearHeader_UI`
- **UT-01** (`Mod_Utils.bas`): добавлена проверка `wb Is Nothing` в `GetSheetByName`
- **LG-01** (`Mod_Logger.bas`): заменён `On Error GoTo ErrHandler` на `On Error Resume Next` в `RotateLogIfNeeded` и `ClearLog` для игнорирования ошибок файловых операций
- **SM-01** (`Sheet1_main.cls`): добавлен обработчик ошибок с восстановлением `EnableEvents` и сбросом флага `isProcessing` в `Worksheet_Change`
- **FB-01** (`Mod_MainButtons.bas`): добавлено восстановление `EnableEvents`, `ScreenUpdating`, `DisplayAlerts` в обработчиках ошибок `Btn_main_Clear` и `Btn_main_Import`

## [0.7.0] — 2026-07-15

### Добавлено
- Новые тесты TC-21..TC-30: покрытие Mod_Logger (WriteLog, RotateLogIfNeeded, ClearLog),
  Mod_SheetOps (ClearMainSheet_UI, ClearHeader_UI, RenameSheetsByGRZ),
  Mod_Import (ImportSheet, ImportDataToMain), граничные случаи FormatDateSQL,
  автоматизированный тест Btn_main_Clear_Click (silent mode)
- Функция `GetTestResults()` в Mod_FullTestRunner для программного сбора результатов
- Параметр `silent` в `ClearMainSheet_UI` для автоматических тестов без MsgBox
- Раздел "Тестирование" в `docs/DEVELOPER.md` с таблицей TC-01..TC-30, покрытием модулей,
  инструкцией по добавлению тестов и описанием CI/CD

### Исправлено
- TC-05: ожидаемое значение `"12377"` → `"123"` (ExtractNumberFromGRZ возвращает первую
  цифровую группу из 3+ цифр, а не все цифры подряд)
- TC-13, кейс 1: ожидаемое значение `"12377"` → `"123"` (аналогично TC-05)

### Изменено
- `scripts/run_tests.py` — полная переработка:
  - Программный сбор результатов через `excel.Run("GetTestResults")`
  - Exit code: 0 если все PASS, 1 если есть FAIL
  - Русский язык вместо транслита
  - Гарантированное закрытие Excel в `finally` блоке
  - Сохранение результатов в `test_results.log`
- `docs/MODERNIZATION_CHECKLIST.md` — отмечены выполненными пункты 3.3 и 3.4

## [0.6.0] — 2026-07-15

### Добавлено
- Новая структура каталогов: `src/modules/` для `.bas` и `src/sheets/` для `.cls` файлов
- Директория `src/` для исходного кода VBA

### Изменено
- Все VBA-модули (.bas) перемещены из корня проекта в `src/modules/`
- Все VBA-классы листов (.cls) перемещены из корня проекта в `src/sheets/`
- Python-скрипты (`export_vba.py`, `impVBA.py`, `run_tests.py`) перемещены в `scripts/`
- `scripts/impVBA.py` — обновлён `MODULES_PATH` на `src`, поиск `.bas`/`.cls` по поддиректориям
- `scripts/export_vba.py` — обновлён `PROJECT_DIR` на `src`, `COMPONENTS` с путями `modules/` и `sheets/`
- `.github/workflows/vba-check.yml` — обновлены пути к VBA-файлам
- `scripts/Import-VbaFromExcel.ps1` — обновлён `$OutputDir` на `src`, `$componentMap` с путями
- `docs/DEVELOPER.md` — обновлены все относительные ссылки на файлы
- `docs/sourcecraft-guide.md` — обновлена схема структуры проекта и ссылки
- `.ycarules` — обновлено правило `[S1]` с описанием новой структуры

## [0.5.0] — 2026-07-14

### Добавлено
- Техническая документация `docs/DEVELOPER.md` с описанием архитектуры VBA-модулей, правил кодировки, процессов импорта/экспорта, настройки окружения и CI/CD
- Секция "Структура проекта" в `README.md` со схемой директорий и ссылками на документацию
- Секция "Быстрый старт" в `README.md` с пошаговой инструкцией начала работы
- Секция "Технологический стек" в `README.md` с детализацией используемых технологий
- Секция "Документация" в `README.md` со ссылками на все руководства

### Изменено
- `README.md` — расширено описание проекта, обновлён состав команды с ролями SourceCraft (Оркестратор, Architect, Code, Debug, Ask)
- `docs/sourcecraft-guide.md` — добавлена ссылка на `docs/DEVELOPER.md` в раздел "Структура проекта"

## [0.4.0] — 2026-07-14

### Добавлено
- GitHub Actions workflow `.github/workflows/vba-check.yml` для автоматической проверки VBA-модулей при пуше в main/dev и Pull Request
- Git-инструкции `docs/git-workflow.md` с описанием веточной стратегии, формата коммитов (Conventional Commits) и pre-commit процедур
- CI-проверки: наличие VBA-файлов, кодировка UTF-8, базовый синтаксис, актуальность CHANGELOG, консистентность .gitignore

### Изменено
- `.gitignore` — добавлены правила для временных директорий скриптов (`_temp_export/`, `_temp_import/`), Python-кэша (`__pycache__/`, `*.pyc`, `*.pyo`, `venv/`) и системных файлов (`.DS_Store`)
- `docs/sourcecraft-guide.md` — добавлена ссылка на `docs/git-workflow.md` в раздел "Структура проекта"

## [0.3.0] — 2026-07-14

### Добавлено
- Интеграция SourceCraft Code Assistant (Этап 1)
- Новые правила в `.ycarules`: директория планов, структура проекта, PowerShell-скрипты и двухфазная кодировка, обновление CHANGELOG.md
- PowerShell-скрипт `scripts/Import-VbaFromExcel.ps1` для импорта VBA-модулей из Excel с конвертацией CP1251 → UTF-8
- Руководство `docs/sourcecraft-guide.md` по работе с SourceCraft Code Assistant
- Директории `scripts/` и `docs/` для скриптов автоматизации и документации

### Изменено
- `.gitattributes` — добавлена нормализация для `*.ps1` файлов

## [0.2.0] — 2026-07-13

### Добавлено
- Полный тестовый модуль `Mod_FullTestRunner.bas` (20 сценариев TC-01..TC-20)
- Автоматические тесты утилит, OrderHeader, импорта, кнопок
- Механизм SKIP для тестов, зависящих от отсутствующих данных
- Сохранение/восстановление состояния листов в тестах

### Исправлено
- Ссылки в `run_tests.py` и `Mod_ButtonDispatcher.bas` на новый тестовый модуль
- Список импортируемых файлов в `import_all_vba.py`

### Изменено
- Замена `Mod_MinimalTestRunner.bas` на `Mod_FullTestRunner.bas`

## [0.1.0] — 2026-07-09

### Добавлено
- Автоматическое заполнение шапки заказ-наряда при вводе номера заказа в ячейку B2 листа main.
- Модуль `Mod_OrderHeader` с независимой функцией `FillHeaderFromOrder`.
- Обработчик `Worksheet_Change` для листа main.
- Инструмент `Write-UTF8BOM.ps1` для сохранения PowerShell-скриптов с BOM.
- Файл правил `.ycarules` для Code Assistant.
- Git-репозиторий с ветками `main` и `dev`.
- Начальная документация (README.md, .gitignore).
