# План релиза SysW v1.0.5

> **Проект:** SysW (рабочая директория: `f:/MiTyA/PROject/SysW`)
> **Статус:** Согласован с юзером (объём релиза утверждён)
> **Автор:** Архитектор (режим Architect)
> **Дата:** 2026-08-21
> **Формат:** Keep a Changelog (см. `[U2]`), язык — русский

---

## 0. Резюме (цель и согласованный объём)

Релиз `v1.0.5` доводит до конца миграцию хранения модельных данных на **SQLite** и фиксирует версию системы. Согласованный объём:

1. **SQLite-миграция: Фазы A И B полностью.** Инфраструктура Этапа 0 уже частично есть (`db/schema.sql`, `scripts/sqlite_schema.py`, константы `DB_*` в `scripts/config.py`) → довести до конца: конвертер, VBA-провайдеры, перевод чтения (Фаза A) и записи/интеграции (Фаза B).
2. **Зависание `build_templates.py` закрывается ОБХОДНЫМ путём** (проблема признана закрытой). Вариант A пересборки через openpyxl **НЕ внедряется**. Шаблоны `base/templates/*` уже созданы и защищены на уровне файлов в v1.0.4; новых действий по пересборке не требуется.
3. **Версия системы обновляется до 1.0.5**, документация актуализируется во всех точках.

> **Ключевое решение юзера:** полная миграция Фаз A+B; зависание шаблонов закрывается обходным путём без внедрения варианта A.

---

## 1. Архивация выполненных планов

Перенести в `plans/_archive/` следующие выполненные планы (правило `[U1]`, `[U2]`):

| Файл | Статус |
|------|--------|
| [`plans/analysis_1.0.4.md`](../plans/analysis_1.0.4.md) | ✅ выполнен |
| [`plans/debug_com_error.md`](../plans/debug_com_error.md) | ✅ выполнен |
| [`plans/integration_1.0.4.md`](../plans/integration_1.0.4.md) | ✅ выполнен |
| [`plans/plan_1.0.4.md`](../plans/plan_1.0.4.md) | ✅ выполнен |
| [`plans/debug_hang_templates_v104.md`](../plans/debug_hang_templates_v104.md) | ✅ закрывается обходным путём (решение юзера) |

> **Остаётся в `plans/` (не архивируется):** [`plans/migration_sqlite_plan.md`](../plans/migration_sqlite_plan.md) — его задачи реализуются в этом релизе. При необходимости после завершения Фаз A+B может быть также перенесён в архив отдельным решением.

**Действия исполнителя (Code):**
- Переместить 5 файлов в `plans/_archive/`.
- Зафиксировать факт архивации в `docs/CHANGELOG.md` (раздел `v1.0.5`, подраздел `Archived`).

---

## 2. Этап 0 — Конвертер `migrate_models_to_sqlite.py` (создание `SysW.db`)

> **Режим:** Code. **Источник:** разделы 1.6, 4 `migration_sqlite_plan.md`.

### 2.1. Что уже готово (не переделывать)
- [`db/schema.sql`](../db/schema.sql) — DDL-схема (v1, `PRAGMA user_version = 1`).
- [`scripts/sqlite_schema.py`](../scripts/sqlite_schema.py) — Python-структура схемы, `init_db()`, агрегаты, `user_version`.
- Константы в [`scripts/config.py`](../scripts/config.py): `APP_VERSION`, `DB_SCHEMA_VERSION`, `DB_PATH`, `DB_SCHEMA_PATH`, `MODELS_DIR`, `MIGRATION_REPORT_FILE`.

### 2.2. Создать конвертер `scripts/migrate_models_to_sqlite.py`
- Только stdlib + `openpyxl`. **Без** `win32com` (обходит COM-зависание).
- Чтение 6 моделей `base/models/{4x4,2170,2180,2190,GAZ,UAZ}.xlsm`; список групп — сканирование `MODELS_DIR` по `*.xlsm` (надёжнее фиксированного списка).
- `load_workbook(..., read_only=True, data_only=True)` — брать **значения** (кэшированные), не формулы; `None` для формул → `0`/пустая строка с логированием.
- Маппинг листов → таблицы (см. таблицу ниже), критерий тождеств: непустой столбец B (OutArticle) И непустой I (Агрегат).
- Идемпотентность: повторный запуск не дублирует данные (`DELETE + INSERT` в транзакции либо `INSERT OR REPLACE`).
- Перед перезаписью — бэкап существующего `SysW.db` в `_backup/SysW_<дата>.db`.
- Выставить `PRAGMA user_version = 1`, `PRAGMA journal_mode=WAL` (см. Этап 4 / ODBC).
- Записать отчёт `logs/migration_report.log` (контрольные числа, см. 2.4).
- Добавить CLI-аргумент `--force` (полный пересозданный прогон).

### 2.3. Маппинг листов → таблицы (заголовки стр. 3, данные с 4-й)
| Лист `{G}.xlsm` | Таблица | Колонки |
|---|---|---|
| `{G}` (все работы) | `works` | A→code, B→name, C→unit, D→norm_hours, E→price, F→note |
| `{G}w` (модельные работы + тождества) | `model_works` | B→out_article, C→out_name, D→norm_hours, G→qty_zn, I→aggregate, J→in_name |
| `z4` (все запчасти) | `parts` | A→code, B→name, C→unit, D→price, E→note |
| `{G}z4` (модельные запчасти + тождества) | `model_parts` | B→out_article, C→out_name, G→qty_zn, F→price, I→aggregate, J→in_catnum, K→in_name |
| — | `matlib_entries` | нормализация тождеств: для работ `entry_code=InName`, `target_code=OutArticle`, `target_type='mod_work'`; для запчастей `entry_code=InCatNum`, `target_code=OutArticle`, `target_type='mod_part'`; `coefficient=qty_zn` |
| служебные `{NN}M` | — | не переносятся |

### 2.4. Контрольные объёмы и валидация
| Источник | Ожидание |
|---|---|
| `works` по всем группам | Σ строк данных в листах `{G}` |
| `parts` по всем группам | Σ строк данных в листах `z4` |
| `model_works` | Σ тождеств в `{G}w` с B и I непустыми |
| `model_parts` | Σ тождеств в `{G}z4` с B и I непустыми |

- Отчёт `logs/migration_report.log`: дата, группы, количества строк по каждой таблице, время, ошибки.
- Валидация: скрипт сравнения контрольных чисел (встроена в конвертер) + `PRAGMA integrity_check`.

**Критерий готовности Этапа 0:** `SysW.db` создан, наполнен из 6 групп, контрольные числа совпадают с xlsm, повторный запуск идемпотентен, отчёт записан.

---

## 3. Этап 1 — VBA-классы провайдеров и рефакторинг `Mod_ModelDB`

> **Режим:** Code. **Источник:** разделы 1.3–1.5, 5 `migration_sqlite_plan.md`.

### 3.1. Новые классы (`src/classes/`)
| Файл | Назначение |
|---|---|
| `IModelDataProvider.cls` | Интерфейс-контракт: методы `GetWorks`, `GetParts`, `GetModelWorks`, `GetModelParts`, `GetMatLibEntries`, `GetWorkIdentities`, `GetPartIdentities`, `GetAllModelGroups`, `CreateModelGroupFile`, `ModelGroupFileExists`. Все методы — заглушки `Err.Raise vbObjectError+1`. Возвращают `Collection` объектов. |
| `WorkEntry.cls` | Класс записи работы (вынести из UDT [`Mod_ModelTypes.bas`](../src/modules/Mod_ModelTypes.bas:17)). Поля: `Code`, `Name`, `Unit`, `NormHours`, `Price`, `Note`. |
| `Mod_SQLiteDB.cls` | SQLite-провайдер через ADO/ODBC (`Implements IModelDataProvider`). `OpenConnection`/`CloseConnection`/`ExecuteScalar`/`ExecuteQuery`/`ExecuteNonQuery`; параметризованные запросы (`Command.Parameters.Append`); `Class_Terminate` закрывает соединение. |
| `Mod_ModelDBProvider.cls` | Excel-реализация интерфейса (`Implements IModelDataProvider`), обёртка над текущей логикой чтения листов. |

### 3.2. Рефакторинг `Mod_ModelDB.bas`
- Добавить фабрику `GetModelDataProvider() As Object`: при наличии `SysW.db` рядом с `work.xlsm` (`ThisWorkbook.Path & "\SysW.db"`) и доступном ODBC-драйвере → `New Mod_SQLiteDB`, иначе → `New Mod_ModelDBProvider`.
- Сохранить служебные функции: `GetModelDBBasePath`, `GetModelGroupFilePath`, `ModelGroupFileExists`, `GetAllModelGroups`.
- Существующие `GetWorks`, `GetWorkIdentities`, `GetPartIdentities` переписать как делегаты к активному провайдеру (для обратной совместимости вызовов из `Mod_AutoMatch` и тестов).
- `WorkEntry` перевести с UDT на класс; обновить обращения.

### 3.3. Fallback-константа в `Mod_Constants.bas`
Добавить в [`src/modules/Mod_Constants.bas`](../src/modules/Mod_Constants.bas):
```vba
Public Const MODELDB_PROVIDER_SQLITE As Boolean = True   ' True = SQLite, False = Excel
```
Фабрика учитывает константу: если `True`, но БД/драйвер недоступны — graceful fallback на Excel-провайдер.

**Критерий готовности Этапа 1:** код компилируется; фабрика возвращает SQLite-провайдер при наличии `SysW.db`; константа добавлена.

---

## 4. Этап 2 (Фаза A) — перевод чтения на SQLite

> **Режим:** Code + Debug. **Источник:** раздел 2.1, 6.2 `migration_sqlite_plan.md`.

### 4.1. Чтение моделей через провайдер
Перевести на SQLite (через `GetModelDataProvider()` / `Mod_SQLiteDB`):
- `GetWorks(groupName, filters)` → `SELECT ... FROM works WHERE group_name = ?` (+ фильтры).
- `GetWorkIdentities(groupName)` → `SELECT out_article, out_name, norm_hours, qty_zn, aggregate, in_name FROM model_works WHERE group_name = ?`.
- `GetPartIdentities(groupName)` → `SELECT ... FROM model_parts WHERE group_name = ?`.

### 4.2. Поведение в Фазе A
- `Mod_AutoMatch` (кнопки АВТО РАБ/АВТО ЗЧ) работает через провайдер (читает тождества из БД).
- `Mod_PickWork` (ручной подбор) в Фазе A остаётся на Excel-провайдере (нужен живой Workbook для UI).
- Excel-файлы `base/models/*.xlsm` — read-only легаси для записи до Фазы B.

### 4.3. Адаптация тестов
- **TC-22..24** (`RunModelDBReadTests`): переписать на `Mod_SQLiteDB`/`GetModelDataProvider` — проверять непустые коллекции из `SysW.db` **без** `OpenModelGroupFile(UAZ.xlsm)`.
- **TC-31..35** (`RunModelDBTests`): `GetModelGroupFilePath`/`ModelGroupFileExists` оставить; **TC-35 OpenModelGroupFile** заменить на проверку, что `GetModelDataProvider()` возвращает SQLite-провайдер и `OpenConnection` успешен.
- **TC-39, TC-40** (`AutoMatchWorks`/`AutoMatchParts`): тождества из `SysW.db` (провайдер), не из `UAZ.xlsm`.
- **TC-44** — остаётся пропущенным (изменяет данные на листе).
- Новые тесты: **TC-S1** (фабрика возвращает SQLite-провайдер при наличии БД), **TC-S2** (`GetWorks` через SQLite эквивалентен Excel), **TC-S3** (конвертер даёт контрольные объёмы).

**Критерий готовности Фазы A:** чтение моделей идёт из `SysW.db` без COM-зависаний; fallback на Excel работает; тесты TC-22..24, 31-35, 39-40, TC-S1..S3 проходят.

---

## 5. Этап 3 (Фаза B) — запись, интеграция, `CreateModelGroupFile`

> **Режим:** Code. **Источник:** раздел 2.2 `migration_sqlite_plan.md`.

### 5.1. `CreateModelGroupFile(groupName)`
- SQLite-реализация: `INSERT OR REPLACE INTO model_groups(group_name) VALUES (?)` (+ пустые наборы не требуются — таблицы уже есть).
- Excel-реализация: создание легаси-файла (резервный вариант).

### 5.2. Интеграция бизнес-логики
- **`Mod_Import`** (`ImportFromB2_UI`, `ImportDataToMain`): после импорта — `FindModelGroupByModel` + `GetMatLibEntries` + подстановка модельных кодов через провайдер.
- **`Mod_OrderHeader.FillHeaderFromOrder`**: чтение данных работ/запчастей через провайдер; автодобавление модели в лист `models` остаётся, группа/цена могут читаться из БД.
- **`Mod_AutoMatch`**: полностью на SQLite-провайдере.
- **Запись тождеств через ADO/ODBC** (если юзер правит соответствия в новом UI) — `INSERT/UPDATE` в `model_works`/`model_parts`/`matlib_entries`.
- Переключить константу `MODELDB_PROVIDER_SQLITE` в активное состояние (если не активна).

**Критерий готовности Фазы B:** полный цикл «импорт ЗН → автоподбор работ/запчастей → заполнение шапки» работает на `SysW.db` без открытия `base/models/*.xlsm`; `CreateModelGroupFile` создаёт запись в БД.

---

## 6. Этап 4 — Тесты, ODBC-драйвер, WAL/busy_timeout

> **Режим:** Debug + Code. **Источник:** разделы 3.2, 7.1, 6.2 `migration_sqlite_plan.md`.

### 6.1. ODBC-драйвер SQLite3
- Проверить наличие драйвера `SQLite3 ODBC Driver` (`sqliteodbc.dll`) в целевой среде.
- Зафиксировать имя драйвера в константе VBA; при старте — проверка доступности, иначе graceful fallback на Excel-провайдер.
- Требуется ссылка **Microsoft ActiveX Data Objects 6.1 Library** (`ADODB`).

### 6.2. WAL и busy_timeout
- В конвертере и `Mod_SQLiteDB.OpenConnection`: `PRAGMA journal_mode=WAL` и `PRAGMA busy_timeout=<N>` (например, 5000 мс) — конкурентный доступ Python-писатель + VBA-читатели.
- Опционально: `PRAGMA synchronous=NORMAL`.

### 6.3. Прогон тестов
- Запуск: `python scripts/run_tests.py` (см. [`scripts/run_tests.py`](../scripts/run_tests.py)).
- Полный прогон (46+ тестов), отсутствие COM-зависаний.
- Зафиксировать результаты в `logs/test_results.log`.

**Критерий готовности Этапа 4:** ODBC-драйвер установлен/проверен; WAL+`busy_timeout` активны; все тесты зелёные без зависаний.

---

## 7. Этап 5 — Версия 1.0.5 и документация

> **Режим:** Architect + Code. **Источник:** `[U2]`, `scripts/update_version.py`.

### 7.1. Обновление версии
Запуск: `python scripts/update_version.py 1.0.5`. Скрипт обновляет:
- [`src/modules/Mod_Constants.bas`](../src/modules/Mod_Constants.bas:18) — `APP_VERSION = "1.0.4"` → `"1.0.5"`.
- [`scripts/config.py`](../scripts/config.py:11), [`scripts/config.ps1`](../scripts/config.ps1:7).
- [`README.md`](../README.md:3), [`docs/DEVELOPER.md`](../docs/DEVELOPER.md:1), [`docs/ROADMAP.md`](../docs/ROADMAP.md:3), [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md:4).
- [`docs/CHANGELOG.md`](../docs/CHANGELOG.md) — новый раздел `## [v1.0.5]`.

После запуска — ручная проверка всех точек (скрипт мог пропустить шаблоны с иным префиксом).

### 7.2. Обновление документации
| Файл | Правки |
|---|---|
| `docs/CHANGELOG.md` | Новый раздел `## [v1.0.5]` (Added/Changed/Removed/Archived) в формате Keep a Changelog. |
| `docs/ROADMAP.md` | Версия → 1.0.5; статусы задач **R-S1..R-S8** (SQLite-миграция) → выполнены; убрать «Планируется». |
| `docs/ARCHITECTURE.md` | Версия → 1.0.5; раздел 5 DDL актуализировать; добавить ADO/ODBC, WAL/busy_timeout; снять «Частично реализовано» (строка 5). |
| `docs/DEVELOPER.md` | Версия → 1.0.5; добавить конвертер, провайдеры (`IModelDataProvider`, `Mod_SQLiteDB`, `Mod_ModelDBProvider`), ODBC, WAL, процедуру миграции/отката. |
| `README.md` | Версия → 1.0.5; блок «SQLite» (строка 27): «Планируется» → «Реализовано». |
| `docs/table.md`, `docs/sourcecraft-guide.md`, `docs/git-workflow.md` | Проверить и синхронизировать упоминания (SQLite, версия, структура). |

**Критерий готовности Этапа 5:** версия 1.0.5 во всех точках; документация актуальна и непротиворечива; CHANGELOG содержит раздел v1.0.5.

---

## 8. Этап 6 — Верификация, чистота проекта, Git

> **Режим:** Debug + Architect. **Источник:** `[S7]`, `[S9]`, `[G10]`.

### 8.1. Верификация
- Повторный прогон `run_tests.py` после всех изменений.
- Проверка отсутствия COM-зависаний.
- Контрольный вопрос: чтение/запись моделей работают через `SysW.db`.

### 8.2. Чистота проекта
- Удалить временные файлы: `__pycache__/`, `*.sync-conflict-*`, `~$*.xlsm` (правило `[S7]`).
- Убедиться, что `SysW.db` намеренно создан и не попадает в мусор (добавить в `.gitignore` при необходимости — только по согласованию, правило `[E3]` на критичные файлы не распространяется на `SysW.db`).

### 8.3. Git (после согласования с юзером)
- Работа ведётся в ветке `dev`.
- Коммиты по этапам (архивация, конвертер, VBA-провайдеры, Фаза A, Фаза B, тесты/ODBC, версия/документация).
- По завершении и подтверждению юзера — слияние `dev → main`, возврат в `dev`.
- Коммит и пуш — **только после явного запроса/согласования с юзером** (правило `.ycarules` «после завершения всех задач - запрос у юзера и выполнение коммита и пуша»). Git-операции — через MCP Git Tools (`[G10]`).

---

## 9. Риски и контрольные вопросы

| # | Риск/вопрос | Влияние | Митигация / откат |
|---|---|---|---|
| 1 | Дублирование схемы: `db/schema.sql` vs `scripts/sqlite_schema.py` | Среднее | Оставить оба, но зафиксировать в `docs/DEVELOPER.md`, что `sqlite_schema.py` — рабочая структура, `schema.sql` — эталон документации; при расхождении — синхронизация вручную. Рекомендация на будущее — единый источник. |
| 2 | ODBC-драйвер SQLite3 отсутствует в целевой среде | Критическое | Проверка драйвера при старте; fallback на Excel-провайдер через константу; документация по установке драйвера. |
| 3 | Несоответствие контрольных объёмов (потеря данных при миграции) | Высокое | Контрольные числа в `logs/migration_report.log`; бэкап `_backup/SysW_<дата>.db`; восстановление из бэкапа. |
| 4 | `data_only=True` возвращает `None` для формул | Высокое | Подстановка 0/пусто с логированием; повторный запуск конвертера. |
| 5 | Рефакторинг `WorkEntry` (UDT→класс) ломает потребителей | Среднее | Компиляция-проверка; замена всех обращений; сохранение обёртки при необходимости. |
| 6 | Конкурентная запись Python + чтение VBA | Среднее | WAL + `busy_timeout`; в Фазе A единственный писатель — Python. |
| 7 | COM-зависание сохраняется в тестах (если fallback на Excel) | Высокое | Тесты всегда на SQLite-провайдере; Excel-файлы не открываются в COM-тестах. |
| 8 | Порядок отката миграции | — | 1) Фаза A: вернуть чтение на Excel-провайдер (константа `False`); 2) Фаза B: восстановить `SysW.db` из бэкапа `_backup/SysW_<дата>.db`. |

**Порядок отката (итог):** константа `MODELDB_PROVIDER_SQLITE` → `False` возвращает Excel-режим (Фаза A); для Фазы B — восстановление `SysW.db` из резервной копии.

---

## 10. Порядок выполнения (сводный чек-лист)

1. Архивация планов в `plans/_archive/` (раздел 1).
2. Конвертер `migrate_models_to_sqlite.py` → `SysW.db` + отчёт (раздел 2).
3. VBA-классы и рефакторинг `Mod_ModelDB` + константа (раздел 3).
4. Фаза A: чтение на SQLite + адаптация тестов (раздел 4).
5. Фаза B: `CreateModelGroupFile`, интеграция, запись через ADO (раздел 5).
6. ODBC/WAL + полный прогон тестов (раздел 6).
7. Версия 1.0.5 + документация (раздел 7).
8. Верификация, чистота, Git-слияние `dev → main` (раздел 8).

> Каждый этап — с критериями готовности (указаны в соответствующих разделах) и проверкой компиляции/тестов.