# Анализ обновления SysW до версии 1.0.4

> **Этап:** Анализ (без изменения кода).
> **Рабочий проект:** `f:/MiTyA/PROject/SysW` — текущая версия **v1.0.3**.
> **Эталонная копия:** `F:/MiTyA/PROject/copy/SysW` — целевая версия **v1.0.4**.
> **Дата:** 2026-08-21.
> **Правила:** строго соблюдать `.ycarules` (приоритет — [G1], [G3], [Z2], [Z4], [Z5], [E3], [S7]).

---

## 1. Резюме

Рабочий проект находится на версии **v1.0.3**, эталонная копия — на **v1.0.4**.
Основное содержание релиза v1.0.4 (из [`docs/CHANGELOG.md`](../docs/CHANGELOG.md) копии):

1. Создан и интегрирован каталог шаблонов `base/templates/` (5 файлов: `work.xlsm`, `work0.xlsm`, `model.xlsm`, `model0.xlsm`, `report0.xlsx`).
2. Применена защита листов шаблонов (`Protect` + `AllowEditRanges` + `FreezePanes A4`) **без изменения VBA-кода**.
3. Версия системы установлена на **v1.0.4** во всех точках единого источника.
4. Удалён каталог `workOt/`.
5. Добавлены новые скрипты и планы (в т.ч. по SQLite — на стадии проектирования).
6. Документация актуализирована (разделы о шаблонах и защите листов).

**Важное замечание пользователя:** номер версии не является гарантией идентичности содержимого — сравнение выполнялось **по содержимому файлов**, а не только по номеру версии.

---

## 2. Версии системы

Единый источник версии — три точки + документация:

| Точка | Рабочий проект | Копия |
|---|---|---|
| [`src/modules/Mod_Constants.bas`](../src/modules/Mod_Constants.bas:18) `APP_VERSION` | `"1.0.3"` | `"1.0.4"` |
| [`scripts/config.py`](../scripts/config.py:11) `APP_VERSION` | `"1.0.3"` | `"1.0.4"` |
| [`scripts/config.ps1`](../scripts/config.ps1:7) `$Script:AppVersion` | `"1.0.3"` | `"1.0.4"` |
| [`README.md`](../README.md:3) | `v1.0.3` | `v1.0.4` |
| [`docs/DEVELOPER.md`](../docs/DEVELOPER.md:1) | `v1.0.3` | `v1.0.4` |
| [`docs/ROADMAP.md`](../docs/ROADMAP.md:3) | `1.0.3` | `1.0.4` |
| [`docs/ARCHITECTURE.md`](../docs/ARCHITECTURE.md:4) | `v1.0.3` | `v1.0.4` |
| [`docs/CHANGELOG.md`](../docs/CHANGELOG.md) | последний релиз `[v1.0.3]` | последний релиз `[v1.0.4]` |

Механизм обновления версии — [`scripts/update_version.py`](../scripts/update_version.py) (идентичен в обоих проектах).

---

## 3. Методология сравнения

- Произведено сравнение структур каталогов обоих проектов (листинг файлов).
- Для файлов, присутствующих в обоих проектах, выполнено сопоставление содержимого: сверены первые строки, служебный заголовок «total lines» и, для ключевых файлов, ключевые фрагменты тела.
- Файлы с совпадающим количеством строк и совпадающим началом признаны идентичными; файлы с расхождением количества строк или начального блока — помечены как отличающиеся.
- Бинарные файлы (`*.xlsm`, `*.xlsx`) сравнивались только на уровне наличия/отсутствия в дереве каталогов.

---

## 4. Реестр отличий между копией (v1.0.4) и рабочим проектом (v1.0.3)

### 4.1. Каталоги/файлы, присутствующие в копии, но ОТСУТСТВУЮЩИЕ в рабочем проекте

| № | Путь (в копии) | Назначение | Действие при интеграции |
|---|---|---|---|
| 1 | `base/templates/work.xlsm` | Шаблон work с кодом | Создать каталог и файл |
| 2 | `base/templates/work0.xlsm` | Шаблон work пустой (без VBA) | Создать |
| 3 | `base/templates/model.xlsm` | Общий модельный шаблон с кодом | Создать |
| 4 | `base/templates/model0.xlsm` | Модельный шаблон пустой (без VBA) | Создать |
| 5 | `base/templates/report0.xlsx` | Шаблон отчёта пустой | Создать |
| 6 | `db/schema.sql` | DDL-схема SQLite `SysW.db` (v1) | Скопировать (**неоднозначность**, см. §6) |
| 7 | `plans/integration_1.0.4.md` | План интеграции v1.0.4 | Интегрировать |
| 8 | `plans/migration_sqlite_plan.md` | Архитектурный план миграции на SQLite | Интегрировать (**неоднозначность**) |
| 9 | `plans/debug_hang_templates_v104.md` | План исправления зависания `build_templates.py` | Интегрировать |
| 10 | `plans/integration_1.0.1.md` | План v1.0.1 (в копии лежит в корне `plans/`) | Решение по переносу |
| 11 | `plans/_promts/promt.md`, `promt_templates.md`, `promt0.20.0.md`, `promt1.md`, `promt2.md`, `promt3.md`, `promt4.md` | Промты пользователя | Решение по интеграции (**неоднозначность**) |
| 12 | `scripts/build_templates.py` | Создание шаблонов `base/templates/` | Скопировать |
| 13 | `scripts/apply_protection_templates.py` | Защита листов шаблонов | Скопировать |
| 14 | `scripts/template_protection.py` | Общая логика защиты листов | Скопировать |
| 15 | `scripts/sqlite_schema.py` | Версия схемы БД SQLite | Скопировать (**неоднозначность**) |
| 16 | `scripts/fix_vbom_and_venv.ps1` | Исправление VBAOM и venv | Скопировать |
| 17 | `scripts/_debug_*.py` (14 шт.) | Отладочные скрипты (`_debug_hang*`, `_debug_open*`, `_debug_isolate.py`, `_debug_trytrust.py`, `_debug_vbaonly.py`, `_debug_vis.py`) | Вероятно мусор ([S7]) — решение |
| 18 | `logs/log.txt`, `logs/log_old.txt` | Файлы логов | Исключены по `.codeassistantignore` (`*.log`) |

### 4.2. Каталоги/файлы, присутствующие в рабочем проекте, но ОТСУТСТВУЮЩИЕ в копии

| № | Путь (в рабочем проекте) | Комментарий |
|---|---|---|
| 1 | `workOt/` | В копии удалён (по указанию пользователя). В рабочем присутствует. Решение о судьбе — см. §6 |
| 2 | `plans/integration_1.0.3.md` | План невыполненных работ v1.0.3 (в копии отсутствует) |
| 3 | `plans/update_plans_and_docs_plan.md` | План актуализации планов/документации |
| 4 | `plans/_promts/promt7.md`, `promt8.md`, `promt_error.md` | Промты пользователя (актуальные в рабочем) |
| 5 | `plans/_archive/*` (часть) | Дополнительные архивные планы: `fix_consistency_discrepancies_plan.md`, `fix_deploy_path_plan.md`, `promt5.md`, `promt6.md`, `sourcecraft_integration_plan.md`, `update_plan_v0.15.0.md`, `promt*.md` и др. |

### 4.3. Файлы, присутствующие в обоих проектах и ОТЛИЧАЮЩИЕСЯ по содержимому

| № | Файл | Отличие копии от рабочего |
|---|---|---|
| 1 | `src/modules/Mod_Constants.bas` | Только `APP_VERSION`: `1.0.3` → `1.0.4` (остальное идентично, 426 строк) |
| 2 | `scripts/config.py` | `APP_VERSION` + добавлен блок SQLite: `DB_SCHEMA_VERSION`, `DB_PATH`, `DB_SCHEMA_PATH`, `MODELS_DIR`, `MIGRATION_REPORT_FILE` (**неоднозначность** — SQLite) |
| 3 | `scripts/config.ps1` | Только `$Script:AppVersion`: `1.0.3` → `1.0.4` |
| 4 | `scripts/check_docs.py` | 304 → 307 строк; в `main` добавлен вызов `check_expected_modules(files, issues)` (+ доп. правки) |
| 5 | `README.md` | Версия v1.0.4; структура дополнена `base/templates/`; НО содержит **регрессии**: путь `L:\PROject\SysW`, «3 класса листов», упоминает `CHANGELOG.md` в корне (**см. §6**) |
| 6 | `docs/CHANGELOG.md` | Добавлен раздел `[v1.0.4]`; версия v1.0.3 оформлена как историческая запись (без коммита) |
| 7 | `docs/DEVELOPER.md` | 894 → 969 строк; версия v1.0.4; добавлен раздел 8 «Шаблоны `base/templates/`» |
| 8 | `docs/ARCHITECTURE.md` | 787 → 805 строк; версия v1.0.4; добавлены `base/templates/` в структуру и подраздел «Защита листов шаблонов (v1.0.4)» |
| 9 | `docs/ROADMAP.md` | 327 → 329 строк; версия 1.0.4; статус актуализирован |
| 10 | `docs/table.md` | 202 → 650 строк; добавлен раздел 0 «Рабочая панель для шаблонов» (структура листов, зоны Protect/AllowEditRanges, кнопки) |
| 11 | `docs/sourcecraft-guide.md` | Упоминания v1.0.4, актуализация про `.sourcecraft` |
| 12 | `docs/git-workflow.md` | Незначительные отличия (версия не фиксируется) |

### 4.4. Файлы, идентичные в обоих проектах (сверено по длине и началу)

- **src/modules:** `Mod_AutoMatch.bas` (384), `Mod_Import.bas` (518), `Mod_Logger.bas` (137), `Mod_ModelDB.bas` (358), `Mod_ModelTypes.bas` (30), `Mod_ButtonDispatcher.bas` (193), `Mod_FullTestRunner.bas` (1583), `Mod_OrderHeader.bas` (280), `Mod_PickWork.bas` (150), `Mod_SheetButtons.bas` (141), `Mod_SheetOps.bas` (229), `Mod_Utils.bas` (143).
- **src/classes:** `PartIdentity.cls` (24), `WorkIdentity.cls` (22).
- **src/sheets:** `Лист2_main.cls` (47).
- **scripts:** `export_vba.py` (184), `impVBA.py` (443), `run_tests.py` (179), `update_version.py` (226), `Set-ExcelTrust.ps1` (222).
- **base/models/*:** совпадают по составу (4x4, 2170, 2180, 2190, GAZ, UAZ, .gitkeep).
- **Конфигурационные:** `.gitattributes`, `.sourcecraft`, `.codeassistantignore` — предполагаются идентичными (требуют подтверждения).

> **Примечание:** VBA-логика в v1.0.4 не менялась (подтверждено [`plans/integration_1.0.4.md`](../plans/_promts/../integration_1.0.4.md) и CHANGELOG: «Реализовано на уровне файлов-шаблонов **без изменения VBA-кода**»). Все проверенные модули совпадают по количеству строк и началу.

---

## 5. Реестр планируемых изменений для обновления до 1.0.4

> Изменения сгруппированы по областям. Критические файлы ([E3]) — только после согласования.

### A. Версия системы (до 1.0.4)
1. `src/modules/Mod_Constants.bas` — `APP_VERSION = "1.0.4"`.
2. `scripts/config.py` — `APP_VERSION = "1.0.4"`.
3. `scripts/config.ps1` — `$Script:AppVersion = "1.0.4"`.
4. `README.md` — `Версия: v1.0.4`.
5. `docs/DEVELOPER.md`, `docs/ROADMAP.md`, `docs/ARCHITECTURE.md` — обновление версии.
6. `docs/CHANGELOG.md` — раздел `## [v1.0.4]`.

*Рекомендуемый способ:* `python scripts/update_version.py 1.0.4` (+ ручное дополнение CHANGELOG).

### B. Шаблоны `base/templates/`
7. Создать каталог `base/templates/` и 5 файлов-шаблонов (из копии или пересоздать через `build_templates.py`).
8. Перенести скрипты: `build_templates.py`, `apply_protection_templates.py`, `template_protection.py`.
9. Применить защиту листов (`Protect` + `AllowEditRanges` + `FreezePanes A4`) — уже заложена в файлах копии.

### C. Скрипты
10. `scripts/check_docs.py` — интегрировать изменения из копии (добавлен `check_expected_modules`).
11. `scripts/fix_vbom_and_venv.ps1` — добавить (новый скрипт).
12. `scripts/config.py` — блок SQLite (решение см. §6).

### D. SQLite (по плану `migration_sqlite_plan.md`)
13. `db/schema.sql` — добавить (решение см. §6).
14. `scripts/sqlite_schema.py` — добавить (решение см. §6).

### E. Документация
15. `docs/table.md` — перенести расширение из копии (раздел 0 о шаблонах/защите).
16. `docs/DEVELOPER.md`, `docs/ARCHITECTURE.md`, `docs/ROADMAP.md`, `docs/sourcecraft-guide.md` — актуализировать до v1.0.4 (шаблоны, защита).
17. `README.md` — согласовать формат слияния (см. §6, вопрос о регрессиях).
18. `docs/CHANGELOG.md` — раздел v1.0.4.

### F. Планы
19. Интегрировать планы из копии: `integration_1.0.4.md`, `migration_sqlite_plan.md`, `debug_hang_templates_v104.md`.
20. Решить судьбу рабочих планов `integration_1.0.3.md`, `update_plans_and_docs_plan.md`.

### G. Прочее
21. `workOt/` — удалить (по решению пользователя, как в копии).
22. Отладочные `scripts/_debug_*.py` — НЕ переносить (мусор [S7]).

### H. Логи
23. `logs/log.txt`, `logs/log_old.txt` — НЕ переносить (исключены `.codeassistantignore`: `*.log`; в рабочем остаётся только `logs/.gitkeep`).

---

## 5.1. Принятые решения пользователя (согласовано 2026-08-21)

1. **Интеграция** — выполняется **полная интеграция из копии**: шаблоны `base/templates/`, защита листов, версия до 1.0.4, SQLite-инфраструктура (`db/schema.sql`, `scripts/sqlite_schema.py`, константы `DB_*` в `config.py`), все планы и скрипты из копии.
2. **Документация** — сохранять актуальные правки рабочего проекта (относительные пути, фактическое число классов листов = 1) и **добавлять только новые разделы из копии** (шаблоны, защита листов, версия v1.0.4). Регрессии копии (`L:\PROject\SysW`, «3 класса листов») не переносятся.
3. **`workOt/`** — удалить в рабочем проекте.
4. **`scripts/_debug_*.py`** — не переносить (временные отладочные скрипты).
5. **Бинарные шаблоны `base/templates/*`** — скопировать готовые файлы напрямую из `F:/MiTyA/PROject/copy/SysW/base/templates/` (без пересоздания через `build_templates.py`).
6. **Промты `plans/_promts/`** — файлы из копии (`promt.md`, `promt_templates.md`, `promt0.20.0.md`, `promt1.md`, `promt2.md`, `promt3.md`, `promt4.md`) добавить в `plans/_promts/`; рабочие `promt7.md`, `promt8.md` — оставить.
7. **Логи `logs/*.log`** — не переносить.

---

## 6. Статус согласования неоднозначностей (решено)

Все неоднозначности решены пользователем (см. §5.1). Краткая сводка решений по каждому пункту анализа:

1. **SQLite-инфраструктура** — интегрировать полностью (`db/schema.sql`, `scripts/sqlite_schema.py`, константы `DB_*` в `config.py`, план `migration_sqlite_plan.md`).
2. **Документация README/ARCHITECTURE** — сохранять актуальные правки рабочего + добавлять новые разделы из копии (регрессии копии не переносятся).
3. **`workOt/`** — удалить в рабочем проекте (по явному решению пользователя).
4. **`scripts/_debug_*.py`** — не переносить (мусор).
5. **Шаблоны `base/templates/*`** — скопировать готовые бинарные файлы напрямую из копии (без пересоздания скриптом).
6. **Промты `plans/_promts/`** — добавить файлы из копии, рабочие `promt7.md`/`promt8.md` оставить.
7. **`plans/integration_1.0.1.md`** — оставить в `plans/_archive/` (исторический план; в корне `plans/` не дублировать).
8. **Логи `logs/*.log`** — не переносить (исключены `.codeassistantignore`).

---

## 7. Порядок дальнейших действий

> Реестр согласован (см. §5.1, §6). Неоднозначностей не осталось.

1. Подготовить детальный план реализации (в режиме Code/Architect) на основе утверждённого реестра.
2. Реализовать изменения (полная интеграция из копии, зафиксированные исключения).
3. Обновить `docs/CHANGELOG.md` разделом v1.0.4.
4. Проверить чистоту проекта ([S7]).
5. Выполнить коммит и пуш через MCP Git Tools ([G10]).