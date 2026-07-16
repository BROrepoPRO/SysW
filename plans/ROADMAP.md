# План развития системы SysW (Roadmap)

> **Версия системы:** 0.7.1
> **Дата:** 2026-07-16
> **Статус:** Актуализирован (Фаза 1 выполнена, аудит стабильности проведён)

---

## 1. Текущее состояние

Система SysW (v0.7.1) — модульная VBA-надстройка для Excel с обвязкой на Python/PowerShell, автоматизирующая обработку заказ-нарядов авторемонта (импорт данных из отчётов, заполнение шапки заказа, поиск по ГРЗ, учёт работ и запчастей).

### Что реализовано

| Область | Статус | Подробности |
|---------|--------|-------------|
| **Структура проекта** | ✅ v0.6.0 | `src/modules/` (10 `.bas`), `src/sheets/` (3 `.cls`), `scripts/`, `docs/`, `plans/` |
| **Тестирование** | ✅ v0.2.0, v0.7.0 | `Mod_FullTestRunner` — 30 тестов (TC-01..TC-30), покрытие ~77% |
| **Техническая документация** | ✅ v0.5.0 | `docs/DEVELOPER.md` — архитектура, кодировка, CI/CD, тестирование |
| **CI/CD** | ✅ v0.4.0 | GitHub Actions (`vba-check.yml`), `docs/git-workflow.md` |
| **SourceCraft интеграция** | ✅ v0.3.0 | `.ycarules`, `docs/sourcecraft-guide.md` |
| **Двухфазная кодировка** | ✅ | UTF-8 на диске, CP1251 в Excel |
| **Скрипты экспорта/импорта** | ✅ | `export_vba.py`, `impVBA.py`, `Import-VbaFromExcel.ps1` |
| **README.md** | ✅ | Архитектура, быстрый старт, состав команды |
| **CHANGELOG.md** | ✅ | История версий в формате Keep a Changelog |
| **Фаза 1 — Критические исправления (P0)** | ✅ Выполнена | R-01..R-06: исправление кодировки, `export_vba.py`, утечки, `isProcessing`, диапазоны, дублирование |
| **Аудит стабильности v0.7.1** | ✅ Выполнен | Проверено 13 модулей, найдено 26 проблем (4 Critical, 7 High, 10 Medium, 5 Low), исправлено 11 проблем (4 Critical + 7 High) |

### Модульная архитектура VBA

```
UI-слой:              Sheet1_main.cls  |  Sheet_work.cls  |  Sheet_z4.cls
                      Mod_ButtonDispatcher  |  Mod_MainButtons  |  Mod_SheetButtons
                           │                          │
Бизнес-логика:       Mod_OrderHeader  |  Mod_Import  |  Mod_SheetOps
                           │                          │
Утилиты:             Mod_Utils  |  Mod_Constants  |  Mod_Logger
                           │
Тестирование:        Mod_FullTestRunner
```

---

## 2. Архитектура и рефакторинг

### ✅ Выполнено (Фаза 1 — Критические исправления, P0)

> **Цель:** устранены проблемы, блокирующие работу или ведущие к потере данных.
> **Статус:** ✅ ВСЕ ЗАДАЧИ ВЫПОЛНЕНЫ

| ID | Задача | Описание | Затрагиваемые модули/файлы | Статус |
|----|--------|----------|---------------------------|--------|
| R-01 | Исправление кодировки в VBA-модулях | Проверить и исправить кракозябры в комментариях и строках. Пересохранить файлы в UTF-8 без BOM через `export_vba.py` | [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas), [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas), [`Mod_SheetOps.bas`](../src/modules/Mod_SheetOps.bas), [`Mod_Import.bas`](../src/modules/Mod_Import.bas), [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas) | [ВЫПОЛНЕНО] |
| R-02 | Исправление `export_vba.py` | Корректный экспорт в UTF-8 без BOM, обработка пустых строк, добавление недостающих модулей (`Mod_SheetButtons`, `Mod_MainButtons`, `Sheet_z4`, `Sheet_work`) в `COMPONENTS` | [`scripts/export_vba.py`](../scripts/export_vba.py) | [ВЫПОЛНЕНО] |
| R-03 | Устранение утечки обработчика | Освобождение обработчика (`wbReport.Close`) во всех модулях при ошибках. Проверить `SearchSheetByGRZ` и другие функции, открывающие внешние книги | [`Mod_Import.bas`](../src/modules/Mod_Import.bas) | [ВЫПОЛНЕНО] |
| R-04 | Замена `isProcessing` на модульный флаг | Убрать глобальный статический флаг `isProcessing` из `Sheet1_main.cls`, сделать локальный для каждого модуля с обработкой ошибок (`On Error GoTo ErrHandler`) | [`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls) | [ВЫПОЛНЕНО] |
| R-05 | Диапазон очистки `Mod_SheetOps` | Исправить обработку граничных значений: заменить `"A2:XFD"` на `"B2:ZZ"` в `ClearMainSheet_UI` и `Btn_main_Import` | [`Mod_Import.bas`](../src/modules/Mod_Import.bas), [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas) | [ВЫПОЛНЕНО] |
| R-06 | Устранение дублирования кода | Вынести повторяющиеся блоки в `Mod_Utils`. Устранить дублирование `Btn_main_Import` (Mod_MainButtons) → делегировать `Mod_Import.ImportSheet` | [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas), [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas) | [ВЫПОЛНЕНО] |

### Фаза 1 — Критические архитектурные изменения (P0, ~18 ч)

> **Цель:** рефакторинг архитектуры — выделение новых модулей, унификация, устранение дублирования.
> **Риск:** средний. Требуется тестирование после каждого изменения.

| ID | Задача | Описание | Затрагиваемые модули/файлы | Критерии готовности |
|----|--------|----------|---------------------------|---------------------|
| R-07 | Выделение `Mod_Logger` в отдельный модуль | Централизованное логирование с ротацией вместо `Mod_Utils.WriteLog`. Реализовать: `WriteLog(Message, Level)`, `SetLogFile(Path)`, `SetMaxSize(Bytes)`, автоматическую ротацию | [`Mod_Logger.bas`](../src/modules/Mod_Logger.bas) (создать/доработать), [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas) (удалить `WriteLog`) | Логи пишутся с уровнями; ротация работает при превышении размера |
| R-08 | Замена `MsgBox` на Logger | Во всех модулях заменить прямые вызовы `MsgBox` на `Mod_Logger`. Создать `_UI`-обёртки для пользовательских диалогов | Все модули | `MsgBox` остаётся только в `_UI`-обёртках; чистые функции не содержат диалогов |
| R-09 | Рефакторинг `Mod_OrderHeader` | Разделение на подфункции, уменьшение связанности. Убрать `MsgBox` из `FillHeaderFromOrder`, вынести в `_UI`-обёртку | [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas), [`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls) | `FillHeaderFromOrder` — чистая функция; `_UI`-обёртка в диспетчере |
| R-10 | Рефакторинг `Mod_Constants` | Вынести магические строки/числа в константы, структурировать по группам. Перенести тип `OrderHeader` из `Mod_Utils`. Создать единый словарь маппинга столбцов | [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas), [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas), [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas), [`Mod_Import.bas`](../src/modules/Mod_Import.bas) | Все константы в одном модуле; магические числа заменены; компиляция проходит |
| R-11 | Маппинг листов | Создать единый словарь маппинга листов вместо разрозненных ссылок по всему коду | [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas), все модули, использующие имена листов | Единая точка изменений для имён листов |
| R-12 | Рефакторинг `Mod_SheetOps` | Разделить на логические блоки, уменьшить размер. Перенести операции с листами из `Mod_Import` | [`Mod_SheetOps.bas`](../src/modules/Mod_SheetOps.bas), [`Mod_Import.bas`](../src/modules/Mod_Import.bas), [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) | `Mod_Import` сокращён; все операции с листами в `Mod_SheetOps` |
| R-13 | Унификация кнопок | `Mod_ButtonDispatcher`, `Mod_MainButtons`, `Mod_SheetButtons` — единый интерфейс. Унифицировать нейминг: `Btn_<лист>_<действие>_Click`. Удалить дублирующиеся обработчики | [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas), [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas), [`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas) | Все 13+ кнопок работают; единый стиль именования; нет дублирования |

### Фаза 2 — Высокий приоритет (P1, ~5.5 ч)

> **Цель:** улучшение качества кода, устранение косметических проблем, подготовка к будущему росту.
> **Риск:** низкий. Изменения изолированные.

| ID | Задача | Описание | Затрагиваемые модули/файлы | Критерии готовности |
|----|--------|----------|---------------------------|---------------------|
| R-14 | Выделение `Mod_Selection` | Модуль работы с выделенным диапазоном. Перенести заглушки из `Mod_MainButtons` и `Mod_SheetButtons`. Реализовать базовую логику подбора запчастей и работ | [`Mod_Selection.bas`](../src/modules/Mod_Selection.bas) (создать), [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas), [`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas), [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) | Заглушки перенесены; кнопки подбора не падают с ошибкой |
| R-15 | Объединение листов | Унификация `Sheet_work`, `Sheet_z4`, `Sheet1_main`. Вынести общую логику в `Mod_SheetCommon` или признать дубль допустимым с синхронизацией комментариев | [`Sheet_work.cls`](../src/sheets/Sheet_work.cls), [`Sheet_z4.cls`](../src/sheets/Sheet_z4.cls), [`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls) | Нет дублирования кода событий; или дубль явно задокументирован |
| R-16 | Расширение `run_tests.py` | Добавить аргументы командной строки (`--module`, `--verbose`, `--output`), генерацию отчётов (JSON/HTML), поддержку выборочного запуска тестов. **Не путать с уже выполненной локализацией вывода (транслит → русский) и добавлением `GetTestResults()` — это следующий шаг развития скрипта** | [`scripts/run_tests.py`](../scripts/run_tests.py) | `run_tests.py --help` показывает все аргументы; отчёты сохраняются |
| R-17 | Обновление документации после архитектурного рефакторинга | После завершения Фазы 1 (R-07..R-13) обновить `DEVELOPER.md` и `README.md` — актуализировать схемы зависимостей, таблицы модулей, маппинг файлов под новую архитектуру. **Не относится к уже выполненному обновлению документации в v0.7.0 (раздел тестирования TC-01..TC-30)** | [`docs/DEVELOPER.md`](../docs/DEVELOPER.md), [`README.md`](../README.md) | Вся документация соответствует актуальной архитектуре после рефакторинга |

---

## 3. Функциональность

### Фаза 3 — Долгосрочные улучшения (P2, опционально)

> **Цель:** масштабирование для больших справочников (10000+ записей).
> **Риск:** высокий. Требует изменений в архитектуре импорта.

| ID | Задача | Описание | Затрагиваемые модули/файлы | Критерии готовности |
|----|--------|----------|---------------------------|---------------------|
| R-18 | Внешние справочники | Выгрузка справочников в отдельные файлы (CSV/JSON). Создать `справочник_запчастей.xlsx` и `справочник_работ.xlsx`, открываются ReadOnly | [`Mod_Selection.bas`](../src/modules/Mod_Selection.bas), новые файлы справочников | Подбор работает с внешними справочниками; поиск через AutoFilter |
| R-19 | SQLite | Замена Excel-справочников на SQLite через ADO. Индексация по VIN/модели/ГРЗ для быстрого поиска среди 700000+ записей | [`Mod_Selection.bas`](../src/modules/Mod_Selection.bas), `SysW.db` | Сложные запросы (JOIN, WHERE, ORDER BY) через SQLite; производительность выше циклов VBA |
| R-20 | CI/CD расширение | Pre-commit hook (проверка кодировки UTF-8 без BOM), автоматические прогоны тестов в GitHub Actions при пуше | [`.github/workflows/vba-check.yml`](../.github/workflows/vba-check.yml), новый `.husky/pre-commit` | Pre-commit проверяет кодировку; CI прогоняет тесты (с SKIP на Linux/MacOS) |
| R-21 | Автоматизация экспорта/импорта | Улучшение скриптов: обработка ошибок, `--dry` режим, логирование операций, параллельный экспорт | [`scripts/export_vba.py`](../scripts/export_vba.py), [`scripts/impVBA.py`](../scripts/impVBA.py), [`scripts/Import-VbaFromExcel.ps1`](../scripts/Import-VbaFromExcel.ps1) | Все скрипты имеют `--dry` режим; ошибки обрабатываются; есть логи операций |

### Фаза 4 — Низкий приоритет (P3)

> **Цель:** концептуальные улучшения, не имеющие срочности.
> **Риск:** варьируется.

| ID | Задача | Описание | Потенциальные модули |
|----|--------|----------|---------------------|
| R-22 | Интеграция с GitHub Projects | Автоматическое создание issues при добавлении задач в ROADMAP. Связь коммитов с issues через Conventional Commits | `.github/workflows/`, `plans/` |
| R-23 | Генерация документации | Autodoc по VBA-комментариям. Извлечение документации из `'''` комментариев в Markdown | Новый скрипт `scripts/generate_docs.py` |
| R-24 | Веб-интерфейс | Просмотр логов и статусов через веб-страницу. Flask/FastAPI бэкенд, чтение логов `Mod_Logger` | Новый модуль `web/` |
| R-25 | Поддержка нескольких языков | Локализация интерфейса (русский/английский/казахский). Вынос строк в ресурсные файлы | `Mod_Constants.bas`, новые `.res` файлы |

---

## 4. Тестирование и CI/CD

### Текущее состояние

- **30 тестов** (TC-01..TC-30) в [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas)
- **Покрытие:** ~77% (26 активных тестов: 20 PASS, 5 SKIP, 0 FAIL, 2 в резерве)
- **Запуск:** `python scripts/run_tests.py` (COM-автоматизация Excel)
- **CI/CD:** GitHub Actions — проверка наличия файлов, кодировки UTF-8, синтаксиса VBA, актуальности CHANGELOG

### План расширения

| ID | Задача | Приоритет | Описание |
|----|--------|-----------|----------|
| R-26 | TC-31..TC-35: Unit-тесты для `Mod_Constants` | P2 | Проверка всех констант и маппингов |
| R-27 | TC-36..TC-40: Unit-тесты для `Mod_Logger` (ротация, уровни) | P2 | Проверка ротации при превышении размера, корректности уровней |
| R-28 | TC-41..TC-45: Интеграционные тесты для `Mod_SheetOps` | P2 | Проверка создания/удаления/копирования листов |
| R-29 | Pre-commit hook: проверка UTF-8 | P3 | Автоматическая проверка кодировки `.bas`/`.cls` файлов перед коммитом |
| R-30 | CI: автоматический прогон тестов | P3 | Добавление `run_tests.py` в GitHub Actions (с SKIP на Linux/MacOS) |
| R-31 | CI: проверка покрытия | P4 | Генерация отчёта о покрытии и сравнение с порогом (цель: >85%) |

---

## 5. Документация

### Текущее состояние

| Документ | Статус | Описание |
|----------|--------|----------|
| [`README.md`](../README.md) | ✅ Актуально | Общее описание, архитектура, быстрый старт, состав команды |
| [`docs/DEVELOPER.md`](../docs/DEVELOPER.md) | ✅ Актуально (v0.7.0) | Техническая документация: архитектура, кодировка, CI/CD, тестирование |
| [`docs/git-workflow.md`](../docs/git-workflow.md) | ✅ Актуально | Веточная стратегия, Conventional Commits, pre-commit процедуры |
| [`docs/sourcecraft-guide.md`](../docs/sourcecraft-guide.md) | ✅ Актуально | Руководство по SourceCraft, роли, процессы |
| [`docs/MODERNIZATION_CHECKLIST.md`](../docs/MODERNIZATION_CHECKLIST.md) | 🟡 Частично | Чек-лист модернизации (некоторые пункты выполнены в v0.7.0) |
| [`plans/architecture_modernization_proposal.md`](../plans/architecture_modernization_proposal.md) | 🟡 Черновик | Предложение по модернизации архитектуры |
| [`plans/architecture_migration_phase3.md`](../plans/architecture_migration_phase3.md) | ✅ Выполнен | План реструктуризации каталогов (v0.6.0) |
| [`plans/ROADMAP.md`](../plans/ROADMAP.md) | 🆕 Настоящий документ | Единый план развития системы |

### План обновления

| ID | Задача | Приоритет | Описание |
|----|--------|-----------|----------|
| R-32 | Обновление `DEVELOPER.md` после Фазы 1 | P2 | Обновить схемы зависимостей, таблицы модулей, маппинг файлов после рефакторинга |
| R-33 | Обновление `README.md` после Фазы 1 | P2 | Синхронизировать описание архитектуры с новыми модулями |
| R-34 | Создание `docs/developer-guide.md` | P3 | Руководство для новых разработчиков: как добавить модуль, кнопку, тест |
| R-35 | Обновление `docs/sourcecraft-guide.md` | P3 | Актуализация схемы структуры проекта и ссылок |
| R-36 | Архивирование устаревших планов | P3 | Перенос выполненных планов в `plans/_archive/` |

---

## 6. Интеграция с SourceCraft/GitHub

### Текущее состояние

- [`.ycarules`](../.ycarules) — 6 групп правил (G, Z, K, U, S, O)
- [`docs/sourcecraft-guide.md`](../docs/sourcecraft-guide.md) — руководство по работе с SourceCraft
- GitHub Actions — базовый CI/CD workflow

### План улучшения

| ID | Задача | Приоритет | Описание |
|----|--------|-----------|----------|
| R-37 | Улучшение `.ycarules` | P2 | Добавить правила для новых модулей (`Mod_Selection`), обновить правило `[S1]` |
| R-38 | Скрипты автоматизации SourceCraft | P3 | Создать скрипты для автоматической генерации планов, проверки соответствия `.ycarules` |
| R-39 | GitHub Projects интеграция | P4 | Автоматическое создание issues из задач ROADMAP; связь коммитов с issues |
| R-40 | Метрики и отчёты | P4 | Генерация отчётов о прогрессе по фазам; дашборд статуса задач |

---

## 7. Сводная информация

### Сводка по фазам

| Фаза | Приоритет | Часы | Задач | Статус |
|------|-----------|------|-------|--------|
| Фаза 1 (P0) — Критические исправления | P0 | ~3.5 | 6 (R-01..R-06) | ✅ Выполнена |
| Фаза 1 — Критические архитектурные изменения | P0 | ~18.0 | 7 (R-07..R-13) | ⬜ Не начата |
| Фаза 2 — Высокий приоритет | P1 | ~5.5 | 4 (R-14..R-17) | ⬜ Не начата |
| Фаза 3 — Долгосрочные улучшения | P2 | — | 4 (R-18..R-21) | ⬜ Опционально |
| Фаза 4 — Низкий приоритет | P3 | — | 4 (R-22..R-25) | 💡 Концепции |
| Тестирование и CI/CD | P2-P4 | — | 6 (R-26..R-31) | ⬜ План |
| Документация | P2-P3 | — | 5 (R-32..R-36) | ⬜ План |
| SourceCraft/GitHub | P2-P4 | — | 4 (R-37..R-40) | ⬜ План |
| **Аудит стабильности v0.7.1** | **P0** | **~4.0** | **11 проблем** | **✅ Выполнен** |
| **Итого** | **P0-P4** | **~27.5** | **40 + аудит** | **⬜ ~10%** |

### Распределение по модулям

| Модуль/файл | Затрагивается задачами |
|-------------|----------------------|
| [`Mod_Utils.bas`](../src/modules/Mod_Utils.bas) | R-06, R-07, R-08, R-10 |
| [`Mod_OrderHeader.bas`](../src/modules/Mod_OrderHeader.bas) | R-01, R-09, R-10 |
| [`Mod_Import.bas`](../src/modules/Mod_Import.bas) | R-01, R-03, R-05, R-10, R-12 |
| [`Mod_ButtonDispatcher.bas`](../src/modules/Mod_ButtonDispatcher.bas) | R-12, R-13, R-14 |
| [`Mod_FullTestRunner.bas`](../src/modules/Mod_FullTestRunner.bas) | R-01 |
| [`Mod_Logger.bas`](../src/modules/Mod_Logger.bas) | R-07, R-08 |
| [`Mod_Constants.bas`](../src/modules/Mod_Constants.bas) | R-01, R-10, R-11 |
| [`Mod_SheetOps.bas`](../src/modules/Mod_SheetOps.bas) | R-01, R-05, R-12 |
| [`Mod_MainButtons.bas`](../src/modules/Mod_MainButtons.bas) | R-05, R-06, R-13, R-14 |
| [`Mod_SheetButtons.bas`](../src/modules/Mod_SheetButtons.bas) | R-13, R-14 |
| **`Mod_Selection.bas`** (новый) | R-14, R-18, R-19 |
| [`Sheet1_main.cls`](../src/sheets/Sheet1_main.cls) | R-04, R-09, R-15 |
| [`Sheet_work.cls`](../src/sheets/Sheet_work.cls) | R-15 |
| [`Sheet_z4.cls`](../src/sheets/Sheet_z4.cls) | R-15 |
| [`scripts/export_vba.py`](../scripts/export_vba.py) | R-02, R-21 |
| [`scripts/impVBA.py`](../scripts/impVBA.py) | R-21 |
| [`scripts/run_tests.py`](../scripts/run_tests.py) | R-16 |
| [`scripts/Import-VbaFromExcel.ps1`](../scripts/Import-VbaFromExcel.ps1) | R-21 |
| [`docs/DEVELOPER.md`](../docs/DEVELOPER.md) | R-17, R-32 |
| [`README.md`](../README.md) | R-17, R-33 |
| [`.ycarules`](../.ycarules) | R-37 |
| [`.github/workflows/vba-check.yml`](../.github/workflows/vba-check.yml) | R-20, R-30 |

---

## 8. Принципы и ограничения

### Архитектурные принципы

1. **SRP (Single Responsibility Principle)** — каждый модуль отвечает за одну зону ответственности
2. **Слабая связанность** — модули общаются через вызовы функций, без глобальных переменных
3. **UI-обёртки** — функции с суффиксом `_UI` содержат диалоги (`MsgBox`), чистые функции — нет
4. **Единая точка входа** — все кнопки проходят через `Mod_ButtonDispatcher`
5. **Тестируемость** — чистые функции можно тестировать без Excel
6. **Масштабируемость** — внешние справочники для больших объёмов данных

### Технические ограничения

| Ограничение | Описание |
|-------------|----------|
| **Двухфазная кодировка** | VBA-файлы на диске — UTF-8 (без BOM), в Excel — CP1251. Конвертация через `export_vba.py` / `impVBA.py`. Запрещено менять кодировку вручную ([правило K1-K3](../.ycarules:65)) |
| **Conventional Commits** | Формат коммитов: `тип(область): описание`. Типы: `feat`, `fix`, `refactor`, `docs`, `chore`, `test`, `ci` |
| **Веточная стратегия** | `main` — стабильная, `dev` — разработка, feature-ветки от `dev`. Запрещён прямой push в `main` |
| **Запрет кириллицы в именах** | Имена файлов, папок, модулей, процедур, переменных — только латиница ([правило Z3](../.ycarules:50)) |
| **Запрет самовольных правок VBA** | Изменения в `.bas`/`.cls`/`.frm` — только после создания и утверждения плана ([правило Z5](../.ycarules:55)) |
| **CHANGELOG** | При любых изменениях обновлять `CHANGELOG.md` в формате Keep a Changelog ([правило U2](../.ycarules:89)) |
| **PowerShell-скрипты** | Должны быть в UTF-8 with BOM (требование PowerShell для кириллицы) |

### Ссылки

| Ресурс | Ссылка |
|--------|--------|
| Правила SourceCraft | [`.ycarules`](../.ycarules) |
| Техническая документация | [`docs/DEVELOPER.md`](../docs/DEVELOPER.md) |
| Git-инструкции | [`docs/git-workflow.md`](../docs/git-workflow.md) |
| Руководство SourceCraft | [`docs/sourcecraft-guide.md`](../docs/sourcecraft-guide.md) |
| Предложение по модернизации | [`plans/architecture_modernization_proposal.md`](../plans/architecture_modernization_proposal.md) |
| Чек-лист модернизации | [`docs/MODERNIZATION_CHECKLIST.md`](../docs/MODERNIZATION_CHECKLIST.md) |
| История версий | [`CHANGELOG.md`](../CHANGELOG.md) |
| Общее описание | [`README.md`](../README.md) |

---

*Документ создан на основе анализа: [`plans/architecture_modernization_proposal.md`](../plans/architecture_modernization_proposal.md), [`plans/architecture_migration_phase3.md`](../plans/architecture_migration_phase3.md), [`docs/MODERNIZATION_CHECKLIST.md`](../docs/MODERNIZATION_CHECKLIST.md), [`docs/DEVELOPER.md`](../docs/DEVELOPER.md), [`docs/sourcecraft-guide.md`](../docs/sourcecraft-guide.md), [`CHANGELOG.md`](../CHANGELOG.md), [`README.md`](../README.md), [`.ycarules`](../.ycarules)*