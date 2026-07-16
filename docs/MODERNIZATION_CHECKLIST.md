# План-чек-лист модернизации SysW

> **Общий бюджет:** ~27 человеко-часов (~3.5 рабочих дня)
> **Статус:** 🟡 Ожидает утверждения

---

## Как пользоваться чек-листом

- ✅ — задача выполнена
- ⬜ — задача ожидает выполнения
- 🔄 — задача в процессе
- ❌ — задача заблокирована

---

## Фаза 0 — Подготовка (0.5 ч)

- [x] ✅ **Сделать бэкап `work.xlsm`** — скопировать файл в `_backup/work_YYYY-MM-DD.xlsm`
- [x] ✅ **Сделать бэкап `report.xlsx`** — скопировать файл в `_backup/report_YYYY-MM-DD.xlsx`
- [x] ✅ **Запустить `export_vba.py`** — убедиться, что текущее состояние VBA-модулей экспортировано на диск
- [x] ✅ **Закоммитить текущее состояние** — `git add -A && git commit -m "chore: backup before modernization phase 1"`

---

## Фаза 1 — Критические исправления (3.5 ч) ✅ ВЫПОЛНЕНА

> **Цель:** устранены проблемы, которые блокировали работу или могли привести к потере данных.
> **Риск:** низкий. Изменения точечные, не затрагивают архитектуру.

### 1.1. Исправить битую кодировку в `Mod_FullTestRunner.bas` (0.5 ч)

- [x] ✅ Открыть `work.xlsm` в Excel
- [x] ✅ Открыть редактор VBA (Alt+F11)
- [x] ✅ Найти модуль `Mod_FullTestRunner`
- [x] ✅ Скопировать весь текст модуля
- [x] ✅ Вставить в файл `Mod_FullTestRunner.bas` (UTF-8 без BOM)
- [x] ✅ Проверить, что русские комментарии отображаются корректно
- [x] ✅ Запустить `python run_tests.py` — убедиться, что тесты проходят

### 1.2. Добавить 4 модуля в скрипты экспорта/импорта (0.5 ч)

- [x] ✅ Открыть `export_vba.py`
- [x] ✅ Добавить в список `COMPONENTS`:
  - `Mod_SheetButtons.bas`
  - `Mod_MainButtons.bas`
  - `Sheet_z4.cls`
  - `Sheet_work.cls`
- [x] ✅ Открыть `impVBA.py`
- [x] ✅ Добавить в список `FILES` те же 4 файла
- [x] ✅ Запустить `python export_vba.py` — убедиться, что модули экспортируются
- [x] ✅ Проверить `git status` — файлы должны появиться в изменениях

### 1.3. Исправить утечку ресурсов в `SearchSheetByGRZ` (0.5 ч)

- [x] ✅ Открыть `Mod_Import.bas`
- [x] ✅ Найти функцию `SearchSheetByGRZ` (строка ~47)
- [x] ✅ Добавить `wbReport.Close False` перед `Exit Function` (строка ~79)
- [x] ✅ Убедиться, что книга закрывается и при нормальном выходе, и при ошибке
- [x] ✅ Проверить, что `Btn_main_Import` в `Mod_MainButtons.bas` не использует открытую книгу после вызова `SearchSheetByGRZ`

### 1.4. Исправить `isProcessing` в `Sheet1_main.cls` (0.5 ч)

- [x] ✅ Открыть `Sheet1_main.cls`
- [x] ✅ Найти обработчик `Worksheet_Change`
- [x] ✅ Добавить обработку ошибок, чтобы `isProcessing` сбрасывался в `False` при любом исходе:
  ```vba
  On Error GoTo ErrHandler
  ' ... основной код ...
  isProcessing = False
  Exit Sub
  ErrHandler:
      isProcessing = False
      MsgBox "Ошибка: " & Err.Description
  ```
- [x] ✅ Проверить, что при ошибке в `FillHeaderFromOrder` флаг сбрасывается

### 1.5. Заменить диапазон `A2:XFD` на `B2:ZZ` (0.5 ч)

- [x] ✅ Открыть `Mod_Import.bas`
- [x] ✅ Найти `ClearMainSheet_UI` (строка ~253)
- [x] ✅ Заменить `"A2:XFD" & lastRow` на `"B2:ZZ" & lastRow`
- [x] ✅ Открыть `Mod_MainButtons.bas`
- [x] ✅ Найти `Btn_main_Import` (строка ~50)
- [x] ✅ Убедиться, что там тоже используется `"B2:ZZ"` (а не `"A2:XFD"`)

### 1.6. Устранить дублирование `Btn_main_Import` (1.5 ч)

- [x] ✅ Открыть `Mod_MainButtons.bas`
- [x] ✅ Найти `Btn_main_Import` (строка ~81)
- [x] ✅ Сравнить логику с `Mod_Import.ImportSheet` (строка ~131)
- [x] ✅ Удалить дублирующийся код из `Btn_main_Import`
- [x] ✅ Заменить на вызов `Mod_Import.ImportSheet`
- [x] ✅ Проверить, что все 8 этапов импорта сохранились
- [x] ✅ Протестировать кнопку импорта на листе main

### ✅ Фаза 1 — Проверка

- [x] ✅ Запустить `python run_tests.py` — все тесты проходят
- [x] ✅ Проверить импорт из `report.xlsx` — данные корректно загружаются
- [x] ✅ Проверить очистку листа — очищается только B2:ZZ, не A1
- [x] ✅ Проверить заполнение шапки — B3:B15 заполняется корректно
- [x] ✅ Закоммитить: `git add -A && git commit -m "feat: phase 1 — critical fixes (encoding, export, leak, dedup)"`

---

## Фаза 2 — Архитектурные изменения (18 ч)

> **Цель:** рефакторинг архитектуры — выделение новых модулей, унификация, устранение дублирования.
> **Риск:** средний. Требуется тестирование после каждого изменения.

### 2.1. Создать `Mod_Logger.bas` — логирование с ротацией (3 ч)

- [ ] ⬜ Создать файл `Mod_Logger.bas`
- [ ] ⬜ Реализовать:
  - `WriteLog(Message, Level)` — запись с уровнем (INFO/WARN/ERROR)
  - `SetLogFile(Path)` — настройка пути к файлу лога
  - `SetMaxSize(Bytes)` — ротация по размеру
  - Автоматическая ротация: при превышении размера — переименовать в `log_YYYYMMDD_HHMMSS.txt`
- [ ] ⬜ Заменить все вызовы `Mod_Utils.WriteLog` на `Mod_Logger.WriteLog` по всему проекту
- [ ] ⬜ Удалить `WriteLog` из `Mod_Utils.bas`
- [ ] ⬜ Проверить, что логи пишутся в новый файл с ротацией

### 2.2. Убрать `MsgBox` из `FillHeaderFromOrder` (1 ч)

- [ ] ⬜ Открыть `Mod_OrderHeader.bas`
- [ ] ⬜ Найти `FillHeaderFromOrder`
- [ ] ⬜ Удалить `MsgBox` из тела функции
- [ ] ⬜ Создать `_UI`-обёртку в `Mod_MainButtons.bas` или `Mod_ButtonDispatcher.bas`:
  ```vba
  Public Sub FillHeaderFromOrder_UI(ByVal OrderNum As Variant)
      Dim result As Boolean
      result = Mod_OrderHeader.FillHeaderFromOrder(OrderNum)
      If Not result Then
          MsgBox "Заказ с номером " & OrderNum & " не найден!", vbExclamation
      End If
  End Sub
  ```
- [ ] ⬜ Перенаправить вызов из `Sheet1_main.cls` на `_UI`-обёртку

### 2.3. Перенести тип `OrderHeader` из `Mod_Utils` в `Mod_OrderHeader` (0.5 ч)

- [ ] ⬜ Открыть `Mod_Utils.bas`
- [ ] ⬜ Найти `Type OrderHeader` (строка ~12)
- [ ] ⬜ Вырезать определение типа
- [ ] ⬜ Вставить в `Mod_OrderHeader.bas` (в начало модуля, до процедур)
- [ ] ⬜ Проверить, что все модули, использующие `OrderHeader`, компилируются (VBA Compile)

### 2.4. Создать `Mod_Constants.bas` — единый модуль констант (2 ч)

- [ ] ⬜ Создать файл `Mod_Constants.bas`
- [ ] ⬜ Перенести туда:
  - Все `SPISOK_COL_*` константы из `Mod_OrderHeader.bas`
  - Все `MODEL_COL_*` константы из `Mod_OrderHeader.bas`
  - Маппинг столбцов из `Mod_Import.bas` (строки ~197-199, 214-217)
  - Константы путей: `REPORT_FILE_NAME = "report.xlsx"`, `LOG_FILE_NAME = "log.txt"`
  - Тип `OrderHeader` (если ещё не перенесён в `Mod_OrderHeader`)
- [ ] ⬜ Удалить дублирующиеся константы из исходных модулей
- [ ] ⬜ Добавить `Attribute VB_Name = "Mod_Constants"` в начало файла
- [ ] ⬜ Проверить компиляцию всех модулей

### 2.5. Вынести маппинг столбцов в `Mod_Constants` (2 ч)

- [ ] ⬜ Открыть `Mod_Import.bas`
- [ ] ⬜ Найти жёстко зашитый маппинг (строки ~197-199, 214-217)
- [ ] ⬜ Создать константы в `Mod_Constants.bas`:
  ```vba
  ' Маппинг для работ: C→L, D→M, H→N
  Public Const MAP_WORKS_SRC_COL As Long = 3  ' C
  Public Const MAP_WORKS_DST_COL As Long = 12 ' L
  ' ... и т.д.
  ```
- [ ] ⬜ Заменить магические числа в `Mod_Import.bas` на константы
- [ ] ⬜ Проверить, что импорт данных работает корректно

### 2.6. Создать `Mod_SheetOps.bas` — операции с листами (4 ч)

- [ ] ⬜ Создать файл `Mod_SheetOps.bas`
- [ ] ⬜ Перенести из `Mod_Import.bas`:
  - `SearchSheetByGRZ` — поиск листа по ГРЗ
  - `RenameSheetsByGRZ` — переименование листов
  - `ImportSheet` — импорт листа из `report.xlsx`
  - `ImportDataToMain` — перенос данных на лист main
  - `ClearMainSheet_UI` — очистка листа main
- [ ] ⬜ Обновить все вызовы в `Mod_Import.bas`, `Mod_MainButtons.bas`, `Mod_ButtonDispatcher.bas`
- [ ] ⬜ Убедиться, что `Mod_Import.bas` после рефакторинга содержит только:
  - `ExtractNumberFromGRZ`
  - Функции высокого уровня, делегирующие `Mod_SheetOps`
- [ ] ⬜ Проверить полный цикл импорта: открытие → поиск → копирование → импорт данных → закрытие

### 2.7. Унифицировать обработчики кнопок (5 ч)

- [ ] ⬜ Открыть `Mod_MainButtons.bas`
- [ ] ⬜ Проанализировать все процедуры:
  - `Btn_main_Import` — уже должен быть делегирован `Mod_Import.ImportSheet` (Фаза 1.6)
  - `Btn_main_Clear` — делегировать `Mod_SheetOps.ClearMainSheet_UI`
  - `Btn_main_SelectParts` — перенести в `Mod_Selection.bas`
  - `Btn_main_SelectWorks` — перенести в `Mod_Selection.bas`
  - Остальные заглушки — перенести в `Mod_Selection.bas`
- [ ] ⬜ Открыть `Mod_SheetButtons.bas`
- [ ] ⬜ Проанализировать все процедуры-заглушки
- [ ] ⬜ Перенести логику в соответствующие модули
- [ ] ⬜ Обновить `Mod_ButtonDispatcher.bas`:
  - Все 13 обработчиков должны вызывать `_UI`-процедуры
  - Ни один обработчик не должен содержать бизнес-логику
- [ ] ⬜ Протестировать все кнопки на всех листах (main, z4, work)

### ⬜ Фаза 2 — Проверка

- [ ] ⬜ Запустить `python run_tests.py` — все тесты проходят
- [ ] ⬜ Проверить импорт из `report.xlsx` — данные корректно загружаются
- [ ] ⬜ Проверить очистку листа main
- [ ] ⬜ Проверить заполнение шапки (B2 → B3:B15)
- [ ] ⬜ Проверить все 13 кнопок на всех листах
- [ ] ⬜ Проверить, что логи пишутся с ротацией
- [ ] ⬜ Закоммитить: `git add -A && git commit -m "feat: phase 2 — architecture refactoring (constants, sheetops, logger, buttons)"`

---

## Фаза 3 — Средний приоритет (5.5 ч)

> **Цель:** улучшение качества кода, устранение косметических проблем, подготовка к будущему росту.
> **Риск:** низкий. Изменения изолированные.

### 3.1. Создать `Mod_Selection.bas` — подбор запчастей и работ (3 ч)

- [ ] ⬜ Создать файл `Mod_Selection.bas`
- [ ] ⬜ Перенести заглушки из `Mod_MainButtons.bas` и `Mod_SheetButtons.bas`
- [ ] ⬜ Реализовать базовую логику подбора:
  - `SelectParts(GRZ)` — подбор запчастей по ГРЗ
  - `SelectWorks(GRZ)` — подбор работ по ГРЗ
  - `SelectParts_UI` / `SelectWorks_UI` — UI-обёртки с MsgBox
- [ ] ⬜ Обновить `Mod_ButtonDispatcher.bas` — перенаправить вызовы на `Mod_Selection`
- [ ] ⬜ Проверить, что кнопки подбора не падают с ошибкой

### 3.2. Объединить `Sheet_z4.cls` / `Sheet_work.cls` (2 ч)

**Вариант А — вынести общую логику в модуль:**
- [ ] ⬜ Создать `Mod_SheetCommon.bas` с процедурой `ActivateSheet_UI`
- [ ] ⬜ В `Sheet_z4.cls` и `Sheet_work.cls` заменить дублирующийся код на вызов `Mod_SheetCommon.ActivateSheet_UI`

**Вариант Б — признать дубль допустимым:**
- [ ] ⬜ Синхронизировать комментарии в обоих файлах
- [ ] ⬜ Добавить поясняющий комментарий, почему дубль допустим

### 3.3. Переписать `run_tests.py` — транслит → русский (0.5 ч)

- [x] ✅ Открыть `run_tests.py`
- [x] ✅ Заменить все строки на транслите на русский язык:
  - `ZAPUSK TESTOV` → `ЗАПУСК ТЕСТОВ`
  - `Sozdanie COM-obekta` → `Создание COM-объекта`
  - И т.д.
- [x] ✅ Добавить программный сбор результатов через `GetTestResults()`
- [x] ✅ Добавить exit code (0 — все PASS, 1 — есть FAIL)
- [x] ✅ Добавить гарантированное закрытие Excel в `finally`
- [x] ✅ Добавить сохранение результатов в `test_results.log`
- [x] ✅ Проверить, что скрипт работает после изменений

### 3.4. Обновить документацию (1 ч)

- [x] ✅ Открыть `docs/DEVELOPER.md`
- [x] ✅ Добавить раздел "Тестирование" с полной таблицей TC-01..TC-30
- [x] ✅ Добавить таблицу покрытия модулей
- [x] ✅ Добавить описание, как добавить новый тест
- [x] ✅ Добавить описание запуска тестов в CI/CD
- [x] ✅ Обновить описание `Mod_FullTestRunner.bas` (564 → 1050+ строк, новые группы)
- [x] ✅ Открыть `CHANGELOG.md`
- [x] ✅ Добавить запись о новой версии с тестами

### 3.5. Добавить план тестирования на реальных данных (0.5 ч)

- [x] ✅ Добавить в `docs/DEVELOPER.md` раздел 5.5 "План тестирования на реальных данных"
- [x] ✅ Включить классификацию тестов TC-01..TC-30 (авто/ручные/требуют данных/опасные)
- [x] ✅ Включить 4 этапа тестирования: базовый, с work.xlsm, с report.xlsx, ручной
- [x] ✅ Включить чек-лист подготовки и проверки исправленных проблем
- [x] ✅ Включить критерии прохождения (PASS/PASS_WITH_SKIPS/FAIL/FAIL_CRITICAL)

### ⬜ Фаза 3 — Проверка

- [ ] ⬜ Запустить `python scripts/run_tests.py` — все тесты проходят
- [ ] ⬜ Проверить, что все кнопки работают
- [ ] ⬜ Проверить, что документация актуальна
- [ ] ⬜ Закоммитить: `git add -A && git commit -m "feat: complete test suite TC-01..TC-30, improve test infrastructure, update docs"`

---

## Фаза 4 — Будущее (опционально)

> **Цель:** масштабирование для больших справочников (10000+ записей).
> **Риск:** высокий. Требует изменений в архитектуре импорта.

- [ ] ⬜ **Вынос справочников в отдельные книги** — когда справочники превысят 10000 записей
- [ ] ⬜ **Переход на SQLite через ADO** — для сложных запросов подбора
- [ ] ⬜ **Автоматическое тестирование в CI/CD** — добавить `run_tests.py` в GitHub Actions
- [ ] ⬜ **Pre-commit hook** — проверка кодировки .bas/.cls файлов (UTF-8 без BOM)

---

## Сводка прогресса

| Фаза | Задач | Часы | Статус |
|------|-------|------|--------|
| Фаза 0 — Подготовка | 4 | 0.5 | ✅ |
| Фаза 1 — Критические исправления | 6 | 3.5 | ✅ |
| Фаза 2 — Архитектурные изменения | 7 | 18.0 | ⬜ |
| Фаза 3 — Средний приоритет | 5 | 6.0 | ✅ |
| Фаза 4 — Будущее | 4 | — | ⬜ |
| **Итого** | **26** | **28.0** | **✅ ~42%** |

---

## Быстрые ссылки

| Ресурс | Ссылка |
|--------|--------|
| План модернизации | [`plans/architecture_modernization_proposal.md`](plans/architecture_modernization_proposal.md) |
| Документация разработчика | [`docs/DEVELOPER.md`](docs/DEVELOPER.md) |
| Git-инструкции | [`docs/git-workflow.md`](docs/git-workflow.md) |
| Правила SourceCraft | [`.ycarules`](.ycarules) |
| Тестовый раннер | [`Mod_FullTestRunner.bas`](Mod_FullTestRunner.bas) |
| Скрипт экспорта VBA | [`export_vba.py`](export_vba.py) |
| Скрипт импорта VBA | [`impVBA.py`](impVBA.py) |
| Запуск тестов | [`run_tests.py`](run_tests.py) |