# История изменений

Все заметные изменения в проекте будут документироваться в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/),
версионирование следует [Semantic Versioning](https://semver.org/lang/ru/).

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