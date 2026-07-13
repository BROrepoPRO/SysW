# История изменений

Все заметные изменения в проекте будут документироваться в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/),
версионирование следует [Semantic Versioning](https://semver.org/lang/ru/).

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