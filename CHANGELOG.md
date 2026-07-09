# История изменений

Все заметные изменения в проекте будут документироваться в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/), 
версионирование следует [Semantic Versioning](https://semver.org/lang/ru/).

## [0.1.0] — 2026-07-09

### Добавлено
- Автоматическое заполнение шапки заказ-наряда при вводе номера заказа в ячейку B2 листа main.
- Модуль `Mod_OrderHeader` с независимой функцией `FillHeaderFromOrder`.
- Обработчик `Worksheet_Change` для листа main.
- Инструмент `Write-UTF8BOM.ps1` для сохранения PowerShell-скриптов с BOM.
- Файл правил `.ycarules` для Code Assistant.
- Git-репозиторий с ветками `main` и `dev`.
- Начальная документация (README.md, .gitignore).