# Руководство по работе с SourceCraft Code Assistant

## Архитектура взаимодействия

### Роли агентов

В проекте SysW используется многоролевая архитектура SourceCraft Code Assistant:

| Роль | Назначение |
|------|-----------|
| **🏗️ Architect** | Проектирование архитектуры, создание планов, стратегические решения |
| **💻 Code** | Написание и модификация кода по готовым планам |
| **❓ Ask** | Анализ кода, ответы на вопросы, объяснение концепций |
| **🪲 Debug** | Диагностика ошибок, поиск первопричин, добавление логирования |
| **🪃 Orchestrator** | Координация работы агентов, декомпозиция задач |

### Процесс разработки

1. **Начальник мира** ставит задачу
2. **Главный помощник** (человек или Orchestrator) декомпозирует задачу в план
3. План сохраняется в директории [`plans/`](../plans/) и согласовывается
4. После утверждения плана **Code Assistant** реализует изменения
5. Результат проверяется и принимается

---

## Ключевые правила из `.ycarules`

### Двухфазная модель кодировки VBA

VBA-файлы (`.bas`, `.cls`) используют двухфазную модель кодировки:

- **На диске** — UTF-8 (для VS Code, Git, code review)
- **В Excel/VBA Editor** — Windows-1251 (требование COM-автоматизации)

Конвертация выполняется скриптами:
- [`impVBA.py`](../scripts/impVBA.py) — UTF-8 → CP1251 (перед импортом в Excel)
- [`export_vba.py`](../scripts/export_vba.py) — CP1251 → UTF-8 (после экспорта из Excel)

### Запрет самовольных правок VBA-модулей

Изменения в `.bas`/`.cls`/`.frm` — только после создания и утверждения плана в [`plans/`](../plans/). Категорически запрещено изменять утверждённую логику (маппинг полей, алгоритмы функций, структуры таблиц) без прямого письменного указания Начальника мира.

### Запрет кириллицы в именах

Имена файлов, папок, модулей, процедур, переменных — только латиница.

### Язык общения

Все ответы агентов — **только на русском языке**.

---

## Рабочий процесс

### Начало сессии

1. Открыть VS Code в `L:\PROject\SysW`
2. Терминал SourceCraft (PowerShell с UTF-8 и `.venv`) настраивается автоматически через `.vscode/settings.json`
3. Убедиться, что активировано виртуальное окружение: `.venv\Scripts\Activate.ps1`

### Создание плана

1. Перед любыми изменениями VBA-модулей создать план в [`plans/`](../plans/)
2. Формат плана: Markdown (`.md`) с чеклистами задач
3. Согласовать план с пользователем
4. Только после утверждения приступать к реализации

### Внесение изменений

1. Работа с VBA-модулями — только через скрипты импорта/экспорта
2. После изменений — запустить тесты: `python scripts/run_tests.py`
3. Обновить [`CHANGELOG.md`](../CHANGELOG.md)

### Обновление CHANGELOG.md

При любых изменениях, влияющих на функциональность проекта, обновлять [`CHANGELOG.md`](../CHANGELOG.md) в формате [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/).

---

## Структура проекта

```
L:\PROject\SysW\
├── .github/              # CI/CD (GitHub Actions)
│   └── workflows/
│       └── vba-check.yml
├── .vscode/              # Настройки VS Code (терминал, кодировка, VBA Language Server)
├── base/                 # Шаблоны и образцы данных
│   └── models/               # Модели данных
├── docs/                 # Документация проекта
│   ├── sourcecraft-guide.md      # Настоящее руководство
│   ├── git-workflow.md           # Git-инструкции (ветки, коммиты, CI)
│   ├── DEVELOPER.md              # Техническая документация разработчика
│   └── ARCHITECTURE_SQLITE.md    # Архитектура выноса данных в SQLite
├── plans/                # Планы изменений, архитектурные решения, отчёты
│   └── _archive/             # Архив выполненных планов
├── scripts/              # Скрипты автоматизации
│   ├── config.py             # Конфигурация проекта (пути, настройки)
│   ├── config.ps1            # Конфигурация окружения PowerShell
│   ├── export_vba.py         # Выгрузка VBA-модулей из Excel (CP1251 → UTF-8)
│   ├── impVBA.py             # Загрузка VBA-модулей в Excel (UTF-8 → CP1251)
│   ├── run_tests.py          # Запуск тестов VBA
│   └── Set-ExcelTrust.ps1    # Настройка доверия Excel
├── src/                  # Исходный код VBA
│   ├── modules/              # 13 .bas модулей
│   │   ├── Mod_AutoMatch.bas
│   │   ├── Mod_ButtonDispatcher.bas
│   │   ├── Mod_Constants.bas
│   │   ├── Mod_FullTestRunner.bas
│   │   ├── Mod_Import.bas
│   │   ├── Mod_Logger.bas
│   │   ├── Mod_MainButtons.bas
│   │   ├── Mod_ModelDB.bas
│   │   ├── Mod_OrderHeader.bas
│   │   ├── Mod_PickWork.bas
│   │   ├── Mod_SheetButtons.bas
│   │   ├── Mod_SheetOps.bas
│   │   └── Mod_Utils.bas
│   └── sheets/               # 3 .cls листа
│       ├── Лист2_main.cls
│       ├── Sheet_work.cls
│       └── Sheet_z4.cls
├── .gitattributes        # Настройки Git для нормализации кодировок
├── .ycarules             # Правила для SourceCraft Code Assistant
├── CHANGELOG.md          # История изменений проекта
├── README.md             # Основное описание проекта
└── work.xlsm             # Excel-файл с макросами (в .gitignore)
```

### Назначение директорий

| Директория | Назначение |
|-----------|-----------|
| `base/` | Шаблоны и образцы данных |
| `base/models/` | Модели данных |
| `plans/` | Планы изменений, архитектурные решения, отчёты. Создаются перед изменением VBA-модулей |
| `plans/_archive/` | Архив выполненных планов |
| `scripts/` | Скрипты автоматизации (Python + PowerShell). Python-скрипты в UTF-8, PowerShell в UTF-8 with BOM |
| `docs/` | Документация проекта: `sourcecraft-guide.md` (руководство по SourceCraft), `git-workflow.md` (Git-инструкции), `DEVELOPER.md` (техническая документация), `ARCHITECTURE_SQLITE.md` (архитектура SQLite) |
| `.vscode/` | Настройки VS Code (кодировка UTF-8, кастомный терминал, VBA Language Server) |

---

## Скрипты

### Python-скрипты (основные)

| Скрипт | Назначение | Кодировка |
|--------|-----------|-----------|
| [`config.py`](../scripts/config.py) | Конфигурация проекта (пути, настройки) | UTF-8 |
| [`export_vba.py`](../scripts/export_vba.py) | Выгрузка VBA-модулей из Excel на диск (CP1251 → UTF-8) | UTF-8 |
| [`impVBA.py`](../scripts/impVBA.py) | Загрузка VBA-модулей с диска в Excel (UTF-8 → CP1251) | UTF-8 |
| [`run_tests.py`](../scripts/run_tests.py) | Запуск тестов VBA | UTF-8 |

### PowerShell-скрипты

| Скрипт | Назначение |
|--------|-----------|
| [`config.ps1`](../scripts/config.ps1) | Конфигурация окружения PowerShell |
| [`Set-ExcelTrust.ps1`](../scripts/Set-ExcelTrust.ps1) | Настройка доверия Excel для работы VBA-макросов |

**Важно:** PowerShell-скрипты должны быть в кодировке UTF-8 with BOM (требование PowerShell для корректной обработки кириллицы).

---

## История версий

| Версия | Дата | Изменения |
|--------|------|-----------|
| 0.15.0 | 2026-07-25 | Тесты TC-31..TC-44 (Mod_ModelDB, Mod_PickWork, Mod_AutoMatch); покрытие 46% → 69%; deprecated MODELDB_BASE_PATH; делегирование Btn_main_ImportVH_Click; COMPONENTS расширен до 13+3; $ProjectPath в Set-ExcelTrust.ps1 |
| 0.14.0 | 2026-07-25 | Централизация логов: LOGS_DIR, директория logs/, перенос логов из корня; удаление мусора из base/models/ |
| 0.13.0 | 2026-07-25 | Добавлены `Mod_ModelDB`, `Mod_PickWork`, `Mod_AutoMatch`; новые скрипты `config.py`, `config.ps1`; удалён `Import-VbaFromExcel.ps1` |
| 0.10.0 | 2026-07-21 | `Mod_Import.ImportFromB2_UI`, исправление ошибки импорта (пропуск заголовка) |
| 0.9.0 | 2026-07-21 | Рефакторинг `Mod_LibName.bas` → объединение с `Mod_Constants.bas` |
| 0.8.0 | 2026-07-19 | Модуль `Mod_FullTestRunner.bas`, 16 тестовых сценариев, новая архитектура импорта |
| 0.7.1 | 2026-07-16 | Защитное программирование и обработка ошибок в модулях VBA (аудит стабильности) |
| 0.7.0 | 2026-07-15 | Новые тесты TC-21..TC-30, функция GetTestResults, silent mode, раздел тестирования в DEVELOPER.md |
| 0.6.0 | 2026-07-15 | Новая структура каталогов: src/modules/, src/sheets/, scripts/ |
| 0.5.0 | 2026-07-14 | Техническая документация DEVELOPER.md, расширение README.md |
| 0.4.0 | 2026-07-14 | GitHub Actions, git-workflow.md, .gitignore |
| 0.3.0 | 2026-07-14 | Интеграция SourceCraft, .ycarules, scripts/, docs/ |
| 0.2.0 | 2026-07-13 | Полный тестовый модуль, автоматические тесты |
| 0.1.0 | 2026-07-09 | Начальная версия: автозаполнение шапки, базовые скрипты, .ycarules |