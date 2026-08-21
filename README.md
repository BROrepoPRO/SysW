# SysW — Система автоматизации обработки заказ-нарядов авторемонта

**Версия:** v1.0.4

**Назначение:** импорт, анализ и учёт данных заказ-нарядов из Excel.

Система автоматизирует заполнение шапки заказ-наряда, импорт данных из отчётов, поиск по ГРЗ и ведение учёта работ и запчастей. Реализована в виде модульной VBA-надстройки для Excel с поддержкой Git-контроля версий и CI/CD.

### Архитектура

Проект построен на модульной архитектуре VBA с разделением на 13 стандартных модулей (`.bas`), 2 класса (`.cls`) и 1 класс листа (`.cls`). Исходный код организован в директории `src/`:
- `src/modules/` — бизнес-логика, утилиты, логирование, обработчики кнопок
- `src/classes/` — классы-объекты (деталь, работа)
- `src/sheets/` — классы-обработчики событий листов Excel

Скрипты автоматизации (Python, PowerShell) вынесены в `scripts/`. Документация — в `docs/`, планы архитектурных изменений — в `plans/`.

---

## Технологический стек

| Технология | Назначение |
|------------|-----------|
| **VBA (Excel)** | Модульная система: импорт, парсинг, заполнение шапки, тестирование |
| **Python** | Скрипты экспорта/импорта VBA-модулей, запуск тестов |
| **PowerShell** | Альтернативный скрипт импорта VBA из Excel |
| **SQLite** | **Планируется** (задачи R-S1..R-S8 в ROADMAP). Текущее хранение — в `work.xlsm` и внешних Excel-файлах (`base/models/*.xlsx`) |
| **Git / GitHub** | Контроль версий, CI/CD (GitHub Actions) |
| **SourceCraft Code Assistant** | Многоролевая архитектура разработки |

---

## Структура проекта

```
SysW (https://github.com/BROrepoPRO/SysW.git)
├── src/                       # Исходный код VBA
│   ├── modules/               # 13 .bas модулей
│   ├── classes/               # 2 .cls класса (PartIdentity, WorkIdentity)
│   └── sheets/                # 1 .cls лист
├── base/                      # Шаблоны и образцы данных
│   ├── templates/             # Шаблоны (work, work0, model, model0, report0)
│   └── models/                # Модели данных (Excel-файлы)
├── scripts/                   # Python + PowerShell скрипты
│   ├── config.py              # Общая конфигурация путей (Python)
│   ├── config.ps1             # Общая конфигурация путей (PowerShell)
│   ├── export_vba.py          # Выгрузка VBA из Excel
│   ├── impVBA.py              # Загрузка VBA в Excel
│   ├── run_tests.py           # Запуск тестов
│   ├── build_templates.py     # Создание шаблонов base/templates/
│   ├── apply_protection_templates.py  # Защита листов шаблонов
│   └── Set-ExcelTrust.ps1     # Настройка доверия Excel к VBA
├── logs/                      # Логи и результаты тестов
│   ├── log.txt                # Основной лог VBA-модулей
│   └── test_results.log       # Результаты прогона тестов
├── docs/                      # Документация проекта
│   ├── sourcecraft-guide.md   # Руководство по SourceCraft
│   ├── git-workflow.md        # Git-инструкции
│   ├── ARCHITECTURE.md        # Архитектура проекта
│   ├── DEVELOPER.md           # Техническая документация разработчика
│   ├── ROADMAP.md             # Единый план развития
│   └── table.md               # Справочник таблиц
├── plans/                     # Планы изменений и архитектурные решения
│   └── _archive/              # Архив выполненных планов
├── .github/workflows/         # CI/CD (GitHub Actions)
│   ├── vba-check.yml          # Проверка VBA-файлов
│   └── docs-check.yml         # Проверка консистентности документации
├── .vscode/                   # Настройки VS Code
├── .ycarules                  # Правила SourceCraft
├── .gitattributes             # Нормализация Git
└── work.xlsm                  # Excel-файл с макросами (в .gitignore)
```

## CI/CD

Проект использует GitHub Actions для автоматических проверок при push и pull request на ветки `main`/`dev`. Workflow расположены в `.github/workflows/`:

- [**vba-check.yml**](.github/workflows/vba-check.yml) — проверка наличия всех VBA-файлов (`src/modules/*.bas`, `src/classes/*.cls`, `src/sheets/*.cls`), их корректной UTF-8 кодировки и базового синтаксиса, а также актуальности `CHANGELOG.md`.
- [**docs-check.yml**](.github/workflows/docs-check.yml) — проверка консистентности документации (`python scripts/check_docs.py --check`).

> Документация по CHANGELOG хранится в [`docs/CHANGELOG.md`](docs/CHANGELOG.md).

---

## Быстрый старт

1. **Клонировать репозиторий:**
   ```bash
   git clone https://github.com/BROrepoPRO/SysW.git
   cd SysW
   ```

2. **Открыть проект в VS Code:**
   ```bash
   code .
   ```

3. **Активировать виртуальное окружение Python:**
   ```powershell
   .venv\Scripts\Activate.ps1
   ```

4. **Открыть `work.xlsm` в Excel** (макросы должны быть включены).

5. **Синхронизировать VBA-модули из Excel на диск** (после работы в Excel):
   ```bash
   python scripts/export_vba.py
   ```

6. **Загрузить VBA-модули с диска в Excel** (после редактирования в VS Code):
   ```bash
   python scripts/impVBA.py
   ```

7. **Запустить тесты:**
   ```bash
   python scripts/run_tests.py
   ```

---

## Состав команды

| Роль | Эмодзи | Обязанности |
|------|--------|-------------|
| **Начальник мира** | 👤 | Генерация идей, постановка задач, ключевые решения, архитектура и приоритеты |
| **Оркестратор** | 🪃 | Координация, декомпозиция задач, промты для Code Assistant, контроль качества |
| **Architect** | 🏗️ | Проектирование архитектуры, создание планов, стратегические решения |
| **Code** | 💻 | Написание и модификация кода по готовым планам |
| **Debug** | 🪲 | Диагностика ошибок, поиск первопричин, добавление логирования |
| **Ask** | ❓ | Анализ кода, ответы на вопросы, объяснение концепций |

---

## Документация

- [**Руководство по SourceCraft**](docs/sourcecraft-guide.md) — архитектура взаимодействия агентов, правила, рабочий процесс, описание скриптов
- [**Git-инструкции**](docs/git-workflow.md) — веточная стратегия, формат коммитов (Conventional Commits), pre-commit процедуры, работа с GitHub
- [**Техническая документация разработчика**](docs/DEVELOPER.md) — архитектура VBA-модулей, двухфазная кодировка, процессы импорта/экспорта, настройка окружения, CI/CD

## Планы

- [**ROADMAP**](docs/ROADMAP.md) — единый план развития проекта, задачи по этапам

---

## История изменений

[CHANGELOG.md](docs/CHANGELOG.md)