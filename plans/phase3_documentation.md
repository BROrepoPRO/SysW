# План Этапа 3: Разработка и интеграция документации проекта

**Версия:** 0.5.0
**Дата:** 2026-07-14
**Статус:** Черновик

---

## Обзор

Обновление документации проекта SysW: расширение `README.md`, создание `docs/DEVELOPER.md`, обновление `CHANGELOG.md` и синхронизация ссылок в `docs/sourcecraft-guide.md`.

---

## Задача 1: Обновление `README.md`

**Файл:** [`README.md`](../README.md)

Текущий README.md содержит базовое описание. Требуется расширить до полноценного landing page проекта.

### Структура нового README.md

```markdown
# SysW — Система автоматизации обработки заказ-нарядов авторемонта

## Описание проекта
(краткое описание системы: импорт, анализ и учёт данных заказ-нарядов из Excel в SQLite)

## Технологический стек
- **VBA (Excel)** — модульная система: импорт, парсинг, заполнение шапки, тестирование
- **Python** — скрипты экспорта/импорта VBA-модулей, запуск тестов
- **PowerShell** — альтернативный скрипт импорта VBA из Excel
- **SQLite** — хранение данных заказ-нарядов, работ, запчастей
- **Git / GitHub** — контроль версий, CI/CD (GitHub Actions)
- **SourceCraft Code Assistant** — многоролевая архитектура разработки

## Структура проекта
(краткая схема директорий со ссылками на docs/)

```
L:\PROject\SysW\
├── docs/                 # Документация проекта
│   ├── sourcecraft-guide.md  # Руководство по SourceCraft
│   ├── git-workflow.md       # Git-инструкции
│   └── DEVELOPER.md          # Техническая документация
├── plans/                # Планы изменений и архитектурные решения
├── scripts/              # PowerShell-скрипты автоматизации
├── .github/workflows/    # CI/CD (GitHub Actions)
├── Mod_*.bas             # VBA-модули
├── Sheet1_main.cls       # VBA-класс листа main
├── *.py                  # Python-скрипты
└── work.xlsm             # Excel-файл с макросами (в .gitignore)
```

## Быстрый старт

1. Клонировать репозиторий
2. Открыть `L:\PROject\SysW` в VS Code
3. Активировать виртуальное окружение Python: `.venv\Scripts\Activate.ps1`
4. Открыть `work.xlsm` в Excel (макросы должны быть включены)
5. Для импорта VBA-модулей из Excel на диск: `python export_vba.py`
6. Для загрузки VBA-модулей с диска в Excel: `python import_all_vba.py`
7. Для запуска тестов: `python run_tests.py`

## Состав команды

| Роль | Эмодзи | Обязанности |
|------|--------|-------------|
| **Начальник мира** | 👤 | Генерация идей, постановка задач, ключевые решения, архитектура и приоритеты |
| **Оркестратор** | 🪃 | Координация, декомпозиция задач, промты для Code Assistant, контроль качества |
| **Architect** | 🏗️ | Проектирование архитектуры, создание планов, стратегические решения |
| **Code** | 💻 | Написание и модификация кода по готовым планам |
| **Debug** | 🪲 | Диагностика ошибок, поиск первопричин, добавление логирования |
| **Ask** | ❓ | Анализ кода, ответы на вопросы, объяснение концепций |

## Документация

- [Руководство по SourceCraft](docs/sourcecraft-guide.md) — архитектура взаимодействия, правила, рабочий процесс
- [Git-инструкции](docs/git-workflow.md) — веточная стратегия, формат коммитов, pre-commit процедуры
- [Техническая документация](docs/DEVELOPER.md) — архитектура VBA-модулей, кодировка, окружение, CI/CD

## История изменений

[CHANGELOG.md](CHANGELOG.md)
```

### Конкретные изменения в README.md

| Секция | Действие | Примечание |
|--------|----------|------------|
| Заголовок | Оставить без изменений | `# SysW — Система автоматизации обработки заказ-нарядов авторемонта` |
| Описание проекта | Расширить | Добавить краткое описание системы |
| Технический состав | Заменить на "Технологический стек" | Детализировать: VBA, Python, PowerShell, SQLite, Git, SourceCraft |
| Структура проекта | Добавить новую секцию | Схема директорий со ссылками на docs/ |
| Быстрый старт | Добавить новую секцию | 7 шагов: клонирование → открытие → .venv → Excel → export → import → тесты |
| Команда проекта | Расширить | Добавить роли SourceCraft: Оркестратор, Architect, Code, Debug, Ask |
| Документация | Добавить новую секцию | Ссылки на sourcecraft-guide.md, git-workflow.md, DEVELOPER.md |
| История изменений | Оставить ссылку на CHANGELOG.md | Без изменений |

---

## Задача 2: Создание `docs/DEVELOPER.md`

**Файл:** [`docs/DEVELOPER.md`](../docs/DEVELOPER.md) — **новый файл**

### Структура DEVELOPER.md

```markdown
# Техническая документация разработчика — SysW

## 1. Архитектура проекта

### 1.1 Общая архитектура

(Описание модульной структуры VBA: как модули взаимодействуют, диаграмма зависимостей)

### 1.2 Схема взаимодействия модулей

```
Sheet1_main.cls (Worksheet_Change)
       │
       ▼
Mod_OrderHeader.FillHeaderFromOrder()
       │
       ├── Mod_Utils.GetSheetByName()
       ├── Mod_Utils.FileExists()
       └── Mod_Utils.FormatDateSQL()

Mod_ButtonDispatcher (обработчики кнопок)
       │
       ├── Mod_Import.ClearMainSheet_UI()
       ├── Mod_Import.ImportSheet_UI()
       ├── Mod_Import.ClearHeader_UI()
       ├── Mod_OrderHeader.FillHeaderFromOrder_UI()
       └── ...

Mod_FullTestRunner.RunAllTests()
       │
       ├── RunUtilsTests()         → Mod_Utils
       ├── RunOrderHeaderTests()   → Mod_OrderHeader
       ├── RunImportTests()        → Mod_Import
       └── RunButtonTests()        → Mod_ButtonDispatcher
```

## 2. Описание модулей VBA

### 2.1 Mod_Utils.bas — Утилитарные функции

**Назначение:** Вспомогательные функции для работы с Excel.

**Ключевые функции:**
- `GetSheetByName(wb, SheetName)` — получение листа по имени без ошибки
- `GetWorkbookPath()` — путь к текущей книге
- `FileExists(FilePath)` — проверка существования файла
- `GetCurrentUser()` — имя пользователя Windows
- `FormatDateSQL(d)` — форматирование даты в ГГГГ-ММ-ДД

**Типы данных:**
- `OrderHeader` — пользовательский тип для хранения данных заказа-наряда (поля: OrderNumber, ModelName, GRZ, VIN, GarageNumber, YearMade, MileageValue, DateValue)

### 2.2 Mod_OrderHeader.bas — Работа с шапкой заказ-наряда

**Назначение:** Заполнение заголовка заказа-наряда (B3:B15) на листе main.

**Ключевые функции:**
- `FillHeaderFromOrder(OrderNum)` — основная функция: заполняет B3:B15 данными из листов spisok и model по номеру заказа
- `FillHeaderFromOrder_UI()` — UI-обёртка с MsgBox для пользователя

**Константы столбцов:**
- `SPISOK_COL_MODEL` (2), `SPISOK_COL_GRZ` (3), `SPISOK_COL_VIN` (4), `SPISOK_COL_GARAGE` (5), `SPISOK_COL_YEAR` (6), `SPISOK_COL_MILEAGE` (7), `SPISOK_COL_DATE` (8), `SPISOK_COL_NOTE` (10)
- `MODEL_COL_GROUP` (2), `MODEL_COL_PRICE` (3), `MODEL_COL_WORKS_ORIG` (4), `MODEL_COL_WORKS_MOD` (5)

### 2.3 Mod_Import.bas — Импорт данных

**Назначение:** Импорт данных из Excel в SQLite и обратно.

**Ключевые функции:**
- `ExtractNumberFromGRZ(GRZ)` — извлечение цифровой группы из ГРЗ
- `SearchSheetByGRZ(GRZ)` — поиск листа в report.xlsx по номеру ГРЗ
- `ClearMainSheet_UI()` — очистка всех данных на листе main
- `ImportSheet_UI()` — запуск импорта из отчёта по ГРЗ
- `ClearHeader_UI()` — очистка шапки заказа (B3:B15)

### 2.4 Mod_ButtonDispatcher.bas — Диспетчер кнопок

**Назначение:** Содержит ТОЛЬКО однострочные вызовы UI-процедур. Является прослойкой между кнопками на формах и бизнес-логикой.

**Обработчики:**
- `Btn_main_Clear_Click()` → `Mod_Import.ClearMainSheet_UI`
- `Btn_main_Import_Click()` → `Mod_Import.ImportSheet_UI`
- `Btn_main_FillHeader_Click()` → `Mod_OrderHeader.FillHeaderFromOrder_UI`
- `Btn_main_ClearHeader_Click()` → `Mod_Import.ClearHeader_UI`
- `Btn_main_ImportByInput_Click()` → импорт по введённому ГРЗ

### 2.5 Mod_FullTestRunner.bas — Тестовый раннер

**Назначение:** Автоматическое тестирование VBA-модулей (20 сценариев TC-01..TC-20).

**Группы тестов:**
- `RunUtilsTests()` — тесты утилит (TC-01..TC-04, TC-06, TC-07, TC-19, TC-20)
- `RunOrderHeaderTests()` — тесты OrderHeader
- `RunImportTests()` — тесты импорта
- `RunButtonTests()` — тесты кнопок

**Механизмы:**
- SKIP для тестов, зависящих от отсутствующих данных
- Сохранение/восстановление состояния листов
- Подсчёт статистики: Total, Passed, Failed, Skipped

### 2.6 Sheet1_main.cls — Основной лист

**Назначение:** Класс листа main. Обрабатывает событие `Worksheet_Change`.

**Логика:**
- При изменении ячейки B2 (номер заказа) автоматически вызывает `Mod_OrderHeader.FillHeaderFromOrder()`
- Защита от рекурсии через `Static isProcessing As Boolean`
- Очистка B3:B15 перед заполнением

## 3. Правила кодировки (двухфазная система)

### 3.1 Принцип

| Фаза | Где | Кодировка | Инструмент |
|------|-----|-----------|------------|
| Диск | VS Code, Git, code review | UTF-8 (без BOM) | `export_vba.py` |
| Excel | VBA Editor, COM-автоматизация | Windows-1251 (CP1251) | `import_all_vba.py` |

### 3.2 Конвертация

- **Экспорт из Excel (CP1251 → UTF-8):** `python export_vba.py`
  - Читает модули из Excel через COM (CP1251)
  - Сохраняет на диск в UTF-8
  - Использует `_temp_export/` как временную директорию

- **Импорт в Excel (UTF-8 → CP1251):** `python import_all_vba.py`
  - Читает файлы с диска (UTF-8, UTF-8 with BOM, или CP1251 fallback)
  - Удаляет `Attribute`-строки (недопустимы в VBA при импорте)
  - Сохраняет во временную директорию `_temp_import/` в CP1251
  - Импортирует в Excel через COM

### 3.3 Git-нормализация

- `.gitattributes` — настройки нормализации UTF-8 для Git
- `.vscode/settings.json` — `files.encoding: "utf8"` для всех VBA-файлов

## 4. Процессы импорта/экспорта

### 4.1 export_vba.py

**Назначение:** Выгрузка VBA-модулей из Excel на диск.

**Процесс:**
1. Запускает Excel через COM (`win32com.client`)
2. Создаёт временную директорию `_temp_export/`
3. Экспортирует каждый компонент VBA-проекта во временную директорию (CP1251)
4. Конвертирует каждый файл из CP1251 в UTF-8
5. Копирует в корень проекта
6. Удаляет временную директорию

**Маппинг компонентов:**
```
Mod_Utils → Mod_Utils.bas
Mod_OrderHeader → Mod_OrderHeader.bas
Mod_Import → Mod_Import.bas
Mod_ButtonDispatcher → Mod_ButtonDispatcher.bas
Mod_FullTestRunner → Mod_FullTestRunner.bas
Sheet1 → Sheet1_main.cls
```

**Использование:**
```bash
python export_vba.py          # Экспорт всех модулей
python export_vba.py --dry    # Просмотр без экспорта
```

### 4.2 import_all_vba.py

**Назначение:** Загрузка VBA-модулей с диска в Excel.

**Процесс:**
1. Читает каждый файл с автоопределением кодировки (UTF-8 with BOM → UTF-8 → CP1251)
2. Удаляет `Attribute`-строки (VBA не принимает их при программном импорте)
3. Сохраняет во временную директорию `_temp_import/` в CP1251
4. Запускает Excel через COM
5. Удаляет существующие компоненты с теми же именами
6. Импортирует новые компоненты из временной директории
7. Сохраняет и закрывает Excel
8. Удаляет временную директорию

**Использование:**
```bash
python import_all_vba.py      # Импорт всех модулей
```

### 4.3 scripts/Import-VbaFromExcel.ps1

**Назначение:** Альтернативный PowerShell-скрипт для импорта VBA-модулей из Excel.

**Процесс:**
1. Запускает Excel через COM
2. Экспортирует каждый компонент VBA-проекта
3. Конвертирует из CP1251 в UTF-8
4. Сохраняет в указанную директорию

**Использование:**
```powershell
.\scripts\Import-VbaFromExcel.ps1              # По умолчанию
.\scripts\Import-VbaFromExcel.ps1 -DryRun      # Режим просмотра
.\scripts\Import-VbaFromExcel.ps1 -ExcelPath "C:\path\to\work.xlsm" -OutputDir "C:\output"
```

## 5. Настройка окружения

### 5.1 VS Code

- Открыть `L:\PROject\SysW` как корневую директорию
- Установить расширение "VBA Language Server" (рекомендуется)
- Настройки в `.vscode/settings.json`:
  - `files.encoding: "utf8"` — кодировка UTF-8 для всех файлов
  - `terminal.integrated.defaultProfile.windows: "PowerShell"` — терминал PowerShell
  - Кастомный профиль терминала SourceCraft с UTF-8 и .venv

### 5.2 Python

- Установить Python 3.12+
- Создать виртуальное окружение: `python -m venv .venv`
- Активировать: `.venv\Scripts\Activate.ps1`
- Установить зависимости: `pip install pywin32` (требуется для COM-автоматизации Excel)

### 5.3 PowerShell

- PowerShell 5.1+ (поставляется с Windows 10)
- Скрипты должны быть в UTF-8 with BOM (требование PowerShell для кириллицы)
- Политика выполнения: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

## 6. CI/CD

### 6.1 GitHub Actions Workflow

**Файл:** `.github/workflows/vba-check.yml`

**Триггеры:** push в main/dev, pull request в main/dev

**Проверки:**
1. **Check VBA files exist** — проверка наличия всех 6 VBA-файлов
2. **Check UTF-8 encoding** — валидация UTF-8 кодировки каждого файла
3. **Check VBA syntax (basic)** — проверка на недопустимые символы
4. **Check CHANGELOG updated** — проверка наличия записи за сегодня
5. **Check .gitignore consistency** — проверка наличия .gitignore

### 6.2 Pre-commit процедура

Перед каждым коммитом:
1. `git status` — проверить статус
2. `python export_vba.py` — синхронизировать VBA-модули (если работали в Excel)
3. Обновить `CHANGELOG.md`
4. Проверить кодировку UTF-8
5. `git add` и `git commit`

---

## Приложение A: Схема зависимостей модулей

```
Sheet1_main.cls
  └── Mod_OrderHeader.FillHeaderFromOrder()
        └── Mod_Utils (GetSheetByName)

Mod_ButtonDispatcher
  ├── Mod_Import (ClearMainSheet_UI, ImportSheet_UI, ClearHeader_UI)
  └── Mod_OrderHeader (FillHeaderFromOrder_UI)

Mod_FullTestRunner
  ├── Mod_Utils (тесты утилит)
  ├── Mod_OrderHeader (тесты OrderHeader)
  ├── Mod_Import (тесты импорта)
  └── Mod_ButtonDispatcher (тесты кнопок)
```

## Приложение B: Маппинг файлов VBA

| VBA-компонент | Файл на диске | Тип |
|--------------|---------------|-----|
| Mod_Utils | Mod_Utils.bas | Стандартный модуль |
| Mod_OrderHeader | Mod_OrderHeader.bas | Стандартный модуль |
| Mod_Import | Mod_Import.bas | Стандартный модуль |
| Mod_ButtonDispatcher | Mod_ButtonDispatcher.bas | Стандартный модуль |
| Mod_FullTestRunner | Mod_FullTestRunner.bas | Стандартный модуль |
| Sheet1 | Sheet1_main.cls | Класс листа |
```

---

## Задача 3: Обновление `CHANGELOG.md`

**Файл:** [`CHANGELOG.md`](../CHANGELOG.md)

Добавить запись о версии 0.5.0 перед версией 0.4.0.

### Содержание новой записи

```markdown
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
```

---

## Задача 4: Синхронизация `docs/sourcecraft-guide.md`

**Файл:** [`docs/sourcecraft-guide.md`](../docs/sourcecraft-guide.md)

Незначительное обновление: добавить ссылку на `docs/DEVELOPER.md` в раздел "Структура проекта" (строка 88).

### Изменение

В секции `Структура проекта` (строка 88) добавить строку:
```markdown
│   ├── DEVELOPER.md            # Техническая документация разработчика
```

---

## Порядок выполнения

1. Обновить `docs/sourcecraft-guide.md` — добавить ссылку на DEVELOPER.md
2. Создать `docs/DEVELOPER.md` — полная техническая документация
3. Обновить `README.md` — расширенное описание проекта
4. Обновить `CHANGELOG.md` — добавить запись 0.5.0
5. Проверить консистентность ссылок между всеми документами