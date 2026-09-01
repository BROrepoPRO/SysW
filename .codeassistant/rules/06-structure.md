# Структура проекта — рабочие правила [S]

> Зеркало модуля «Структура проекта» из `.ycarules`. Коды правил идентичны.

## [S1] ОРГАНИЗАЦИЯ ПАПОК

| Путь | Содержимое |
|------|------------|
| `base/models/` | Модели групп: `4x4.xlsm`, `2170.xlsm`, `2180.xlsm`, `2190.xlsm`, `GAZ.xlsm`, `UAZ.xlsm` |
| `base/templates/` | Шаблоны: `model.xlsm`, `model0.xlsm`, `report0.xlsx`, `work.xlsm`, `work0.xlsm` |
| `db/` | База данных: `schema.sql` |
| `docs/` | Документация: `ARCHITECTURE.md`, `DEVELOPER.md`, `git-workflow.md`, `sourcecraft-guide.md`, `ROADMAP.md`, `table.md`, `CHANGELOG.md` |
| `logs/` | Лог-файлы |
| `plans/` | Планы, отчёты, решения |
| `plans/_archive/` | Архив выполненных планов |
| `scripts/` | Python/PowerShell скрипты автоматизации |
| `src/classes/` | VBA-классы (`.cls`) – `PartIdentity.cls`, `WorkIdentity.cls`, `WorkEntry.cls`, `IModelDataProvider.cls`, `Mod_ModelDBProvider.cls`, `Mod_SQLiteDB.cls` |
| `src/modules/` | VBA-модули (`.bas`) – основная бизнес-логика |
| `src/sheets/` | VBA-классы листов (`.cls`) |
| `.codeassistant/rules/` | Рабочие правила рабочей области (зеркало `.ycarules`) |
| `.github/` | GitHub Actions |
| `.vscode/` | Настройки VS Code |
| Корень | `work.xlsm`, `report.xlsx`, `README.md`, `.ycarules`, `.gitattributes`, `.gitignore`, `.codeassistantignore` |

## [S2] VBA-МОДУЛИ (`src/modules/`)

| Модуль | Назначение |
|--------|------------|
| `Mod_AutoMatch.bas` | Автоматическое сопоставление данных |
| `Mod_ButtonDispatcher.bas` | Диспетчер кнопок |
| `Mod_Constants.bas` | Константы |
| `Mod_FullTestRunner.bas` | Запуск тестов |
| `Mod_Import.bas` | Импорт данных |
| `Mod_Logger.bas` | Логирование |
| `Mod_ModelDB.bas` | Работа с базой моделей |
| `Mod_ModelTypes.bas` | Типы данных |
| `Mod_OrderHeader.bas` | Шапка заказа |
| `Mod_PickWork.bas` | Выбор работ |
| `Mod_SheetButtons.bas` | Кнопки листов |
| `Mod_SheetOps.bas` | Операции с листами |
| `Mod_Utils.bas` | Утилиты |

## [S3] КЛАССЫ (`src/classes/`)

| Класс | Назначение |
|-------|------------|
| `PartIdentity.cls` | Объект детали |
| `WorkIdentity.cls` | Объект работы |
| `WorkEntry.cls` | Элемент работы |
| `IModelDataProvider.cls` | Интерфейс провайдера данных модели |
| `Mod_ModelDBProvider.cls` | Провайдер базы моделей |
| `Mod_SQLiteDB.cls` | Работа с SQLite |

## [S4] POWERSHELL-СКРИПТЫ (`scripts/`)

- `config.ps1` – конфигурация.
- `Set-ExcelTrust.ps1` – настройка доверия Excel.
- `fix_vbom_and_venv.ps1` – исправление VBOM и виртуального окружения.
- `monitor_long.ps1` – мониторинг длительных процессов.
- Кодировка – **UTF‑8 с BOM**.
- **Запрещено** изменять VBA-модули напрямую – только через скрипты импорта/экспорта.

## [S5] PYTHON-СКРИПТЫ (`scripts/`)

Полный перечень (актуальный список см. в каталоге `scripts/`):
`apply_protection_templates.py`, `apply_sheet_format.py`, `build_all.py`, `build_global_parts.py`,
`build_templates.py`, `check_docs.py`, `check_vba_syntax.py`, `check_z4_fallback.py`,
`clean_system.py`, `config.py`, `export_vba.py`, `impVBA.py`, `initiate_models.py`,
`migrate_models_to_sqlite.py`, `run_oab_reconcile_business_test.py`, `run_p1_business_test.py`,
`run_tests.py`, `sqlite_schema.py`, `template_protection.py`, `update_version.py`, `watch_p2.py`.

## [S6] МОДЕЛИ (`base/models/`)

- `4x4.xlsm`, `2170.xlsm`, `2180.xlsm`, `2190.xlsm`, `GAZ.xlsm`, `UAZ.xlsm`.

## [S7] ОЧИСТКА ПРОЕКТА

Перед завершением задачи удалить временные файлы: `__pycache__`, `*.sync-conflict-*`, `~$*.xlsm`.
Проект должен оставаться чистым.

## [S8] ОБЪЕДИНЕНИЕ КОМАНД

Если пользователь должен выполнить несколько однотипных команд – объединить их в одну (через `&&` или скрипт).

## [S9] ПРИОРИТЕТ ИСПОЛЬЗОВАНИЯ POWERSHELL 7

**Все команды терминала и скрипты в среде Windows выполняются через PowerShell 7** (`pwsh.exe`).
- `cmd.exe` и `activate.bat` **не используются**.
- Виртуальное окружение активируется исключительно через `.venv\Scripts\Activate.ps1` (dot-source).
- Все PowerShell-скрипты и команды должны быть написаны с учётом совместимости с PowerShell 7.
- При вызове команд используйте `pwsh` вместо `powershell` (если не требуется обратная совместимость с Windows PowerShell 5.1).
- В случае, если скрипт требует старой версии, это должно быть явно оговорено в комментариях.
- **Профиль терминала VS Code `SourceCraft`** использует PowerShell 7 (`pwsh.exe`) с автоактивацией venv через `.venv\Scripts\Activate.ps1`.