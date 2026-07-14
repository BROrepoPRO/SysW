# Git Workflow — SysW

## Веточная стратегия

```
main  ───●────────────●────────── (стабильная, только релизы)
          \          /
dev   ─────●──●──●──●──────────── (разработка, интеграция)
              \
feature/fix ───●──●────────────── (рабочие ветки от dev)
```

| Ветка | Назначение | Кто пушит |
|-------|-----------|-----------|
| `main` | Стабильная версия, только релизы. Защищена от прямых пушей | Только через PR из `dev` |
| `dev` | Интеграционная ветка разработки. Все изменения проходят через неё | Разработчики |
| `fix/*` | Ветки для исправлений (например, `fix/python-import`) | Разработчики |
| `feature/*` | Ветки для новых функций (например, `feature/new-report`) | Разработчики |

## Формат сообщений коммитов

```
тип(область): краткое описание в повелительном наклонении

Тело сообщения (опционально) — подробное описание изменений,
их причин и последствий.
```

### Типы коммитов

| Тип | Когда использовать |
|-----|-------------------|
| `feat` | Новая функциональность |
| `fix` | Исправление ошибки |
| `docs` | Изменения в документации |
| `chore` | Обслуживание (настройка CI, .gitignore, .gitattributes) |
| `refactor` | Рефакторинг без изменения функциональности |
| `test` | Добавление или изменение тестов |
| `style` | Форматирование, отступы, точки с запятой (не CSS) |

### Примеры

```
feat(import): add order number validation before import
fix(orderheader): handle empty customer name field
docs(git): add branching strategy documentation
chore(ci): add GitHub Actions workflow for VBA checks
refactor(utils): extract string formatting to separate function
test(runner): add test case for empty order list
```

### Области (scope) для проекта SysW

| Область | Описание |
|---------|----------|
| `import` | Mod_Import, импорт данных |
| `orderheader` | Mod_OrderHeader, заполнение шапки |
| `utils` | Mod_Utils, утилиты |
| `buttons` | Mod_ButtonDispatcher, кнопки |
| `tests` | Mod_FullTestRunner, run_tests.py |
| `vba` | Любые изменения в .bas/.cls файлах |
| `docs` | Документация |
| `ci` | GitHub Actions, CI/CD |
| `scripts` | Python/PowerShell скрипты автоматизации |

## Что делать перед коммитом

### 1. Проверить статус

```bash
git status
```

Убедиться, что:
- Нет случайных файлов (`.tmp`, `_temp_*`, `__pycache__`)
- `work.xlsm` **НЕ** отображается в списке изменений (он в `.gitignore`)
- Все нужные изменения проиндексированы

### 2. Синхронизировать VBA-модули (если работали в Excel)

Если изменения вносились в Excel/VBA Editor:

```bash
# Windows: экспорт VBA-модулей из Excel на диск
python export_vba.py
```

### 3. Обновить CHANGELOG.md

Добавить запись в `CHANGELOG.md` в формате [Keep a Changelog](https://keepachangelog.com/ru/1.0.0/):

```markdown
## [версия] — YYYY-MM-DD

### Добавлено
- ...

### Изменено
- ...

### Исправлено
- ...
```

### 4. Проверить кодировку

VBA-файлы должны быть в UTF-8 (без BOM). Проверка:

```bash
python3 -c "
with open('Mod_Utils.bas', 'rb') as f:
    raw = f.read()
raw.decode('utf-8')  # вызовет ошибку, если не UTF-8
print('OK: UTF-8')
"
```

### 5. Проиндексировать и закоммитить

```bash
git add <файлы>
git commit -m "тип(область): описание"
```

### 6. Отправить на GitHub

```bash
# Если ветка новая — первый пуш с -u
git push -u origin dev

# Если ветка уже существует
git push
```

## Синхронизация с GitHub

### Забрать изменения из remote

```bash
# Забрать изменения из origin/dev
git pull origin dev

# Если есть конфликты — разрешить и закоммитить
```

### Pull Request (PR)

1. Создать PR из `dev` → `main` через интерфейс GitHub
2. Дождаться прохождения GitHub Actions checks
3. Получить approval от ревьюера
4. Merge через "Squash and merge" (один коммит в `main`)

## Полезные команды

```bash
# Просмотр лога
git log --oneline --graph --all

# Просмотр изменений
git diff
git diff --staged

# Отмена индексации файла
git restore --staged <файл>

# Переключение веток
git checkout dev
git checkout -b fix/some-bug  # создать и переключиться

# Отмена незакоммиченных изменений (осторожно!)
git restore <файл>
```

## .gitignore

Файл `.gitignore` уже настроен и игнорирует:
- `work.xlsm` — Excel-файл с макросами (бинарный, не подлежит контролю версий)
- `_temp_export/`, `_temp_import/` — временные директории скриптов
- `__pycache__/`, `*.pyc`, `*.pyo` — кэш Python
- `.venv/`, `venv/` — виртуальное окружение Python
- `*.tmp`, `*.log`, `~$*.xls*` — временные и лог-файлы
- `Thumbs.db`, `Desktop.ini`, `.DS_Store` — системные файлы ОС

**Не добавляйте эти файлы в Git!**