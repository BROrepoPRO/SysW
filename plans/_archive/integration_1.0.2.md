# План интеграции обновлений — релиз v1.0.2

## Цель релиза

Релиз v1.0.2 объединяет несколько независимых улучшений проекта SysW:

1. **Исправление загрузки виртуального окружения Python** в терминале VS Code (профиль `SourceCraft`).
2. **Полное удаление Syncthing** из проекта (файлы синхронизации и все упоминания в документации).
3. **Проверка и закрепление конфигурации MCP-серверов** (`filesystem`, `git`) с дополнением документации.
4. **Обновление номера версии** с 1.0.0 до 1.0.2 во всех точках проекта.
5. **Очистка проекта от мусора** и обновление документации.

Высший приоритет — соблюдение правил `.ycarules`. Исходное задание — `plans/_archive/promt1.0.2mcp.md`. Согласие пользователя на реализацию получено.

---

## Раздел 1. Виртуальное окружение Python

**Проблема:** профиль терминала `SourceCraft` в [`.vscode/settings.json`](.vscode/settings.json) использует относительный путь
`args: ["/K", ".venv\\Scripts\\activate.bat"]`. Если рабочая директория терминала не совпадает с корнем проекта — путь не резолвится, и venv не активируется.

**Факт:** окружение `.venv` в `L:\PROject\SysW\.venv` полностью здорово (Python 3.14.6, pywin32/win32com на месте).

**Правка:**

1. Открыть [`.vscode/settings.json`](.vscode/settings.json).
2. В профиле `SourceCraft` заменить строку:
   ```json
   "args": ["/K", ".venv\\Scripts\\activate.bat"]
   ```
   на:
   ```json
   "args": ["/K", "${workspaceFolder}\\.venv\\Scripts\\activate.bat"]
   ```
3. Сохранить файл.

**Проверка:** открыть новый терминал VS Code с профилем `SourceCraft` и убедиться, что командная строка показывает активацию venv (например, префикс `(.venv)`).

---

## Раздел 2. Удаление Syncthing

Пользователь явно согласился на полное удаление всего, что связано с синхронизацией.

**Удаление файлов:**

1. Удалить файл [`scripts/sync_fix.ps1`](scripts/sync_fix.ps1).
2. Удалить файл [`.stignore`](.stignore).

**Вычистка упоминаний из документации:**

Выполнить поиск упоминаний `Syncthing`, `sync_fix`, `syncthing`, `.stignore` в следующих файлах и удалить связанные фрагменты (секции, пункты, абзацы):

1. [`docs/CHANGELOG.md`](docs/CHANGELOG.md)
2. [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)
3. [`docs/DEVELOPER.md`](docs/DEVELOPER.md)
4. [`docs/ROADMAP.md`](docs/ROADMAP.md)
5. [`docs/git-workflow.md`](docs/git-workflow.md)

**Проверка:** глобальный поиск по репозиторию не должен возвращать совпадений по `Syncthing|sync_fix|stignore` (кроме истории git).

---

## Раздел 3. MCP-серверы

**Факт:** в [`.sourcecraft`](.sourcecraft) уже прописаны два MCP-сервера:
- `filesystem`: `npx -y @modelcontextprotocol/server-filesystem "${workspaceFolder}"`
- `git`: `npx -y @modelcontextprotocol/server-git --repository "${workspaceFolder}"`

**Действия:**

1. Проверить, что обе записи корректны и используют переменную `${workspaceFolder}` (без захардкоженных абсолютных путей).
2. При необходимости закрепить/выровнять формат записи.
3. Дополнить [`docs/sourcecraft-guide.md`](docs/sourcecraft-guide.md):
   - описание серверов `filesystem` и `git`;
   - примеры использования (список файлов, статус git);
   - примечание про порядок запуска `npx -y`.

**Проверка:** `.sourcecraft` содержит валидный JSON; guide описывает оба сервера и их применение.

---

## Раздел 4. Обновление версии

**Текущая версия:** 1.0.0. **Целевая версия:** 1.0.2.

**Шаги:**

1. Запустить скрипт обновления версии из корня проекта:
   ```bash
   python scripts/update_version.py 1.0.2
   ```
   (на Windows — `python scripts\update_version.py 1.0.2`; при необходимости использовать `py`).

2. Скрипт автоматически обновит файлы:
   - [`src/modules/Mod_Constants.bas`](src/modules/Mod_Constants.bas)
   - [`scripts/config.py`](scripts/config.py)
   - [`scripts/config.ps1`](scripts/config.ps1)
   - [`README.md`](README.md)
   - [`docs/DEVELOPER.md`](docs/DEVELOPER.md)
   - [`docs/ROADMAP.md`](docs/ROADMAP.md)
   - [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)

3. Проверить результат в [`docs/CHANGELOG.md`](docs/CHANGELOG.md). Если скрипт не создал запись — добавить раздел вручную в начало списка релизов:

   ```markdown
   ## [v1.0.2]

   ### Изменения

   - Исправлена активация виртуального окружения Python в терминале VS Code.
   - Полное удаление Syncthing (скрипт `sync_fix.ps1`, `.stignore`, упоминания в docs).
   - Проверка и документирование MCP-серверов (`filesystem`, `git`).
   - Обновление номера версии до 1.0.2.
   - Очистка проекта от временных/мусорных файлов.
   ```

4. Убедиться, что все точки версии в проекте показывают 1.0.2.

---

## Раздел 5. Документация и чистота

**Обновление документации:**

1. Проверить согласованность упоминаний версии 1.0.2 в `docs/*.md` и `README.md`.
2. Убедиться, что в `docs` не осталось ссылок на удалённые файлы (`sync_fix.ps1`, `.stignore`).

**Очистка проекта от мусора:**

1. Удалить все каталоги `__pycache__` (рекурсивно, кроме `node_modules`).
2. Удалить файлы вида `*.sync-conflict-*`.
3. Удалить временные файлы Excel `~$*.xlsm`.
4. Удалить файлы `*.bak` (например, `src/sheets/Sheet_work.cls.bak`, `src/sheets/Sheet_z4.cls.bak`) — по согласованию с пользователем (если не требуются для восстановления).

**Проверка чистоты:** в дереве файлов не должно быть `__pycache__`, `*.sync-conflict-*`, `~$*.xlsm`, `*.bak`.

---

## Раздел 6. Коммит и пуш

Все изменения фиксируются единым релизным коммитом в Conventional Commit стиле.

1. Проверить статус и подготовить индексацию:
   ```bash
   git add -A
   git status
   ```

2. Просмотреть дифф, чтобы убедиться в отсутствии лишних файлов:
   ```bash
   git diff --cached --stat
   ```

3. Создать коммит (пример, тип `fix`, т.к. релиз содержит исправление venv):
   ```bash
   git commit -m "fix: release v1.0.2 — venv activation, remove syncthing, mcp docs"
   ```
   > Если релиз в первую очередь про новый функционал MCP — допустим тип `feat`. Ключевое изменение — исправление venv, поэтому рекомендуется `fix`.

4. Отправить изменения в удалённый репозиторий:
   ```bash
   git push
   ```

5. При необходимости (по правилам проекта) — создать git-тег `v1.0.2`:
   ```bash
   git tag v1.0.2
   git push --tags
   ```

---

## Проверка перед завершением (чек-лист)

- [ ] [`.vscode/settings.json`](.vscode/settings.json): профиль `SourceCraft` использует `${workspaceFolder}\\.venv\\Scripts\\activate.bat`.
- [ ] Терминал VS Code активирует venv без ошибок.
- [ ] [`scripts/sync_fix.ps1`](scripts/sync_fix.ps1) удалён.
- [ ] [`.stignore`](.stignore) удалён.
- [ ] В `docs/*.md` и `README.md` нет упоминаний Syncthing / sync_fix / .stignore.
- [ ] [`.sourcecraft`](.sourcecraft) содержит валидные записи `filesystem` и `git`.
- [ ] [`docs/sourcecraft-guide.md`](docs/sourcecraft-guide.md) дополнен описанием MCP-серверов.
- [ ] Версия обновлена до 1.0.2 во всех точках (Mod_Constants.bas, config.py, config.ps1, README.md, docs).
- [ ] В [`docs/CHANGELOG.md`](docs/CHANGELOG.md) создан раздел `## [v1.0.2]`.
- [ ] Проект очищен от мусора (`__pycache__`, `*.sync-conflict-*`, `~$*.xlsm`, `*.bak`).
- [ ] `git add -A` выполнен, `git status` чист и соответствует ожиданиям.
- [ ] Коммит создан в Conventional Commit стиле (`feat`/`fix`).
- [ ] Изменения отправлены командой `git push`.