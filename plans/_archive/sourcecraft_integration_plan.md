# План: Интеграция .sourcecraft в проект SysW

## Цель

Создать корневой файл `.sourcecraft` (JSON), который:
1. Содержит конфигурацию MCP-серверов с **относительными путями** (вместо абсолютных `L:\PROject\SysW`)
2. Интегрирует правила из `.ycarules` как единую конфигурацию
3. Заменяет `.vscode/mcp.json` (который будет удалён)

---

## 1. Структура `.sourcecraft`

Файл создаётся в корне проекта: `l:/PROject/SysW/.sourcecraft`

```jsonc
{
  "$schema": "https://raw.githubusercontent.com/RooVetGit/Roo-Code/main/schemas/sourcecraft.json",
  "mcpServers": {
    "filesystem": {
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-filesystem",
        "${workspaceFolder}"
      ]
    },
    "git": {
      "command": "npx",
      "args": [
        "-y",
        "@modelcontextprotocol/server-git",
        "--repository",
        "${workspaceFolder}"
      ]
    }
  },
  "customInstructions": {
    "file": ".ycarules"
  },
  "rules": {
    "exclude": [
      "*.log",
      "_temp_export/*",
      "_temp_import/*",
      "__pycache__/*",
      "*.pyc",
      "*.pyo",
      ".venv/*",
      "venv/*",
      ".DS_Store",
      "Thumbs.db",
      "Desktop.ini",
      "node_modules/*",
      "workOt/*",
      "*.sync-conflict-*",
      "logs/*.log",
      "logs/*.txt"
    ],
    "include": [
      "docs/DEVELOPER.md",
      "docs/ARCHITECTURE_SQLITE.md",
      "src/modules/Mod_Constants.bas"
    ],
    "critical": [
      "work.xlsm",
      "table.md",
      "report.xlsx",
      "README.md",
      "CHANGELOG.md",
      ".ycarules",
      ".sourcecraft",
      "plans/ROADMAP.md",
      "base/",
      "base/models/"
    ]
  }
}
```

### Пояснение полей

| Поле | Назначение |
|------|-----------|
| `$schema` | Схема валидации JSON (опционально) |
| `mcpServers` | Конфигурация MCP-серверов с `${workspaceFolder}` — переменная, которая резолвится в корень проекта независимо от абсолютного пути на диске |
| `customInstructions.file` | Ссылка на `.ycarules` — файл с правилами для ассистента. SourceCraft автоматически подгрузит его содержимое как кастомные инструкции |
| `rules.exclude` | Паттерны исключений из `.ycarules` (секция `[E1]`) |
| `rules.include` | Паттерны обязательного включения из `.ycarules` (секция `[E2]`) |
| `rules.critical` | Критические файлы из `.ycarules` (секция `[E3]`) |

---

## 2. Что происходит с `.ycarules`

Файл `.ycarules` **остаётся** в проекте. Он подключается через `customInstructions.file`.

**Почему не объединяем всё в один файл:**
- `.ycarules` содержит развёрнутую документацию (легенду, описания правил, примеры) — это не чистая конфигурация, а инструкции для ассистента
- `.sourcecraft` — это машиночитаемый JSON-конфиг
- Разделение конфигурации и инструкций — лучшая практика

**Дублирование правил:**
- `rules.exclude`, `rules.include`, `rules.critical` в `.sourcecraft` дублируют соответствующие секции из `.ycarules` для того, чтобы SourceCraft мог применять их на уровне платформы (фильтрация файлов, защита от изменений)
- `.ycarules` остаётся единственным источником истины для текстовых инструкций ассистенту

---

## 3. Что происходит с `.vscode/mcp.json`

Файл `.vscode/mcp.json` **удаляется**, так как его функциональность полностью перенесена в `.sourcecraft`.

**Важно:** Если пользователь использует VS Code без SourceCraft (например, для обычной разработки), MCP-серверы не будут доступны. Это приемлемо, так как MCP-серверы нужны только для работы SourceCraft.

---

## 4. Обновление `.vscode/settings.json`

Добавить ссылку на `.sourcecraft` (если SourceCraft поддерживает такую настройку). Если нет — изменений не требуется.

---

## 5. Обновление документации

- `docs/sourcecraft-guide.md` — добавить раздел о `.sourcecraft`, обновить структуру проекта
- `CHANGELOG.md` — добавить запись о v0.16.0 (уже сделано)

---

## 6. Git-коммит

После всех изменений выполнить коммит:

```
chore(ci): add .sourcecraft configuration with MCP servers and ycarules integration
```

---

## Чеклист задач

- [ ] Создать файл `.sourcecraft` в корне проекта с MCP-серверами (относительные пути `${workspaceFolder}`)
- [ ] Добавить в `.sourcecraft` секцию `customInstructions.file` со ссылкой на `.ycarules`
- [ ] Добавить в `.sourcecraft` секцию `rules` (exclude, include, critical) на основе `.ycarules`
- [ ] Удалить файл `.vscode/mcp.json`
- [ ] Обновить `docs/sourcecraft-guide.md` (структура проекта, раздел о `.sourcecraft`)
- [ ] Обновить `CHANGELOG.md` (запись о v0.16.0 уже есть — проверить актуальность)
- [ ] Выполнить `git add .sourcecraft .vscode/mcp.json CHANGELOG.md docs/sourcecraft-guide.md`
- [ ] Выполнить `git commit` с сообщением в формате Conventional Commits
- [ ] Выполнить `git push`