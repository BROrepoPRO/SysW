# План: Адаптированный `.vscode/settings.json` для проекта SysW

## Контекст

Проект `L:\PROject\SysW` — VBA-разработка для Excel (`work.xlsm`). Файлы `.bas` (5 шт.) содержат кракозябры из-за двойной перекодировки UTF-8 → Windows-1251 → UTF-8. Файл `Sheet1_main.cls` — корректный UTF-8. Скрипт `import_vba.ps1` конвертирует UTF-8 → Windows-1251 при импорте в Excel.

Текущий `.vscode/settings.json` минимален — не хватает настроек терминала, VBA LSP, исключений.

## Задача

Создать новый `.vscode/settings.json` со следующими блоками:

---

### Блок 1: Глобальные настройки кодировки

```json
"files.encoding": "utf8",
"files.autoGuessEncoding": true,
"files.eol": "\n",
"files.trimTrailingWhitespace": true
```

- `files.autoGuessEncoding: true` — добавлен для безопасности, чтобы VS Code сам определял кодировку, если файл вдруг окажется не в UTF-8
- `files.trimTrailingWhitespace: true` — держит код чистым

---

### Блок 2: Терминал — профиль SourceCraft с UTF-8

```json
"terminal.integrated.profiles.windows": {
    "SourceCraft": {
        "path": "powershell.exe",
        "args": [
            "-NoExit",
            "-Command",
            "chcp 65001 | Out-Null"
        ],
        "icon": "terminal-powershell",
        "color": "terminal.ansiGreen"
    }
},
"terminal.integrated.defaultProfile.windows": "SourceCraft",
"terminal.integrated.cwd": "L:/PROject/SysW"
```

**Почему так:**
- На Windows 10 `chcp` по умолчанию = 866 (OEM-кириллица). Это вызывает кракозябры в выводе PowerShell.
- Профиль `SourceCraft` принудительно переключает кодировку терминала на UTF-8 (65001) при каждом запуске.
- `terminal.integrated.cwd` сразу открывает терминал в корне проекта.
- Используется `powershell.exe` (PowerShell 5), т.к. `import_vba.ps1` использует COM-объекты Excel, которые работают только в Windows PowerShell.

---

### Блок 3: VBA LSP

```json
"vba.enable": true,
"vba.languageServer.enable": true,
"vba.excel.path": "L:/PROject/SysW/work.xlsm",
"vba.exportOnSave": true
```

- Включает поддержку VBA: автодополнение, подсветка синтаксиса, навигация.
- `vba.excel.path` указывает на целевую книгу.
- `vba.exportOnSave: true` — автоматический экспорт модулей при сохранении.

---

### Блок 4: Специфичные настройки для VBA-файлов

```json
"[bas]": {
    "files.encoding": "utf8"
},
"[cls]": {
    "files.encoding": "utf8"
}
```

- `.bas` — стандартные модули VBA (5 файлов).
- `.cls` — class-модули (`Sheet1_main.cls`).
- Кодировка `utf8` для корректной работы с кириллицей.

---

### Блок 5: Исключения

```json
"files.exclude": {
    "workOt/": true,
    "**/*.xlsm": true,
    "**/*.xlsx": true
},
"search.exclude": {
    "workOt/": true,
    "**/*.xlsm": true,
    "**/*.xlsx": true,
    "check_hex.ps1": true
}
```

- `workOt/` — директория с шаблонами Excel.
- `*.xlsm`, `*.xlsx` — бинарные файлы Excel, не редактируются в коде.
- `check_hex.ps1` — утилита для HEX-дампа, не нужна в поиске.
- `import_vba.ps1` **не исключается** — это основной рабочий скрипт.

---

## Полный файл `.vscode/settings.json`

```json
{
    // ============================================================
    // Блок 1: Глобальные настройки кодировки и форматирования
    // ============================================================
    "files.encoding": "utf8",
    "files.autoGuessEncoding": true,
    "files.eol": "\n",
    "files.trimTrailingWhitespace": true,

    // ============================================================
    // Блок 2: Настройки терминала
    // Профиль SourceCraft с принудительной UTF-8 кодировкой (chcp 65001)
    // Решает проблему кракозябр в выводе PowerShell и совместимости
    // с VBA-скриптами, работающими через COM-объекты Excel.
    // ============================================================
    "terminal.integrated.profiles.windows": {
        "SourceCraft": {
            "path": "powershell.exe",
            "args": [
                "-NoExit",
                "-Command",
                "chcp 65001 | Out-Null"
            ],
            "icon": "terminal-powershell",
            "color": "terminal.ansiGreen"
        }
    },
    "terminal.integrated.defaultProfile.windows": "SourceCraft",
    "terminal.integrated.cwd": "L:/PROject/SysW",

    // ============================================================
    // Блок 3: Настройки VBA-расширения (vba-lsp)
    // Обеспечивает автодополнение, подсветку синтаксиса,
    // навигацию по коду VBA-модулей.
    // ============================================================
    "vba.enable": true,
    "vba.languageServer.enable": true,
    "vba.excel.path": "L:/PROject/SysW/work.xlsm",
    "vba.exportOnSave": true,

    // ============================================================
    // Блок 4: Специфичные настройки для VBA-файлов
    // .bas — стандартные модули VBA
    // .cls — class-модули (в т.ч. Sheet1_main.cls)
    // Кодировка utf8 для корректной работы с кириллицей.
    // ============================================================
    "[bas]": {
        "files.encoding": "utf8"
    },
    "[cls]": {
        "files.encoding": "utf8"
    },

    // ============================================================
    // Блок 5: Исключения для проводника и поиска
    // workOt/ — временные/шаблонные файлы Excel
    // *.xlsm, *.xlsx — бинарные файлы Excel (не редактируются в коде)
    // check_hex.ps1 — утилита для дампа HEX (не нужна в поиске)
    // import_vba.ps1 НЕ исключается — это основной рабочий скрипт
    // ============================================================
    "files.exclude": {
        "workOt/": true,
        "**/*.xlsm": true,
        "**/*.xlsx": true
    },
    "search.exclude": {
        "workOt/": true,
        "**/*.xlsm": true,
        "**/*.xlsx": true,
        "check_hex.ps1": true
    }
}
```

## Порядок применения

1. Переключиться в **Code mode**
2. Открыть файл `.vscode/settings.json`
3. Заменить содержимое на полный JSON выше
4. Сохранить файл
5. Перезагрузить VS Code (чтобы применились настройки терминала)
6. Открыть терминал — должен автоматически запуститься профиль `SourceCraft` с `chcp 65001`