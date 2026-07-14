# Отчёт: Механизм глобальных правил SourceCraft / Roo Code

## 1. Общая архитектура

**SourceCraft Code Assistant** (v0.11.127) — это форк/вариация **Roo Code**, распространяемый как расширение VS Code под идентификатором `yandex-cloud.yandex-code-assist`. Механизм правил унаследован от Roo Code.

Система правил разделена на **два уровня**:

| Уровень | Файл | Где находится | Область действия |
|---------|------|---------------|------------------|
| **Проектный (project)** | `.ycarules` | Корень проекта (`L:/PROject/SysW/.ycarules`) | Только текущий проект |
| **Глобальный (global)** | `.ycarules` | В `globalStoragePath` расширения | Все проекты |

## 2. Текущее состояние в системе

### 2.1. Проектный `.ycarules`

**Найден** и активен:
- **Путь:** [`L:/PROject/SysW/.ycarules`](.ycarules)
- **Содержит:** правила поведения для агентов, правила кодировки VBA, запреты и т.д.

### 2.2. Глобальный `.ycarules`

**Не найден.** Поиск выполнялся по следующим путям:
- `%USERPROFILE%` — не найден
- `%APPDATA%` — не найден
- `%LOCALAPPDATA%` — не найден
- `%USERPROFILE%\.vscode` — не найден

### 2.3. Другие файлы правил

- `.roomodes` — не найден (глобальные кастомные режимы)
- `.clinerules` — не найден (наследие Claude Code)
- `.windsurfrules` — не найден (наследие Windsurf)

## 3. Как работает механизм загрузки правил (по данным анализа кода)

На основе анализа обфусцированного кода [`extension.js`](C:\Users\root\.vscode\extensions\yandex-cloud.yandex-code-assist-0.11.127\out\extension.js):

### 3.1. Загрузка `.ycarules`

Расширение ищет файл `.ycarules` в **корне открытого проекта** (рабочей директории). Если файл найден — его содержимое добавляется в системный промпт как `customInstructions`.

### 3.2. Глобальные инструкции (Global Instructions)

В коде обнаружена строка (около line 258):
```
Global Instructions: ${...}
```
Это указывает на то, что расширение поддерживает **глобальные инструкции**, которые передаются через параметр `globalCustomInstructions` (line 6586):
```
globalCustomInstructions:I
```

### 3.3. Глобальное хранилище (globalStoragePath)

Расширение использует `globalStoragePath` VS Code для хранения глобальных данных. Стандартный путь для глобального хранилища расширений VS Code:

```
%APPDATA%\Code\User\globalStorage\yandex-cloud.yandex-code-assist\
```

или если задан `customStoragePath` (настраивается в `settings.json`):
```json
"sourcecraft-code-assist.customStoragePath": "путь_к_кастомному_хранилищу"
```

### 3.4. Разделение project/global для custom modes

В коде (line 5925) обнаружено:
```javascript
let n = e.endsWith(QY) ? "project" : "global";
```
Это указывает, что файлы, оканчивающиеся на определённый суффикс (вероятно, `.roomodes`), считаются проектными, а все остальные — глобальными.

## 4. Как добавить глобальные правила

### Способ 1: Создать глобальный `.ycarules` в `globalStoragePath`

Создать файл:
```
%APPDATA%\Code\User\globalStorage\yandex-cloud.yandex-code-assist\.ycarules
```

С содержимым, например:
```markdown
# Глобальные правила для всех проектов
- Всегда использовать UTF-8 для VBA-файлов
- ...
```

### Способ 2: Использовать `customStoragePath` (если настроен)

Если в `settings.json` VS Code задан кастомный путь:
```json
"sourcecraft-code-assist.customStoragePath": "C:/Users/root/.sourcecraft"
```

Тогда глобальный `.ycarules` нужно разместить по этому пути.

### Способ 3: Через настройки VS Code (если поддерживается)

Проверить, нет ли в `settings.json` ключа для глобальных инструкций:
```json
"sourcecraft-code-assist.globalInstructions": "..."
```
(На данный момент такой ключ **не обнаружен** в `package.json` расширения.)

## 5. Рекомендация для данного проекта

В текущей конфигурации **глобальный `.ycarules` отсутствует**, и все правила задаются через проектный файл [`L:/PROject/SysW/.ycarules`](.ycarules).

Если требуется добавить правило, которое будет применяться **ко всем проектам**, рекомендуется:

1. Создать директорию глобального хранилища (если не существует):
   ```
   %APPDATA%\Code\User\globalStorage\yandex-cloud.yandex-code-assist\
   ```

2. Создать в ней файл `.ycarules` с глобальными правилами.

3. Перезапустить VS Code.

## 6. Проверка работоспособности

Чтобы убедиться, что глобальные правила загружаются, можно:
1. Добавить тестовое правило в глобальный `.ycarules`
2. Открыть другой проект (не SysW)
3. Проверить в логах расширения (`%APPDATA%\Code\logs\...\exthost\yandex-cloud.yandex-code-assist\`), загрузилось ли правило

## 7. Ссылки

- **Проектный `.ycarules`:** [`L:/PROject/SysW/.ycarules`](.ycarules)
- **Расширение:** `C:\Users\root\.vscode\extensions\yandex-cloud.yandex-code-assist-0.11.127`
- **Логи расширения:** `%APPDATA%\Code\logs\*\window1\exthost\yandex-cloud.yandex-code-assist\`
- **Глобальное хранилище (предполагаемое):** `%APPDATA%\Code\User\globalStorage\yandex-cloud.yandex-code-assist\`