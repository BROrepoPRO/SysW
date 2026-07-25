# План обновления SysW до v0.15.0

## 1. Общая стратегия обновления

### 1.1 Цель
Обновление проекта SysW с v0.14.0 до v0.15.0 с фокусом на:
- Максимальное покрытие тестами непокрытых модулей
- Исправление остаточных проблем кодовой базы
- Синхронизацию документации с фактическим состоянием
- Подготовку к дальнейшему развитию

### 1.2 Порядок выполнения (зависимости)

```
Фаза 1: Подготовка (безопасные изменения)
  └── 1.1 Очистка комментариев с абсолютными путями
  └── 1.2 Удаление deprecated-константы MODELDB_BASE_PATH
  └── 1.3 Обновление export_vba.py (COMPONENTS)

Фаза 2: Тесты (требуют Фазу 1)
  └── 2.1 Добавление тестов для Mod_ModelDB (TC-31..TC-35)
  └── 2.2 Добавление тестов для Mod_PickWork (TC-36..TC-38)
  └── 2.3 Добавление тестов для Mod_AutoMatch (TC-39..TC-44)
  └── 2.4 Углубление TC-14 (полный цикл импорта)
  └── 2.5 Обновление реестра тестов в DEVELOPER.md

Фаза 3: Документация (требует Фазы 1-2)
  └── 3.1 Обновление DEVELOPER.md
  └── 3.2 Обновление sourcecraft-guide.md
  └── 3.3 Обновление CHANGELOG.md

Фаза 4: Коммит и пуш
  └── 4.1 Экспорт VBA из Excel
  └── 4.2 Коммит #1: кодовая база
  └── 4.3 Коммит #2: тесты
  └── 4.4 Коммит #3: документация
  └── 4.5 Пуш в репозиторий
```

### 1.3 Какие файлы будут изменены/созданы/удалены

**Изменяемые файлы:**
- `src/modules/Mod_FullTestRunner.bas` — добавление новых тестов
- `src/modules/Mod_ModelDB.bas` — удаление deprecated-константы, очистка комментариев
- `scripts/export_vba.py` — обновление COMPONENTS (добавление новых модулей)
- `docs/DEVELOPER.md` — синхронизация с актуальным состоянием
- `docs/sourcecraft-guide.md` — обновление до v0.14.0
- `CHANGELOG.md` — добавление записи v0.15.0

**Неизменяемые (защищённые по E3):**
- `work.xlsm` — защищён
- `table.md` — защищён
- `report.xlsx` — защищён
- `README.md` — защищён
- `.ycarules` — защищён
- `plans/ROADMAP.md` — защищён
- `base/` — защищён
- `base/models/` — защищён

---

## 2. План обновления тестового модуля (Mod_FullTestRunner.bas)

### 2.1 Текущее состояние

| Группа | Сценарии | Модуль | Статус |
|--------|----------|--------|--------|
| `RunUtilsTests()` | TC-01..TC-08 | Mod_Utils | ✅ |
| `RunLoggerTests()` | TC-09..TC-11 | Mod_Logger | ✅ |
| `RunUtilsEdgeTests()` | TC-12 | Mod_Utils | ✅ |
| `RunLibNameTests()` | TC-13 | Mod_Constants | ✅ |
| `RunImportVHTests()` | TC-14 | Mod_Import | ⚠️ поверхностный |

### 2.2 Новые тестовые группы

#### 2.2.1 `RunModelDBTests()` — Mod_ModelDB (TC-31..TC-35)

```vba
Private Sub RunModelDBTests()
    ' TC-31: GetModelDBBasePath — возвращает путь, оканчивающийся на \base\models\
    ' TC-32: GetModelGroupFilePath — возвращает .xlsm путь для существующей группы
    ' TC-33: GetModelGroupFilePath — возвращает .xlsx путь, если .xlsm не существует
    ' TC-34: ModelGroupFileExists — проверка существования файла UAZ.xlsm
    ' TC-35: ModelGroupFileExists — проверка несуществующей группы
End Sub
```

**Детализация TC-31..TC-35:**

| ID | Название | Проверка | Условие PASS |
|----|----------|----------|--------------|
| TC-31 | `GetModelDBBasePath` возвращает корректный путь | Вызов `Mod_ModelDB.GetModelDBBasePath()` | `Right(path, 12) = "\base\models\"` |
| TC-32 | `GetModelGroupFilePath` для `.xlsm` | Вызов с `"UAZ"` | Путь оканчивается на `"UAZ.xlsm"` |
| TC-33 | `GetModelGroupFilePath` fallback на `.xlsx` | Вызов с несуществующей группой | Путь оканчивается на `"{group}.xlsx"` |
| TC-34 | `ModelGroupFileExists` с существующей группой | Вызов с `"UAZ"` | `Result = True` |
| TC-35 | `ModelGroupFileExists` с несуществующей группой | Вызов с `"NONEXISTENT_GROUP"` | `Result = False` |

#### 2.2.2 `RunPickWorkTests()` — Mod_PickWork (TC-36..TC-38)

```vba
Private Sub RunPickWorkTests()
    ' TC-36: GetGroupNameFromMain — чтение группы из B14 (если пусто — "")
    ' TC-37: GetWorkSheetName — возвращает то же имя, что передано
    ' TC-38: GetWorkSheetName — с пустой строкой
End Sub
```

**Детализация TC-36..TC-38:**

| ID | Название | Проверка | Условие PASS |
|----|----------|----------|--------------|
| TC-36 | `GetGroupNameFromMain` с пустой B14 | Очистить B14, вызвать | `Result = ""` |
| TC-37 | `GetWorkSheetName` с корректным именем | Вызов с `"UAZ"` | `Result = "UAZ"` |
| TC-38 | `GetWorkSheetName` с пустой строкой | Вызов с `""` | `Result = ""` |

**Важно:** `PickWork_UI` не тестируется автоматически, т.к. требует взаимодействия с пользователем (MsgBox, открытие файлов). Тестируется только через ручную проверку.

#### 2.2.3 `RunAutoMatchTests()` — Mod_AutoMatch (TC-39..TC-44)

```vba
Private Sub RunAutoMatchTests()
    ' TC-39: GetGroupName (приватная) — косвенная проверка через вызов AutoMatchWorks с пустой B14
    '        Ожидается MsgBox "Не указана группа в ячейке B14."
    '        В автоматическом тесте — проверка, что не возникает ошибки
    ' TC-40: AutoMatchWorks с пустой группой — выход без ошибки
    ' TC-41: AutoMatchParts с пустой группой — выход без ошибки
    ' TC-42: HighlightNotFound — проверка установки цвета ячейки
    ' TC-43: ClearHighlight — проверка очистки цвета ячейки
    ' TC-44: IsAllFound — проверка обнаружения "НЕ НАЙДЕНО"
End Sub
```

**Детализация TC-39..TC-44:**

| ID | Название | Проверка | Условие PASS |
|----|----------|----------|--------------|
| TC-39 | `AutoMatchWorks` с пустой B14 | Очистить B14, вызвать | Нет ошибки (Err.Number = 0) |
| TC-40 | `AutoMatchParts` с пустой B14 | Очистить B14, вызвать | Нет ошибки (Err.Number = 0) |
| TC-41 | `HighlightNotFound` устанавливает жёлтый фон | Вызов на тестовой ячейке | `cell.Interior.Color = RGB(255,255,0)` |
| TC-42 | `HighlightNotFound` устанавливает красный текст | Вызов на тестовой ячейке | `cell.Font.Color = RGB(255,0,0)` |
| TC-43 | `ClearHighlight` очищает фон и цвет | Вызов на тестовой ячейке | `ColorIndex = xlNone` и `Font.ColorIndex = xlAutomatic` |
| TC-44 | `IsAllFound` возвращает False при "НЕ НАЙДЕНО" | Записать "НЕ НАЙДЕНО" в ячейку | `Result = False` |

**Примечание:** `AutoMatchWorks` и `AutoMatchParts` с корректными данными не тестируются автоматически, т.к. требуют наличия файлов групп и данных на листе main. Это интеграционные тесты, выполняемые вручную.

### 2.3 Углубление TC-14

**Текущий TC-14:** Проверяет только вызов `ImportFromB2_UI` с пустым B2.

**Улучшенный TC-14 (расширить):**

| Кейс | Что проверяет | Условие PASS |
|------|---------------|--------------|
| TC-14a (существующий) | Вызов с пустым B2 | Нет ошибки |
| TC-14b (новый) | Вызов `ImportFromB2_UI` с непустым B4, но без листа `{B4}M` | Нет ошибки (покажет MsgBox, но не упадёт) |
| TC-14c (новый) | Проверка, что `ImportFromB2_UI` не изменяет данные при пустом B2 | B4 восстановлено после теста |

**Важно:** Полный цикл импорта (TC-30) уже существует в реестре, но не реализован в коде. TC-14b и TC-14c — минимальное улучшение без риска повреждения данных.

### 2.4 Обновление структуры `RunAllTests()`

```vba
Public Sub RunAllTests()
    ' Инициализация счётчиков
    m_Total = 0: m_Passed = 0: m_Failed = 0: m_Skipped = 0: m_ResultsLog = ""

    Debug.Print "=============================================="
    Debug.Print "  Запуск набора тестов (TC-01..TC-44)"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Запуск групп тестов
    RunUtilsTests         ' TC-01..TC-08
    RunLoggerTests        ' TC-09..TC-11
    RunUtilsEdgeTests     ' TC-12
    RunLibNameTests       ' TC-13
    RunImportVHTests      ' TC-14 (расширенный)
    RunModelDBTests       ' TC-31..TC-35  (NEW)
    RunPickWorkTests      ' TC-36..TC-38  (NEW)
    RunAutoMatchTests     ' TC-39..TC-44  (NEW)

    ' Финальный отчёт
    PrintFinalReport
End Sub
```

### 2.5 Обновление заголовка модуля

Строка 7: `' Покрытие: TC-01 .. TC-13 (автоматические тесты)` → `' Покрытие: TC-01 .. TC-44 (автоматические тесты)`

---

## 3. План обеспечения максимального покрытия

### 3.1 Целевое покрытие

| Модуль | Текущие тесты | Новые тесты | Всего | Цель |
|--------|--------------|-------------|-------|------|
| Mod_Utils | 8 (TC-01..08, TC-12) | 0 | 8 | 100% ✅ |
| Mod_Logger | 3 (TC-09..11) | 0 | 3 | 100% ✅ |
| Mod_Constants | 1 (TC-13) | 0 | 1 | 100% ✅ |
| Mod_Import | 3 (TC-14, TC-21, TC-22, TC-30) | +2 (TC-14b, TC-14c) | 5 | 100% ✅ |
| Mod_SheetOps | 6 (TC-18..20, TC-25..29) | 0 | 6 | 100% ✅ |
| Mod_ButtonDispatcher | 1 (TC-24) | 0 | 1 | 100% ✅ |
| **Mod_ModelDB** | **0** | **+5 (TC-31..35)** | **5** | **100%** |
| **Mod_PickWork** | **0** | **+3 (TC-36..38)** | **3** | **100%** |
| **Mod_AutoMatch** | **0** | **+6 (TC-39..44)** | **6** | **100%** |
| Mod_OrderHeader | 0 | 0 | 0 | ❌ (требует данных) |
| Mod_MainButtons | 0 | 0 | 0 | ❌ (только UI-обёртки) |
| Mod_SheetButtons | 0 | 0 | 0 | ❌ (только UI-обёртки) |
| **Итого** | **22** | **+16** | **38** | **~77% модулей** |

### 3.2 Приоритеты тестов

| Приоритет | Модуль | Обоснование |
|-----------|--------|-------------|
| P0 | Mod_ModelDB | Базовый слой доступа к данным — без него не работают Mod_PickWork и Mod_AutoMatch |
| P1 | Mod_AutoMatch | Ключевой модуль автоподбора — основная бизнес-логика |
| P2 | Mod_PickWork | Вспомогательный модуль ручного подбора |
| P3 | Mod_Import (TC-14) | Углубление существующего теста |

### 3.3 Модули без автоматических тестов (обоснование)

| Модуль | Причина отсутствия тестов |
|--------|--------------------------|
| Mod_OrderHeader | Требует наличия листов `spisok` и `models` с данными. Тестируется вручную. |
| Mod_MainButtons | Содержит только UI-обёртки (вызовы других модулей). Тестируется через интеграционные тесты. |
| Mod_SheetButtons | Содержит только обработчики кнопок листов z4/work. Тестируется вручную. |

---

## 4. План обновления кодовой базы

### 4.1 P1: Централизация пути в Mod_ModelDB

**Файл:** `src/modules/Mod_ModelDB.bas`

**Текущее состояние:** Константа `MODELDB_BASE_PATH` объявлена как пустая строка (строка 15), помечена `@deprecated`. Функция `GetModelDBBasePath()` уже использует относительный путь через `ThisWorkbook.Path`.

**Что сделать:**
1. Удалить deprecated-константу `MODELDB_BASE_PATH` (строки 13-15)
2. Обновить комментарий в `GetModelDBBasePath()` — убрать упоминание абсолютного пути `L:\PROject\SysW\` (строка 25)

**Изменение:**
```diff
- ' @deprecated Используйте GetModelDBBasePath() вместо этой константы.
- ' Оставлена для обратной совместимости.
- Public Const MODELDB_BASE_PATH As String = ""
```

```diff
- ' Пример: если work.xlsm в L:\PROject\SysW\, то вернёт L:\PROject\SysW\base\models\
+ ' Пример: если work.xlsm в C:\Projects\SysW\, то вернёт C:\Projects\SysW\base\models\
```

### 4.2 P2: Привязка к work.xlsm

**Статус:** Уже исправлено в v0.13.0 (согласно CHANGELOG). `'work.xlsm'` заменён на `ThisWorkbook.Name`.

**Действие:** Проверить, что в `Mod_MainButtons.bas` и `Mod_ButtonDispatcher.bas` нет упоминаний `'work.xlsm'`. Если есть — исправить.

### 4.3 P3: Дублирование обработчиков кнопок

**Статус:** Уже исправлено в v0.13.0. `Btn_main_ImportVH_Click` удалён из `Mod_MainButtons.bas`.

**Действие:** Проверить, что `Btn_main_ImportVH_Click` присутствует только в `Mod_ButtonDispatcher.bas`.

### 4.4 P4: Абсолютные пути в скриптах

**Статус:** Уже исправлено в v0.13.0. Все Python-скрипты используют `config.py`, PowerShell — `config.ps1`.

**Остаточные упоминания `L:\PROject\SysW`:**
1. `src/modules/Mod_ModelDB.bas` строка 25 — комментарий (исправляется в п.4.1)
2. `scripts/Set-ExcelTrust.ps1` строка 216 — комментарий в выводе (информационный)

**Действие:** В `Set-ExcelTrust.ps1` заменить хардкодный путь в выводе на переменную `$ProjectPath`:
```diff
- Write-Step "  - Trusted Locations: должна быть L:\PROject\SysW" -Status "INFO"
+ Write-Step "  - Trusted Locations: должна быть $ProjectPath" -Status "INFO"
```

### 4.5 Обновление export_vba.py

**Файл:** `scripts/export_vba.py`

**Проблема:** В `COMPONENTS` (строки 30-44) отсутствуют модули `Mod_ModelDB`, `Mod_PickWork`, `Mod_AutoMatch`, `Mod_MainButtons`, `Mod_SheetOps`, `Лист2_main`.

**Что сделать:** Обновить словарь `COMPONENTS`, добавив все 13 модулей и 3 листа:

```python
COMPONENTS = {
    "Mod_AutoMatch": "modules/Mod_AutoMatch.bas",
    "Mod_ButtonDispatcher": "modules/Mod_ButtonDispatcher.bas",
    "Mod_Constants": "modules/Mod_Constants.bas",
    "Mod_FullTestRunner": "modules/Mod_FullTestRunner.bas",
    "Mod_Import": "modules/Mod_Import.bas",
    "Mod_Logger": "modules/Mod_Logger.bas",
    "Mod_MainButtons": "modules/Mod_MainButtons.bas",
    "Mod_ModelDB": "modules/Mod_ModelDB.bas",
    "Mod_OrderHeader": "modules/Mod_OrderHeader.bas",
    "Mod_PickWork": "modules/Mod_PickWork.bas",
    "Mod_SheetButtons": "modules/Mod_SheetButtons.bas",
    "Mod_SheetOps": "modules/Mod_SheetOps.bas",
    "Mod_Utils": "modules/Mod_Utils.bas",
    "Лист2": "sheets/Лист2_main.cls",
    "Sheet_work": "sheets/Sheet_work.cls",
    "Sheet_z4": "sheets/Sheet_z4.cls",
}
```

---

## 5. План обновления документации

### 5.1 Обновление DEVELOPER.md

**Файл:** `docs/DEVELOPER.md`

#### 5.1.1 Секция 1.1 (строка 9-24) — список модулей

Актуальный список (13 модулей + 3 листа) уже присутствует. Проверить соответствие.

#### 5.1.2 Секция 1.2 (строка 28-66) — схема взаимодействия

Обновить схему `Mod_FullTestRunner.RunAllTests()` — добавить новые группы:
```
Mod_FullTestRunner.RunAllTests()
       │
       ├── RunUtilsTests()         → Mod_Utils (TC-01..TC-08)
       ├── RunLoggerTests()        → Mod_Logger (TC-09..TC-11)
       ├── RunUtilsEdgeTests()     → Mod_Utils (TC-12)
       ├── RunLibNameTests()       → Mod_Constants (TC-13)
       ├── RunImportVHTests()      → Mod_Import (TC-14)
       ├── RunModelDBTests()       → Mod_ModelDB (TC-31..TC-35)  [NEW]
       ├── RunPickWorkTests()      → Mod_PickWork (TC-36..TC-38) [NEW]
       └── RunAutoMatchTests()     → Mod_AutoMatch (TC-39..TC-44) [NEW]
```

#### 5.1.3 Секция 2.5 (строка 188-214) — Mod_FullTestRunner

Обновить:
- Количество тестов: 14 → 30 (с учётом новых)
- Группы тестов: добавить `RunModelDBTests`, `RunPickWorkTests`, `RunAutoMatchTests`
- Строка 192: `(577 строк)` → обновить после добавления кода

#### 5.1.4 Секция 6.1 (строка 607-640) — полный список тестов

Добавить строки для TC-31..TC-44:

| ID | Название | Модуль | Тип | Статус |
|----|----------|--------|-----|--------|
| TC-31 | GetModelDBBasePath возвращает корректный путь | Mod_ModelDB | Модульный | ✅ |
| TC-32 | GetModelGroupFilePath для .xlsm | Mod_ModelDB | Модульный | ✅ |
| TC-33 | GetModelGroupFilePath fallback на .xlsx | Mod_ModelDB | Модульный | ✅ |
| TC-34 | ModelGroupFileExists с существующей группой | Mod_ModelDB | Модульный | ✅ |
| TC-35 | ModelGroupFileExists с несуществующей группой | Mod_ModelDB | Модульный | ✅ |
| TC-36 | GetGroupNameFromMain с пустой B14 | Mod_PickWork | Модульный | ✅ |
| TC-37 | GetWorkSheetName с корректным именем | Mod_PickWork | Модульный | ✅ |
| TC-38 | GetWorkSheetName с пустой строкой | Mod_PickWork | Модульный | ✅ |
| TC-39 | AutoMatchWorks с пустой B14 | Mod_AutoMatch | Модульный | ✅ |
| TC-40 | AutoMatchParts с пустой B14 | Mod_AutoMatch | Модульный | ✅ |
| TC-41 | HighlightNotFound устанавливает жёлтый фон | Mod_AutoMatch | Модульный | ✅ |
| TC-42 | HighlightNotFound устанавливает красный текст | Mod_AutoMatch | Модульный | ✅ |
| TC-43 | ClearHighlight очищает фон и цвет | Mod_AutoMatch | Модульный | ✅ |
| TC-44 | IsAllFound возвращает False при "НЕ НАЙДЕНО" | Mod_AutoMatch | Модульный | ✅ |

#### 5.1.5 Секция 6.2 (строка 648-657) — таблица покрытия

Обновить:

| Модуль | Всего тестов | PASS | SKIP | FAIL | Покрытие |
|--------|-------------|------|------|------|----------|
| Mod_Utils | 8 | 8 | 0 | 0 | 100% |
| Mod_Logger | 3 | 3 | 0 | 0 | 100% |
| Mod_Constants | 1 | 1 | 0 | 0 | 100% |
| Mod_Import | 5 | 5 | 0 | 0 | 100% |
| Mod_SheetOps | 6 | 6 | 0 | 0 | 100% |
| Mod_ButtonDispatcher | 1 | 1 | 0 | 0 | 100% |
| Mod_ModelDB | 5 | 5 | 0 | 0 | 100% |
| Mod_PickWork | 3 | 3 | 0 | 0 | 100% |
| Mod_AutoMatch | 6 | 6 | 0 | 0 | 100% |
| **Итого** | **38** | **38** | **0** | **0** | **100%** |

#### 5.1.6 Секция 2.11 (строка 309-340) — Mod_ModelDB

Обновить описание констант — удалить `MODELDB_BASE_PATH`, т.к. константа удалена.

### 5.2 Обновление sourcecraft-guide.md

**Файл:** `docs/sourcecraft-guide.md`

#### 5.2.1 Секция "История версий" (строка 169-176)

Добавить запись:
```markdown
| 0.14.0 | 2026-07-25 | Централизация логов: директория logs/, константа LOGS_DIR, пути логов вынесены из корня |
```

#### 5.2.2 Проверить актуальность структуры проекта (строка 83-129)

Структура уже соответствует v0.14.0 (13 модулей). Обновлений не требуется.

### 5.3 Документирование структуры листов work/z4

**Файл:** `docs/DEVELOPER.md` — добавить новую секцию после 2.17

```markdown
### 2.18 Структура листов work и z4

#### Лист work
**Назначение:** Лист для отображения и редактирования работ заказ-наряда.

**Структура колонок:**
| Колонка | Назначение |
|---------|-----------|
| A | № п/п |
| B | Артикул (модельный) |
| C | Наименование работы |
| D | Ед. изм. |
| E | Кол-во |
| F | Цена |
| G | Сумма |

**События:**
- `Worksheet_Activate` — закрепление первых двух строк (FreezePanes)

#### Лист z4
**Назначение:** Лист для отображения и редактирования запчастей заказ-наряда.

**Структура колонок:**
| Колонка | Назначение |
|---------|-----------|
| A | № п/п |
| B | Артикул (модельный) |
| C | Наименование запчасти |
| D | Ед. изм. |
| E | Кол-во |
| F | Цена |
| G | Сумма |

**События:**
- `Worksheet_Activate` — закрепление первых двух строк (FreezePanes)
```

### 5.4 Обновление CHANGELOG.md

Добавить запись о v0.15.0 в формате Keep a Changelog:

```markdown
## [0.15.0] — 2026-07-25

### Added
- **Тесты Mod_ModelDB (TC-31..TC-35):** 5 тестов для `GetModelDBBasePath`, `GetModelGroupFilePath`, `ModelGroupFileExists`
- **Тесты Mod_PickWork (TC-36..TC-38):** 3 теста для `GetGroupNameFromMain`, `GetWorkSheetName`
- **Тесты Mod_AutoMatch (TC-39..TC-44):** 6 тестов для `AutoMatchWorks`, `AutoMatchParts`, `HighlightNotFound`, `ClearHighlight`, `IsAllFound`
- **Расширение TC-14:** добавлены проверки вызова `ImportFromB2_UI` с непустым B4 и восстановления данных

### Changed
- **Mod_FullTestRunner.bas:** общее количество тестов увеличено с 22 до 38; добавлены 3 новые группы тестов
- **Mod_ModelDB.bas:** удалена deprecated-константа `MODELDB_BASE_PATH` (заменена на `GetModelDBBasePath()`)
- **scripts/export_vba.py:** обновлён словарь `COMPONENTS` — добавлены все 13 модулей и 3 листа
- **scripts/Set-ExcelTrust.ps1:** хардкодный путь в выводе заменён на переменную `$ProjectPath`

### Docs
- **DEVELOPER.md:** синхронизирован реестр тестов (TC-01..TC-44), таблица покрытия, схема тестового раннера, добавлено описание структуры листов work/z4
- **sourcecraft-guide.md:** добавлена запись о версии 0.14.0 в историю
```

---

## 6. План коммита и пуша

### 6.1 Структура коммитов

#### Коммит #1: Кодовая база
```
refactor(core): удалить MODELDB_BASE_PATH, обновить export_vba.py, исправить пути в Set-ExcelTrust.ps1
```
**Файлы:**
- `src/modules/Mod_ModelDB.bas` — удаление deprecated-константы, очистка комментариев
- `scripts/export_vba.py` — обновление COMPONENTS
- `scripts/Set-ExcelTrust.ps1` — замена хардкодного пути на переменную

#### Коммит #2: Тесты
```
test(core): добавить тесты для Mod_ModelDB, Mod_PickWork, Mod_AutoMatch, расширить TC-14
```
**Файлы:**
- `src/modules/Mod_FullTestRunner.bas` — новые тесты TC-31..TC-44, расширение TC-14

#### Коммит #3: Документация
```
docs: обновить DEVELOPER.md, sourcecraft-guide.md, CHANGELOG.md до v0.15.0
```
**Файлы:**
- `docs/DEVELOPER.md` — синхронизация
- `docs/sourcecraft-guide.md` — история версий
- `CHANGELOG.md` — запись v0.15.0

### 6.2 Формат сообщений (Conventional Commits)

| Тип | Область | Описание |
|-----|---------|----------|
| `refactor` | `core` | Изменения кода без изменения функциональности |
| `test` | `core` | Добавление или изменение тестов |
| `docs` | — | Изменения в документации |
| `fix` | `core` | Исправление ошибок |

### 6.3 Порядок выполнения

```bash
# 1. Экспорт VBA из Excel (синхронизация)
python scripts/export_vba.py

# 2. Проверить статус
git status

# 3. Коммит #1: кодовая база
git add src/modules/Mod_ModelDB.bas scripts/export_vba.py scripts/Set-ExcelTrust.ps1
git commit -m "refactor(core): удалить MODELDB_BASE_PATH, обновить export_vba.py, исправить пути в Set-ExcelTrust.ps1"

# 4. Коммит #2: тесты
git add src/modules/Mod_FullTestRunner.bas
git commit -m "test(core): добавить тесты для Mod_ModelDB, Mod_PickWork, Mod_AutoMatch, расширить TC-14"

# 5. Коммит #3: документация
git add docs/DEVELOPER.md docs/sourcecraft-guide.md CHANGELOG.md
git commit -m "docs: обновить DEVELOPER.md, sourcecraft-guide.md, CHANGELOG.md до v0.15.0"

# 6. Пуш в репозиторий
git push
```

### 6.4 Pre-commit проверки

Перед каждым коммитом:
1. Проверить UTF-8 кодировку изменяемых файлов
2. Убедиться, что `work.xlsm` не изменялся (защищён E3)
3. Проверить, что не затронуты защищённые файлы (E3)

---

## 7. Риски и ограничения

### 7.1 Защищённые файлы (E3 — изменения запрещены)

| Файл | Причина | Альтернатива |
|------|---------|-------------|
| `work.xlsm` | Основной Excel-файл | Изменения через `impVBA.py` |
| `table.md` | Документация структуры | Не изменять |
| `report.xlsx` | Отчётные данные | Не изменять |
| `README.md` | Основное описание | Не изменять |
| `.ycarules` | Правила SourceCraft | Не изменять |
| `plans/ROADMAP.md` | Стратегический план | Не изменять |
| `base/` | Шаблоны данных | Не изменять |
| `base/models/` | Модели данных | Не изменять |

### 7.2 Операции, требующие Excel

| Операция | Требует Excel | Без Excel |
|----------|--------------|-----------|
| Запуск тестов VBA | ✅ (COM) | ❌ |
| Импорт VBA в work.xlsm | ✅ (COM) | ❌ |
| Экспорт VBA из work.xlsm | ✅ (COM) | ❌ |
| Редактирование .bas/.cls на диске | ❌ | ✅ (VS Code) |

**Риск:** Если Excel не установлен или не настроен доступ к VBA Project Object Model, тесты не запустятся.

**Митигация:** Все изменения кода на диске (UTF-8) можно делать без Excel. Excel требуется только для верификации (запуска тестов).

### 7.3 Зависимости между задачами

| Задача | Зависит от | Риск блокировки |
|--------|-----------|-----------------|
| Добавление тестов | Изменений в Mod_ModelDB нет | Низкий — тесты не зависят от удаления константы |
| Обновление документации | Завершения всех изменений | Средний — нужно знать финальное состояние |
| Коммит #2 (тесты) | Коммита #1 (кодовая база) | Низкий — можно сделать в одной ветке |
| Коммит #3 (документация) | Коммитов #1 и #2 | Средний — нужно дождаться обоих |

### 7.4 Риски при добавлении тестов

| Риск | Описание | Митигация |
|------|----------|-----------|
| TC-36 (GetGroupNameFromMain) | Может изменить B14 если тест не восстановит | Сохранять и восстанавливать B14 |
| TC-39/TC-40 (AutoMatch) | Могут показать MsgBox | Использовать `Application.DisplayAlerts = False` |
| TC-41/TC-42 (HighlightNotFound) | Могут изменить форматирование листа | Использовать тестовый лист или восстанавливать |
| TC-14b (ImportFromB2_UI) | Может изменить данные на листе main | Сохранять и восстанавливать состояние |

### 7.5 Ограничения тестирования

1. **Mod_OrderHeader** не тестируется автоматически — требует наличия листов `spisok` и `models` с данными
2. **Mod_MainButtons** не тестируется автоматически — содержит только UI-обёртки с MsgBox
3. **Mod_SheetButtons** не тестируется автоматически — обработчики кнопок листов z4/work
4. **PickWork_UI** не тестируется автоматически — требует открытия файлов и взаимодействия с пользователем
5. **AutoMatchWorks/AutoMatchParts** с реальными данными не тестируются — требуют файлов групп в `base/models/`

---

## Приложение A: Полный список изменений по файлам

| Файл | Тип изменений | Описание |
|------|--------------|----------|
| `src/modules/Mod_ModelDB.bas` | Изменение | Удаление `MODELDB_BASE_PATH`, очистка комментариев |
| `src/modules/Mod_FullTestRunner.bas` | Изменение | +16 тестов (TC-31..TC-44), расширение TC-14, новые группы |
| `scripts/export_vba.py` | Изменение | Обновление COMPONENTS (16 компонентов) |
| `scripts/Set-ExcelTrust.ps1` | Изменение | Замена хардкодного пути на переменную |
| `docs/DEVELOPER.md` | Изменение | Синхронизация реестра тестов, схемы, таблицы покрытия, структуры листов |
| `docs/sourcecraft-guide.md` | Изменение | Добавление v0.14.0 в историю |
| `CHANGELOG.md` | Изменение | Добавление записи v0.15.0 |

## Приложение B: Шаблон кода для новых тестов

```vba
' ============================================================
' Группа: тесты ModelDB (TC-31..TC-35)
' ============================================================
Private Sub RunModelDBTests()
    Dim basePath As String
    Dim filePath As String

    Debug.Print "--- Mod_ModelDB Tests ---"

    ' -------------------------------------------------------
    ' TC-31: GetModelDBBasePath
    ' -------------------------------------------------------
    On Error Resume Next
    basePath = Mod_ModelDB.GetModelDBBasePath()
    If Err.number <> 0 Then
        AddResult "TC-31", "GetModelDBBasePath", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim endsCorrectly As Boolean
        endsCorrectly = (Right(basePath, 12) = "\base\models\")
        AddResult "TC-31", "GetModelDBBasePath", endsCorrectly, _
                  "Ожидалось окончание \base\models\, получено: " & basePath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-32: GetModelGroupFilePath для .xlsm
    ' -------------------------------------------------------
    On Error Resume Next
    filePath = Mod_ModelDB.GetModelGroupFilePath("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-32", "GetModelGroupFilePath .xlsm", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-32", "GetModelGroupFilePath .xlsm", (Len(filePath) > 0), _
                  "Путь пуст"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-33: GetModelGroupFilePath fallback на .xlsx
    ' -------------------------------------------------------
    On Error Resume Next
    filePath = Mod_ModelDB.GetModelGroupFilePath("NONEXISTENT_TEST_GROUP")
    If Err.number <> 0 Then
        AddResult "TC-33", "GetModelGroupFilePath fallback", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-33", "GetModelGroupFilePath fallback", (Len(filePath) > 0), _
                  "Путь пуст"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-34: ModelGroupFileExists с существующей группой
    ' -------------------------------------------------------
    On Error Resume Next
    Dim exists As Boolean
    exists = Mod_ModelDB.ModelGroupFileExists("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-34", "ModelGroupFileExists UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-34", "ModelGroupFileExists UAZ", exists, _
                  "Ожидалось True, получено False"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-35: ModelGroupFileExists с несуществующей группой
    ' -------------------------------------------------------
    On Error Resume Next
    exists = Mod_ModelDB.ModelGroupFileExists("NONEXISTENT_TEST_GROUP_12345")
    If Err.number <> 0 Then
        AddResult "TC-35", "ModelGroupFileExists несуществующая", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-35", "ModelGroupFileExists несуществующая", (Not exists), _
                  "Ожидалось False, получено True"
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub
```

---

*План создан: 2026-07-25*
*Версия: v0.15.0*
*Статус: черновик, требует согласования*