# Debug-отчёт

**Дата:** 2026-07-13
**Ветка:** dev
**Задача:** Верификация VBA-модулей — синтаксис, кодировка, соответствие эталону

---

## 1. Кодировка (Windows-1251)

| Файл | Статус | Примечание |
|------|--------|------------|
| `Mod_OrderHeader.bas` | ✅ OK | Windows-1251 (no BOM), первые байты: `41 74 74 72 69 62 75 74 65 20 56 42` = "Attribute VB" |
| `Mod_Import.bas` | ✅ OK | Windows-1251 (no BOM) |
| `Mod_ButtonDispatcher.bas` | ✅ OK | Windows-1251 (no BOM) |
| `Sheet1_main.cls` | ✅ OK | Windows-1251 (no BOM), первые байты: `56 45 52 53 49 4F 4E 20 31 2E 30 20` = "VERSION 1.0 " |
| `Mod_Utils.bas` | ✅ OK | Windows-1251 (no BOM) |
| `.ycarules` | ✅ OK | Windows-1251 (no BOM) |

**Итог:** Все файлы сохранены в Windows-1251 без BOM.
**Вердикт: PASS**

---

## 2. Синтаксис VBA (баланс конструкций)

Проверялись: `Sub`/`End Sub`, `Function`/`End Function`, блочный `If`/`End If`, `For`/`Next`.

| Файл | Sub/End Sub | Func/End Func | If/End If | For/Next | Вердикт |
|------|-------------|---------------|-----------|----------|---------|
| `Mod_OrderHeader.bas` | 0/0 | 2/2 | 6/6 | 0/0 | ✅ OK |
| `Mod_Import.bas` | 4/4 | 2/2 | 22/22 | 6/6 | ✅ OK |
| `Mod_ButtonDispatcher.bas` | 2/2 | 0/0 | 2/2 | 0/0 | ✅ OK |
| `Sheet1_main.cls` | 1/1 | 0/0 | 1/1 | 0/0 | ✅ OK |
| `Mod_Utils.bas` | 1/1 | 4/4 | 0/0 | 0/0 | ✅ OK |

**Примечание:** Однострочные `If ... Then` (без `End If`) не учитывались в подсчёте блочных `If`.
**Вердикт: PASS**

---

## 3. Соответствие эталону

### 3.1. `Mod_OrderHeader.bas`

| Проверка | Статус | Детали |
|----------|--------|--------|
| `FillHeaderFromOrder` — `Function`, а не `Sub` | ✅ OK | Строка 13: `Public Function FillHeaderFromOrder(OrderNum As Variant) As Boolean` |
| Сигнатура `(OrderNum As Variant) As Boolean` | ✅ OK | Строка 13 |
| Константы `SPISOK_COL_MODEL`, `SPISOK_COL_GRZ` и т.д. | ✅ OK | Строки 22-29: `SPISOK_COL_MODEL=2`, `SPISOK_COL_GRZ=3`, `SPISOK_COL_VIN=4`, `SPISOK_COL_GARAGE=5`, `SPISOK_COL_YEAR=6`, `SPISOK_COL_MILEAGE=7`, `SPISOK_COL_DATE=8`, `SPISOK_COL_NOTE=10` |
| Константы `MODEL_COL_GROUP`, `MODEL_COL_PRICE` и т.д. | ✅ OK | Строки 32-35: `MODEL_COL_GROUP=2`, `MODEL_COL_PRICE=3`, `MODEL_COL_WORKS_ORIG=4`, `MODEL_COL_WORKS_MOD=5` |
| Нет `Offset` | ⚠️ **FAIL** | В `FillHeaderFromOrder` — ✅ нет `Offset`. Но в `FindOrder` (строки 114-120) используется `FoundCell.Offset(0, n).Value` для заполнения полей `OrderHeader`. |
| Поиск model начинается со строки 3 | ✅ OK | Строка 78: `ModelRow.Row >= 3` |
| B3 = `"00" & CStr(OrderNum) & "-20"` | ✅ OK | Строка 61: `wsMain.Cells(3, 2).Value = "00" & CStr(OrderNum) & "-20"` |

### 3.2. `Mod_Import.bas`

| Проверка | Статус | Детали |
|----------|--------|--------|
| `ExtractNumberFromGRZ` использует `For` + `Mid` + сравнение `>= "0"` (не `Like`) | ✅ OK | Строки 22-32: цикл `For i = 1 To Len(GRZ)`, `Mid(GRZ, i, 1) >= "0" And Mid(GRZ, i, 1) <= "9"` |
| `SearchSheetByGRZ` открывает книгу через `Workbooks.Open` (не проверяет `FileExists`) | ✅ OK | Строка 62: `Workbooks.Open(...)` без `FileExists` |
| `RenameSheetsByGRZ` ищет "автомобиль" через `Find` | ✅ OK | Строка 105: `ws.Cells.Find(What:="автомобиль", ...)` |
| `ImportSheet` принимает один параметр `GRZ` | ✅ OK | Строка 131: `Public Sub ImportSheet(GRZ As String)` |
| `ImportDataToMain` очищает `L2:N` и `X2:AA` | ✅ OK | Строки 175-176: `wsMain.Range("L2:N" & lastRow).ClearContents`, `wsMain.Range("X2:AA" & lastRow).ClearContents` |
| `ImportFromReport` существует и копирует A, B, C (игнорирует D) | ✅ OK | Строки 233-277: функция существует. Строки 268-274: копирует `Cells(i,1)`, `Cells(i,2)`, `Cells(i,3)` — колонки D не копируются |

### 3.3. `Mod_ButtonDispatcher.bas`

| Проверка | Статус | Детали |
|----------|--------|--------|
| `Btn_main_Import_Click` вызывает `Mod_Import.ImportFromReport` | ✅ OK | Строка 35: `Call Mod_Import.ImportFromReport` |
| `Btn_main_Clear_Click` запрашивает подтверждение, очищает `A2:XFD` до последней строки | ✅ OK | Строка 20: `MsgBox(... vbYesNo ...)`. Строка 25: `wsMain.Range("A2:XFD" & lastRow).ClearContents` |

### 3.4. `Sheet1_main.cls`

| Проверка | Статус | Детали |
|----------|--------|--------|
| `Static isProcessing As Boolean` | ✅ OK | Строка 19: `Static isProcessing As Boolean` |
| Проверка `Target.CountLarge > 1` | ✅ OK | Строка 23: `If Target.CountLarge > 1 Then Exit Sub` |
| `Intersect(Target, Me.Range("B2"))` | ✅ OK | Строка 24: `If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub` |
| Исключение B3:B15 по `Target.Row` | ✅ OK | Строка 27: `If Target.Row >= 3 And Target.Row <= 15 Then Exit Sub` |
| Вызов `Mod_OrderHeader.FillHeaderFromOrder(Me.Range("B2").Value)` | ✅ OK | Строка 31: `Call Mod_OrderHeader.FillHeaderFromOrder(Me.Range("B2").Value)` |

**Вердикт по разделу 3: FAIL** — в `Mod_OrderHeader.bas.FindOrder` (строки 114-120) используется `Offset`, что противоречит требованию "Нет Offset".

---

## 4. Имена файлов

Проверка, что все `.bas` и `.cls` файлы имеют латинские имена (не `Лист*.cls/.bas`).

| Файл | Статус |
|------|--------|
| `Mod_ButtonDispatcher.bas` | ✅ OK — латиница |
| `Mod_FullTestRunner.bas` | ✅ OK — латиница |
| `Mod_Import.bas` | ✅ OK — латиница |
| `Mod_OrderHeader.bas` | ✅ OK — латиница |
| `Mod_Utils.bas` | ✅ OK — латиница |
| `Sheet1_main.cls` | ✅ OK — латиница |

**Вердикт: PASS**

---

## 5. Мусор

Проверка наличия временных/дублирующихся файлов:

| Проверка | Статус | Детали |
|----------|--------|--------|
| Файл `$null` | ⚠️ **FAIL** | Файл `$null` существует в корне проекта (`l:\PROject\SysW\$null`) |
| Файлы `write_*.ps1` | ✅ OK | Не найдены |
| Дубликаты `.cls` файлов | ✅ OK | Только один `.cls` файл: `Sheet1_main.cls` |

**Вердикт: FAIL** — обнаружен мусорный файл `$null`.

---

## Итог

| Раздел | Вердикт |
|--------|---------|
| 1. Кодировка | ✅ PASS |
| 2. Синтаксис | ✅ PASS |
| 3. Соответствие эталону | ⚠️ **FAIL** — `Mod_OrderHeader.bas.FindOrder` использует `Offset` |
| 4. Имена файлов | ✅ PASS |
| 5. Мусор | ⚠️ **FAIL** — файл `$null` |

### Общий вердикт: **FAIL**

### Найденные проблемы (без исправления):

1. **`Mod_OrderHeader.bas` (строки 114-120):** Функция `FindOrder` использует `FoundCell.Offset(0, n).Value` для доступа к соседним ячейкам. Эталон требует отсутствия `Offset`. Рекомендуется заменить на `FoundCell.Cells(1, n+1).Value` или `ws.Cells(FoundCell.Row, FoundCell.Column + n).Value`.

2. **`$null`:** В корне проекта обнаружен мусорный файл `$null`. Рекомендуется удалить.