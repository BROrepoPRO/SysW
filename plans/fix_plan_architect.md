# План исправлений VBA-модулей проекта SysW

**Дата:** 2026-07-13
**Ветка:** dev
**Проект:** L:\PROject\SysW
**Роль:** Architect (анализ и планирование)

---

## 1. Mod_OrderHeader.bas — Критические несоответствия

### 1.1. Отсутствует проверка `OrderNum` на число

**Текущий код (строка 25):**
```vba
Set FoundCell = wsSpisok.Columns("A").Find(What:=OrderNum, LookAt:=xlWhole, LookIn:=xlValues)
```

**Эталон:** Перед поиском необходимо проверить, что `OrderNum` (значение B2) — число. Если не число — MsgBox и Exit Function.

**Исправление:** Добавить проверку `IsNumeric(OrderNum)` или `IsNumeric(wsMain.Range("B2").Value)` перед поиском.

### 1.2. Отсутствует очистка B3:B15 при ненайденном заказе с сообщением

**Текущий код (строки 54-57):**
```vba
Else
    ' Очищаем B3:B15
    wsMain.Range("B3:B15").ClearContents
End If
```

**Эталон:** При ненайденном заказе нужно не только очистить B3:B15, но и вывести MsgBox с сообщением, после чего Exit Function.

**Исправление:** Добавить MsgBox "Заказ не найден" перед очисткой и Exit Sub после неё.

### 1.3. Использование `Offset` вместо `Cells(row, col)` с константами

**Текущий код (строки 33-53):** Везде используется `FoundCell.Offset(0, N).Value` и `ModelFound.Offset(0, N).Value`.

**Эталон:** Запрещено использовать `Offset`. Только прямое обращение `Cells(row, columnNumber)` с именованными константами.

**Исправление:** Заменить все `Offset` на `wsSpisok.Cells(FoundCell.Row, SPISOK_COL_MODEL)` и т.д. Добавить константы столбцов в начале функции.

### 1.4. Отсутствуют именованные константы столбцов

**Текущий код:** Константы не объявлены.

**Эталон:** В начале функции `FillHeaderFromOrder` должны быть объявлены:
```vba
Const SPISOK_COL_MODEL As Long = 2
Const SPISOK_COL_GRZ As Long = 3
Const SPISOK_COL_VIN As Long = 4
Const SPISOK_COL_GARAGE As Long = 5
Const SPISOK_COL_YEAR As Long = 6
Const SPISOK_COL_MILEAGE As Long = 7
Const SPISOK_COL_DATE As Long = 8
Const MODEL_COL_GROUP As Long = 2
Const MODEL_COL_PRICE As Long = 3
Const MODEL_COL_WORKS_ORIG As Long = 4
Const MODEL_COL_WORKS_MOD As Long = 5
```

### 1.5. Неверный маппинг B11:B14 из листа `model`

**Текущий код (строки 46-49):**
```vba
wsMain.Range("B11").Value = ModelFound.Offset(0, 2).Value ' C: название модели
wsMain.Range("B12").Value = ModelFound.Offset(0, 1).Value ' B: кузов
wsMain.Range("B13").Value = ModelFound.Offset(0, 3).Value ' D: год выпуска
wsMain.Range("B14").Value = ModelFound.Offset(0, 4).Value ' E: VIN
```

**Эталон:**
- B11 = столбец C (Цена н/ч) → `MODEL_COL_PRICE` (3)
- B12 = столбец B (Группа) → `MODEL_COL_GROUP` (2)
- B13 = столбец D (Работы исх) → `MODEL_COL_WORKS_ORIG` (4)
- B14 = столбец E (Работы мод) → `MODEL_COL_WORKS_MOD` (5)

**Исправление:** Полностью изменить маппинг B11:B14 в соответствии с эталоном.

### 1.6. Поиск модели должен начинаться со строки 3

**Текущий код (строка 44):**
```vba
Set ModelFound = wsModel.Columns("A").Find(What:=ModelCode, LookAt:=xlWhole, LookIn:=xlValues)
```

**Эталон:** Поиск в столбце A листа `model` должен начинаться **со строки 3**, а не с 1.

**Исправление:** Изменить диапазон поиска на `wsModel.Range("A3:A" & wsModel.Cells(wsModel.Rows.Count, "A").End(xlUp).Row)`.

### 1.7. Отсутствует проверка `FoundCell` на Nothing после поиска модели

**Текущий код (строка 43-51):** Проверка `If ModelCode <> ""` есть, но если модель не найдена — ничего не происходит.

**Рекомендация:** Если модель не найдена — очистить B11:B14 или просто пропустить. Для надёжности стоит добавить обработку.

### 1.8. `FindOrder` — без изменений (работает корректно)

Функция `FindOrder` (строки 62-90) использует `Offset`, но в эталоне сказано "оставить без изменений, если работает корректно". Оставляем как есть.

---

## 2. Mod_Import.bas — Критические несоответствия

### 2.1. `ExtractNumberFromGRZ` — извлекает ВСЕ цифры, а не первую группу из 3-4 цифр

**Текущий код (строки 10-20):**
```vba
For i = 1 To Len(GRZ)
    If Mid(GRZ, i, 1) Like "[0-9]" Then
        Result = Result & Mid(GRZ, i, 1)
    End If
Next i
```

**Пример:** Из "А123ВС77" возвращает "12377", а должно вернуть "123".

**Эталон:** Извлекать **первую непрерывную последовательность из 3 или 4 цифр**. Если подходящей группы нет — вернуть пустую строку.

**Исправление:** Переписать логику: найти первую последовательность цифр, проверить что её длина 3 или 4, вернуть её.

### 2.2. `SearchSheetByGRZ` — ищет в `ThisWorkbook`, а не в `report.xlsx`

**Текущий код (строка 50):**
```vba
For Each ws In ThisWorkbook.Sheets
```

**Эталон:** Должен открывать книгу `report.xlsx` из `ThisWorkbook.Path` (только для чтения) и искать лист в ней.

**Исправление:** Добавить открытие `report.xlsx` через `Workbooks.Open`, поиск листа в открытой книге, возврат найденного листа. Книгу не закрывать.

### 2.3. `RenameSheetsByGRZ` — работает с `ThisWorkbook`, а должен с `report.xlsx`

**Текущий код (строка 28):**
```vba
For Each ws In ThisWorkbook.Sheets
```

**Эталон:** Открывает `report.xlsx` **для записи**, работает с её листами.

**Исправление:** Открывать `report.xlsx` для записи, переименовывать листы в ней, сохранять и закрывать.

### 2.4. `RenameSheetsByGRZ` — неверные условия фильтрации листов

**Текущий код (строка 30):**
```vba
If GRZ <> "" And ws.Name <> "main" And ws.Name <> "spisok" And ws.Name <> "model" Then
```

**Эталон:** Исключать листы `"report"` и `"spisok"` (не `"main"` и `"model"`).

**Исправление:** Изменить условие на `ws.Name <> "report" And ws.Name <> "spisok"`.

### 2.5. `RenameSheetsByGRZ` — неверный источник GRZ

**Текущий код (строка 29):**
```vba
GRZ = Trim(ws.Range("B2").Value)
```

**Эталон:** На каждом видимом листе найти **первую сверху** ячейку, содержащую слово `"автомобиль"` (без учёта регистра), и из неё извлечь GRZ.

**Исправление:** Искать ячейку с "автомобиль", извлекать GRZ через `ExtractNumberFromGRZ`.

### 2.6. `RenameSheetsByGRZ` — неверный формат имени листа

**Текущий код (строка 33):**
```vba
ws.Name = "GRZ_" & SheetNum
```

**Эталон:** Переименовывать лист в строку, возвращённую `ExtractNumberFromGRZ` (т.е. просто цифры, без префикса "GRZ_").

**Исправление:** `ws.Name = SheetNum` (с проверкой на существование имени).

### 2.7. `RenameSheetsByGRZ` — отсутствует удаление дублирующегося листа

**Эталон:** Если лист с таким именем уже существует, удалить его перед переименованием.

**Исправление:** Добавить проверку `SheetExists(wb, SheetNum)` и удаление существующего листа.

### 2.8. `ImportSheet` — не копирует лист в `ThisWorkbook` и не читает GRZ/OrderNum из main

**Текущий код (строки 61-71):**
```vba
Public Sub ImportSheet(ByVal GRZ As String)
    Dim ws As Worksheet
    Set ws = SearchSheetByGRZ(GRZ)
    If ws Is Nothing Then
        MsgBox "Лист с госномером " & GRZ & " не найден.", vbExclamation, "Импорт"
        Exit Sub
    End If
    ImportDataToMain ws
End Sub
```

**Эталон:**
- Должен скопировать найденный лист в `ThisWorkbook`
- Переименовать скопированный лист в `main!B2 & "M"`
- Вызвать `ImportDataToMain` для скопированного листа
- Должен сам читать GRZ из `main!B4` и OrderNum из `main!B2`, если параметры не переданы

**Исправление:**
- Изменить сигнатуру на `Public Sub ImportSheet(Optional ByVal GRZ As String = "", Optional ByVal OrderNum As String = "")`
- В начале процедуры, если GRZ пуст — читать `wsMain.Range("B4").Value`
- Если OrderNum пуст — читать `wsMain.Range("B2").Value`
- Добавить копирование листа: `ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)`
- Переименовать копию в `OrderNum & "M"`
- Вызвать `ImportDataToMain` для копии

### 2.9. `ImportDataToMain` — отсутствует очистка целевых диапазонов

**Текущий код (строки 94-128):** Нет очистки перед записью.

**Эталон:** Очистить на листе `main` диапазоны L:N и X:AA (со 2 строки до последней использованной).

**Исправление:** Добавить `wsMain.Range("L2:N" & LastRowMain).ClearContents` и `wsMain.Range("X2:AA" & LastRowMain).ClearContents`.

### 2.10. `ImportDataToMain` — нет поиска таблиц "Выполненные работы" и "Расходная накладная"

**Текущий код (строки 94-128):** Проходит по всем строкам источника (по колонке B) и маппит все подряд.

**Эталон:** Должен найти на `wsSource` две таблицы:
1. **"Выполненные работы"** — ориентир строка с "Выполненные работы по заказ-наряду" или первый блок с "Наименование". Данные со строки 2. Маппинг: C→L, D→M, H→N.
2. **"Расходная накладная"** — следующий блок. Маппинг: B→X, C→Y, D→Z, G→AA.

Если таблица не найдена — предупреждение, но продолжить.

**Исправление:** Полностью переработать логику: поиск маркеров таблиц, определение границ, маппинг по эталону.

### 2.11. `ImportFromReport` — УДАЛИТЬ (согласовано с пользователем)

**Решение:** Полностью удалить процедуру `ImportFromReport()` (строки 130-155) из модуля `Mod_Import.bas`.

**Обоснование:** Логика импорта через лист "report" признана ошибочной. Правильная логика реализуется через `ImportSheet`, которая работает с `report.xlsx` и копирует листы по GRZ.

### 2.12. `ImportIncomingDocNumber` — использует `SearchSheetByGRZ` (требует проверки)

**Текущий код (строки 74-91):** Использует `SearchSheetByGRZ`, который после исправления будет работать с `report.xlsx`. Нужно убедиться, что это поведение корректно для данной процедуры.

**Рекомендация:** Проверить, должна ли `ImportIncomingDocNumber` также работать с `report.xlsx`, или ей нужен поиск в `ThisWorkbook`. Если нужен поиск в `ThisWorkbook` — создать отдельную функцию.

---

## 3. Mod_ButtonDispatcher.bas — Изменения

### 3.1. `Btn_main_Import_Click` — заменить вызов на `ImportSheet`

**Текущий код (строка 11):**
```vba
Mod_Import.ImportFromReport
```

**Исправление:**
```vba
Mod_Import.ImportSheet
```
(без параметров — `ImportSheet` сама прочитает GRZ из B4 и OrderNum из B2)

### 3.2. `Btn_main_Clear_Click` — неверный диапазон очистки

**Текущий код (строка 31):**
```vba
.Range("A2:XFD" & .UsedRange.Rows.Count).ClearContents
```

**Проблема:** Используется `.UsedRange.Rows.Count`, который может не совпадать с фактической последней строкой.

**Эталон:** Очищать от A2 до последней использованной строки (по всему листу, столбцы A–XFD). Строка 1 (заголовки) не должна затираться.

**Исправление:** Использовать `ws.Cells(ws.Rows.Count, "A").End(xlUp).Row` для определения последней строки. Альтернатива: `.Range("A2:XFD" & lastRow).ClearContents`.

---

## 4. Sheet1_main.cls — Несоответствия

### 4.1. Условие исключения B3:B15 избыточно, но не вредно

**Текущий код (строка 37):**
```vba
If Not Intersect(Target, Me.Range("B3:B15")) Is Nothing Then Exit Sub
```

**Вердикт:** Условие можно оставить как дополнительную защиту от рекурсии. Не критично.

### 4.2. Нет проверки на пустое значение B2

**Текущий код (строка 49):**
```vba
Mod_OrderHeader.FillHeaderFromOrder Trim(Target.Value), Me, wsSpisok, wsModel
```

**Рекомендация:** Добавить проверку, что `Trim(Target.Value)` не пустая строка, перед вызовом `FillHeaderFromOrder`. Если B2 очистили — не вызывать заполнение.

---

## 5. .ycarules — Дополнения

### 5.1. Отсутствует правило о запрете самовольных правок

**Исправление:** Добавить правило:
```
## Запрет самовольных правок VBA-модулей
Запрещается вносить изменения в VBA-модули (.bas, .cls, .frm) без явного согласования с пользователем. Любые правки должны быть предварительно описаны в плане и утверждены.
```

---

## 6. Сводная таблица изменений

| Файл | Изменение | Серьёзность | Строки |
|------|-----------|-------------|--------|
| `Mod_OrderHeader.bas` | Добавить проверку `IsNumeric(OrderNum)` | Критическая | После строки 22 |
| `Mod_OrderHeader.bas` | Добавить MsgBox при ненайденном заказе | Критическая | Строки 54-57 |
| `Mod_OrderHeader.bas` | Заменить `Offset` на `Cells(row, col)` с константами | Критическая | Строки 33-53 |
| `Mod_OrderHeader.bas` | Добавить именованные константы столбцов | Критическая | Начало функции |
| `Mod_OrderHeader.bas` | Исправить маппинг B11:B14 (model) | Критическая | Строки 46-49 |
| `Mod_OrderHeader.bas` | Поиск модели начинать со строки 3 | Средняя | Строка 44 |
| `Mod_Import.bas` | Переписать `ExtractNumberFromGRZ` (первая группа 3-4 цифр) | Критическая | Строки 10-20 |
| `Mod_Import.bas` | `SearchSheetByGRZ` — открывать `report.xlsx` | Критическая | Строки 40-58 |
| `Mod_Import.bas` | `RenameSheetsByGRZ` — работать с `report.xlsx` | Критическая | Строки 23-37 |
| `Mod_Import.bas` | `RenameSheetsByGRZ` — исправить фильтр листов | Средняя | Строка 30 |
| `Mod_Import.bas` | `RenameSheetsByGRZ` — искать "автомобиль", не B2 | Критическая | Строка 29 |
| `Mod_Import.bas` | `RenameSheetsByGRZ` — убрать префикс "GRZ_" | Средняя | Строка 33 |
| `Mod_Import.bas` | `RenameSheetsByGRZ` — удалять дубликаты имён | Средняя | После строки 32 |
| `Mod_Import.bas` | `ImportSheet` — читать GRZ/OrderNum из main, копировать лист | Критическая | Строки 61-71 |
| `Mod_Import.bas` | `ImportDataToMain` — очистка L:N и X:AA | Критическая | Начало процедуры |
| `Mod_Import.bas` | `ImportDataToMain` — поиск таблиц по маркерам | Критическая | Строки 94-128 |
| `Mod_Import.bas` | **УДАЛИТЬ** `ImportFromReport` | Критическая | Строки 130-155 |
| `Mod_ButtonDispatcher.bas` | Заменить вызов на `Mod_Import.ImportSheet` | Критическая | Строка 11 |
| `Mod_ButtonDispatcher.bas` | Уточнить диапазон очистки в `Btn_main_Clear_Click` | Средняя | Строка 31 |
| `Sheet1_main.cls` | Добавить проверку на пустое B2 | Низкая | Перед строкой 49 |
| `.ycarules` | Добавить правило о запрете самовольных правок | Средняя | Новое правило |

---

## 7. Примечания по кодировке и именам файлов

1. **Кодировка:** Все файлы `.bas` и `.cls` должны быть в кодировке **Windows-1251**. При сохранении использовать PowerShell:
   ```powershell
   [System.IO.File]::WriteAllText("путь\файл.bas", $content, [System.Text.Encoding]::GetEncoding(1251))
   ```
2. **Имена файлов:** Строго латиница. Текущие имена соответствуют требованию.
3. **Имена модулей (Attribute VB_Name):** Должны совпадать с именами файлов (без расширения). Текущие имена корректны.

---

## 8. Рекомендации для Code-агента

### Порядок выполнения

1. **`Mod_Import.bas` — начать с фундаментальных изменений:**
   - `ExtractNumberFromGRZ` (от этой функции зависят другие)
   - `SearchSheetByGRZ` (работа с report.xlsx)
   - `RenameSheetsByGRZ` (полная переработка)
   - `ImportSheet` (чтение B2/B4, копирование листа)
   - `ImportDataToMain` (очистка + поиск таблиц по маркерам)
   - **Удалить `ImportFromReport`** (строки 130-155)

2. **`Mod_ButtonDispatcher.bas`:**
   - Заменить `Mod_Import.ImportFromReport` на `Mod_Import.ImportSheet`
   - Уточнить диапазон очистки

3. **`Mod_OrderHeader.bas`:**
   - Добавить константы столбцов
   - Заменить `Offset` на `Cells(row, col)`
   - Добавить проверку `IsNumeric`
   - Исправить маппинг B11:B14
   - Поиск model со строки 3

4. **`Sheet1_main.cls`:**
   - Добавить проверку на пустое B2

5. **`.ycarules`:**
   - Добавить правило о запрете правок

6. **Синхронизация:**
   - Записать все файлы в кодировке Windows-1251
   - Синхронизировать с книгой `work.xlsm` через задачу `Terminal → Run Task → VBA: Import from Excel`

### Важные предостережения

- **Не использовать `Offset`** в `FillHeaderFromOrder` — только `Cells(row, col)`.
- **Не путать маппинг model**: B11=Цена н/ч (столбец C), B12=Группа (столбец B), B13=Работы исх (столбец D), B14=Работы мод (столбец E).
- **`ExtractNumberFromGRZ`** должна вернуть первую группу из 3 или 4 цифр, а не все цифры подряд.
- **`SearchSheetByGRZ`** ищет в `report.xlsx`, а не в `ThisWorkbook`.
- **`ImportDataToMain`** должен искать таблицы по маркерам, а не маппить все строки подряд.
- **`ImportFromReport`** — полностью удалить, не оставлять закомментированным.
- После изменений запустить `run_tests.py` для проверки.

### Проверка после исправлений

1. Визуально проверить, что все `Offset` заменены на `Cells(row, col)`.
2. Проверить, что все константы объявлены и используются.
3. Проверить, что `ExtractNumberFromGRZ("А123ВС77")` возвращает "123", а не "12377".
4. Проверить, что `ImportFromReport` удалена.
5. Проверить, что `Btn_main_Import_Click` вызывает `Mod_Import.ImportSheet`.
6. Запустить тесты через `run_tests.py`.