# План выноса UI-логики из Mod_ButtonDispatcher

## 1. Анализ текущего состояния

### 1.1. Обработчики в Mod_ButtonDispatcher (13 шт.)

| # | Обработчик | Тип логики | Содержит UI-логику? |
|---|-----------|-----------|-------------------|
| 1 | `Btn_main_Clear_Click` | Прямая реализация | **Да** — MsgBox с подтверждением, очистка листа |
| 2 | `Btn_main_Import_Click` | Вызов + ошибки | **Нет** — только вызов `Mod_Import.ImportSheet` |
| 3 | `Btn_main_FillHeader_Click` | Вызов + проверки | **Да** — проверка `IsEmpty`, MsgBox об успехе/неудаче |
| 4 | `Btn_main_ClearHeader_Click` | Прямая реализация | **Да** — очистка диапазона, MsgBox об успехе |
| 5 | `Btn_main_ImportByInput_Click` | InputBox + вызов | **Да** — InputBox, MsgBox об успехе |
| 6 | `Btn_main_RunTests_Click` | Вызов + ошибки | **Нет** — только вызов `Mod_FullTestRunner.RunAllTests` |
| 7 | `Btn_main_WriteLog_Click` | InputBox + вызов | **Да** — InputBox, MsgBox об успехе |
| 8 | `Btn_main_RenameSheets_Click` | Вызов + MsgBox | **Да** — MsgBox об успехе |
| 9 | `Btn_main_ImportDataToMain_Click` | Проверки + вызов | **Да** — проверка `wsSource Is Nothing`, MsgBox |
| 10 | `Btn_main_FindOrder_Click` | InputBox + вызов + MsgBox | **Да** — InputBox, построение сообщения, MsgBox |
| 11 | `Btn_main_ShowWorkbookPath_Click` | Вызов + MsgBox | **Да** — MsgBox с результатом |
| 12 | `Btn_main_ShowCurrentUser_Click` | Вызов + MsgBox | **Да** — MsgBox с результатом |
| 13 | `Btn_main_CheckFileExists_Click` | InputBox + вызов + MsgBox | **Да** — InputBox, MsgBox с результатом |

### 1.2. Текущие проблемы

1. **Дублирование обработки ошибок** — шаблон `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog` повторяется в 12 из 13 обработчиков.
2. **Прямая логика в диспетчере** — `Btn_main_Clear_Click` содержит полную реализацию очистки листа.
3. **Разбросанные MsgBox** — сообщения об успехе/неудаче размазаны по диспетчеру, вместо того чтобы быть в вызываемых модулях.
4. **Проверки входных данных** — `Btn_main_FillHeader_Click` проверяет `IsEmpty(orderNum)`, `Btn_main_ImportDataToMain_Click` проверяет `wsSource Is Nothing` — это должно быть в `_UI`-обёртках.

---

## 2. Проектируемые _UI-процедуры

### 2.1. Mod_OrderHeader — новые процедуры

#### `FillHeaderFromOrder_UI()`
```vba
' Запрашивает номер заказа из B2, вызывает FillHeaderFromOrder,
' показывает результат пользователю
Public Sub FillHeaderFromOrder_UI()
```
**Логика, переносимая из диспетчера:**
- Чтение `ThisWorkbook.Sheets("main").Range("B2").Value`
- Проверка `IsEmpty(orderNum) Or orderNum = ""` с MsgBox
- Вызов `Mod_OrderHeader.FillHeaderFromOrder(orderNum)`
- MsgBox об успехе или ошибке
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `FindOrder_UI()`
```vba
' Запрашивает номер заказа через InputBox, вызывает FindOrder,
' форматирует и показывает результат
Public Sub FindOrder_UI()
```
**Логика, переносимая из диспетчера:**
- `InputBox("Введите номер заказа для поиска:")`
- Проверка на пустой ввод
- Вызов `Mod_OrderHeader.FindOrder(orderNum, Header)`
- Построение строки `msg` с полями `OrderHeader`
- MsgBox с результатом (найден/не найден)
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

---

### 2.2. Mod_Import — новые процедуры

#### `ImportSheet_UI()`
```vba
' Импортирует лист по ГРЗ из ячейки B4
Public Sub ImportSheet_UI()
```
**Логика, переносимая из диспетчера:**
- Чтение `Me.Range("B4").Value` (передаётся как параметр или через ThisWorkbook)
- Вызов `Mod_Import.ImportSheet(grz)`
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

> **Примечание:** текущий `Btn_main_Import_Click` уже почти чистый — только вызов и обработка ошибок. Можно оставить как есть, но для единообразия создать `_UI`-обёртку.

#### `ImportByInput_UI()`
```vba
' Запрашивает ГРЗ через InputBox, вызывает ImportSheet
Public Sub ImportByInput_UI()
```
**Логика, переносимая из диспетчера:**
- `InputBox("Введите ГРЗ для импорта:")`
- Проверка на пустой ввод
- Вызов `Mod_Import.ImportSheet(grz)`
- MsgBox об успехе
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `ClearMainSheet_UI()`
```vba
' Очищает все данные на листе main с подтверждением пользователя
Public Sub ClearMainSheet_UI()
```
**Логика, переносимая из диспетчера:**
- `MsgBox("Очистить все данные на листе main?", vbYesNo + vbQuestion)`
- При `vbYes`: определение `lastRow`, очистка `Range("A2:XFD" & lastRow).ClearContents`
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `ClearHeader_UI()`
```vba
' Очищает шапку заказа B3:B15 на листе main
Public Sub ClearHeader_UI()
```
**Логика, переносимая из диспетчера:**
- Очистка `wsMain.Range("B3:B15").ClearContents`
- MsgBox об успехе
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `RenameSheets_UI()`
```vba
' Переименовывает листы в report.xlsx по ГРЗ
Public Sub RenameSheets_UI()
```
**Логика, переносимая из диспетчера:**
- Вызов `Mod_Import.RenameSheetsByGRZ`
- MsgBox об успехе
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `ImportDataToMain_UI()`
```vba
' Переносит данные с активного листа в лист main
Public Sub ImportDataToMain_UI()
```
**Логика, переносимая из диспетчера:**
- `Set wsSource = ActiveSheet`
- Проверка `wsSource Is Nothing` с MsgBox
- Проверка `wsSource.Name = "main"` с MsgBox
- Вызов `Mod_Import.ImportDataToMain(wsSource)`
- MsgBox об успехе
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

---

### 2.3. Mod_Utils — новые процедуры

#### `WriteLog_UI()`
```vba
' Запрашивает сообщение через InputBox и записывает в лог
Public Sub WriteLog_UI()
```
**Логика, переносимая из диспетчера:**
- `InputBox("Введите сообщение для записи в лог:")`
- Проверка на пустой ввод
- Вызов `Mod_Utils.WriteLog("Btn_main_WriteLog_Click: " & msg)`
- MsgBox об успехе
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `ShowWorkbookPath_UI()`
```vba
' Показывает путь к текущей книге через MsgBox
Public Sub ShowWorkbookPath_UI()
```
**Логика, переносимая из диспетчера:**
- Вызов `Mod_Utils.GetWorkbookPath()`
- MsgBox с результатом
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `ShowCurrentUser_UI()`
```vba
' Показывает имя текущего пользователя Windows через MsgBox
Public Sub ShowCurrentUser_UI()
```
**Логика, переносимая из диспетчера:**
- Вызов `Mod_Utils.GetCurrentUser()`
- MsgBox с результатом
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

#### `CheckFileExists_UI()`
```vba
' Запрашивает путь к файлу через InputBox и проверяет его существование
Public Sub CheckFileExists_UI()
```
**Логика, переносимая из диспетчера:**
- `InputBox("Введите полный путь к файлу:")`
- Проверка на пустой ввод
- Вызов `Mod_Utils.FileExists(filePath)`
- MsgBox с результатом (существует/не найден)
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

---

### 2.4. Mod_FullTestRunner — новая процедура

#### `RunAllTests_UI()`
```vba
' Запускает все тесты (TC-01..TC-20)
Public Sub RunAllTests_UI()
```
**Логика, переносимая из диспетчера:**
- Вызов `Mod_FullTestRunner.RunAllTests`
- `On Error GoTo ErrHandler` + `MsgBox` + `WriteLog`

> **Примечание:** `RunAllTests` уже сам показывает `MsgBox` с результатом в `PrintFinalReport`. Обработка ошибок в диспетчере — единственное, что нужно перенести.

---

## 3. Итоговый Mod_ButtonDispatcher.bas

После рефакторинга каждый обработчик будет состоять из одной строки вызова `_UI`-процедуры. Обработка ошибок полностью исключена из диспетчера.

```vba
Attribute VB_Name = "Mod_ButtonDispatcher"
Option Explicit

' ============================================================
' Модуль: Mod_ButtonDispatcher
' Назначение: Диспетчер кнопок на формах
' Содержит ТОЛЬКО однострочные вызовы _UI-процедур
' ============================================================

' ============================================================
' ОБРАБОТЧИКИ ЛИСТА MAIN
' ============================================================

Public Sub Btn_main_Clear_Click()
    Call Mod_Import.ClearMainSheet_UI
End Sub

Public Sub Btn_main_Import_Click()
    Call Mod_Import.ImportSheet_UI
End Sub

Public Sub Btn_main_FillHeader_Click()
    Call Mod_OrderHeader.FillHeaderFromOrder_UI
End Sub

Public Sub Btn_main_ClearHeader_Click()
    Call Mod_Import.ClearHeader_UI
End Sub

Public Sub Btn_main_ImportByInput_Click()
    Call Mod_Import.ImportByInput_UI
End Sub

Public Sub Btn_main_RunTests_Click()
    Call Mod_FullTestRunner.RunAllTests_UI
End Sub

Public Sub Btn_main_WriteLog_Click()
    Call Mod_Utils.WriteLog_UI
End Sub

Public Sub Btn_main_RenameSheets_Click()
    Call Mod_Import.RenameSheets_UI
End Sub

Public Sub Btn_main_ImportDataToMain_Click()
    Call Mod_Import.ImportDataToMain_UI
End Sub

Public Sub Btn_main_FindOrder_Click()
    Call Mod_OrderHeader.FindOrder_UI
End Sub

' ============================================================
' ОБРАБОТЧИКИ АДМИНИСТРИРОВАНИЯ
' ============================================================

Public Sub Btn_main_ShowWorkbookPath_Click()
    Call Mod_Utils.ShowWorkbookPath_UI
End Sub

Public Sub Btn_main_ShowCurrentUser_Click()
    Call Mod_Utils.ShowCurrentUser_UI
End Sub

Public Sub Btn_main_CheckFileExists_Click()
    Call Mod_Utils.CheckFileExists_UI
End Sub
```

---

## 4. Сводная таблица новых _UI-процедур

| Модуль | Процедура | Сигнатура | Откуда берёт логику |
|--------|----------|-----------|-------------------|
| `Mod_Import` | `ClearMainSheet_UI` | `Public Sub ClearMainSheet_UI()` | `Btn_main_Clear_Click` (полностью) |
| `Mod_Import` | `ImportSheet_UI` | `Public Sub ImportSheet_UI()` | `Btn_main_Import_Click` |
| `Mod_Import` | `ImportByInput_UI` | `Public Sub ImportByInput_UI()` | `Btn_main_ImportByInput_Click` |
| `Mod_Import` | `ClearHeader_UI` | `Public Sub ClearHeader_UI()` | `Btn_main_ClearHeader_Click` |
| `Mod_Import` | `RenameSheets_UI` | `Public Sub RenameSheets_UI()` | `Btn_main_RenameSheets_Click` |
| `Mod_Import` | `ImportDataToMain_UI` | `Public Sub ImportDataToMain_UI()` | `Btn_main_ImportDataToMain_Click` |
| `Mod_OrderHeader` | `FillHeaderFromOrder_UI` | `Public Sub FillHeaderFromOrder_UI()` | `Btn_main_FillHeader_Click` |
| `Mod_OrderHeader` | `FindOrder_UI` | `Public Sub FindOrder_UI()` | `Btn_main_FindOrder_Click` |
| `Mod_Utils` | `WriteLog_UI` | `Public Sub WriteLog_UI()` | `Btn_main_WriteLog_Click` |
| `Mod_Utils` | `ShowWorkbookPath_UI` | `Public Sub ShowWorkbookPath_UI()` | `Btn_main_ShowWorkbookPath_Click` |
| `Mod_Utils` | `ShowCurrentUser_UI` | `Public Sub ShowCurrentUser_UI()` | `Btn_main_ShowCurrentUser_Click` |
| `Mod_Utils` | `CheckFileExists_UI` | `Public Sub CheckFileExists_UI()` | `Btn_main_CheckFileExists_Click` |
| `Mod_FullTestRunner` | `RunAllTests_UI` | `Public Sub RunAllTests_UI()` | `Btn_main_RunTests_Click` |

---

## 5. Шаблон обработки ошибок (единый для всех _UI-процедур)

Каждая `_UI`-процедура должна следовать этому шаблону:

```vba
Public Sub Example_UI()
    On Error GoTo ErrHandler

    ' ... основная логика ...

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в Example_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Example_UI: " & Err.Description)
End Sub
```

---

## 6. Порядок реализации

1. **Mod_Import.bas** — добавить 6 процедур: `ClearMainSheet_UI`, `ImportSheet_UI`, `ImportByInput_UI`, `ClearHeader_UI`, `RenameSheets_UI`, `ImportDataToMain_UI`
2. **Mod_OrderHeader.bas** — добавить 2 процедуры: `FillHeaderFromOrder_UI`, `FindOrder_UI`
3. **Mod_Utils.bas** — добавить 4 процедуры: `WriteLog_UI`, `ShowWorkbookPath_UI`, `ShowCurrentUser_UI`, `CheckFileExists_UI`
4. **Mod_FullTestRunner.bas** — добавить 1 процедуру: `RunAllTests_UI`
5. **Mod_ButtonDispatcher.bas** — заменить все обработчики на однострочные вызовы

---

## 7. Диаграмма потока вызовов (до/после)

### До рефакторинга:
```
Btn_main_Clear_Click
  ├── MsgBox (подтверждение)
  ├── Прямая очистка листа
  └── Нет обработки ошибок

Btn_main_FillHeader_Click
  ├── On Error GoTo ErrHandler
  ├── Проверка IsEmpty(orderNum)
  ├── MsgBox (предупреждение)
  ├── Call Mod_OrderHeader.FillHeaderFromOrder
  ├── MsgBox (успех/неудача)
  └── ErrHandler: MsgBox + WriteLog
```

### После рефакторинга:
```
Btn_main_Clear_Click
  └── Call Mod_Import.ClearMainSheet_UI
        ├── MsgBox (подтверждение)
        ├── Очистка листа
        └── On Error GoTo ErrHandler → MsgBox + WriteLog

Btn_main_FillHeader_Click
  └── Call Mod_OrderHeader.FillHeaderFromOrder_UI
        ├── On Error GoTo ErrHandler
        ├── Проверка IsEmpty(orderNum)
        ├── MsgBox (предупреждение)
        ├── Call FillHeaderFromOrder
        ├── MsgBox (успех/неудача)
        └── ErrHandler: MsgBox + WriteLog
```

---

## 8. Расширение модульности

### 8.1. Принцип: диспетчер ничего не знает о реализации

`Mod_ButtonDispatcher` не должен содержать:
- Обработку ошибок
- MsgBox / InputBox
- Проверки данных
- Прямые манипуляции с листами

Всё это — ответственность `_UI`-процедур в соответствующих модулях.

### 8.2. Добавление нового обработчика

Чтобы добавить новую кнопку, нужно:

1. Создать бизнес-функцию в соответствующем модуле (например, `Mod_SomeFeature.DoSomething`)
2. Создать `_UI`-обёртку в том же модуле (`Mod_SomeFeature.DoSomething_UI`) с:
   - `On Error GoTo ErrHandler`
   - Ввод/вывод для пользователя (InputBox/MsgBox)
   - Вызов бизнес-функции
3. В `Mod_ButtonDispatcher` добавить одну строку:
   ```vba
   Public Sub Btn_main_NewFeature_Click()
       Call Mod_SomeFeature.DoSomething_UI
   End Sub
   ```

### 8.3. Правила именования

| Суффикс | Назначение | Где находится |
|---------|-----------|---------------|
| (без суффикса) | Бизнес-логика, чистая работа с данными | Модуль (например, `Mod_Import`) |
| `_UI` | Обёртка с пользовательским вводом/выводом и обработкой ошибок | Тот же модуль |
| `_Test` | Тестовые процедуры | `Mod_FullTestRunner` |

### 8.4. Модульная структура (диаграмма зависимостей)

```
Mod_ButtonDispatcher (только вызовы _UI)
    │
    ├──► Mod_Import._UI
    │       └──► Mod_Import (бизнес-логика)
    │
    ├──► Mod_OrderHeader._UI
    │       └──► Mod_OrderHeader (бизнес-логика)
    │
    ├──► Mod_Utils._UI
    │       └──► Mod_Utils (бизнес-логика)
    │
    └──► Mod_FullTestRunner._UI
            └──► Mod_FullTestRunner (бизнес-логика)
```

### 8.5. Расширение на новые модули

Если в будущем появится новый функциональный модуль (например, `Mod_Reporting` для отчётов):

1. Создать `Mod_Reporting.bas` с бизнес-функциями
2. Создать `Mod_Reporting._UI`-обёртки
3. В `Mod_ButtonDispatcher` добавить вызовы

Никакие существующие модули не нужно модифицировать — принцип Open/Closed.

---

## 9. Риски и замечания

1. **Совместимость с тестами** — `Mod_FullTestRunner` вызывает `FillHeaderFromOrder` напрямую (с дополнительными параметрами `wsMain`, `wsSpisok`, `wsModel`). Текущая `FillHeaderFromOrder` в `Mod_OrderHeader` принимает только `OrderNum`. Необходимо убедиться, что сигнатура не меняется.
2. **`Btn_main_Import_Click`** — использует `Me.Range("B4").Value`. В `_UI`-процедуре нужно обращаться к `ThisWorkbook.Sheets("main").Range("B4").Value`, так как `Me` в контексте модуля недоступен.
3. **`Btn_main_ImportDataToMain_Click`** — использует `ActiveSheet`. Это нормально для `_UI`-процедуры в `Mod_Import`.
4. **Порядок выполнения** — рекомендуется реализовывать модули снизу вверх: сначала `Mod_Utils`, затем `Mod_OrderHeader` и `Mod_Import`, потом `Mod_FullTestRunner`, и в последнюю очередь — `Mod_ButtonDispatcher`.