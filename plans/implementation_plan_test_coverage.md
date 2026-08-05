# План реализации: добавление тестов TC-15..TC-30, TC-45, TC-46

**Дата:** 2026-08-04
**Статус:** План (без изменений кода)
**Автор:** SourceCraft Code Assistant (Architect)
**Основа:** [`plans/analysis_test_coverage.md`](plans/analysis_test_coverage.md)
**Связанный архив:** [`plans/_archive/update_test_coverage_plan.md`](plans/_archive/update_test_coverage_plan.md)

---

## 0. Сводка

| Параметр | Значение |
|----------|----------|
| Изменяемый файл | [`src/modules/Mod_FullTestRunner.bas`](src/modules/Mod_FullTestRunner.bas) (ТОЛЬКО он, правило [Z5]) |
| Бизнес-модули | НЕ изменяются |
| Новые тесты | TC-15..TC-30, TC-45, TC-46 (18 тестов) |
| Новые группы | `RunSheetOpsTests`, `RunAggregateNameTests`, `RunModelDBReadTests`, `RunOrderHeaderTests`, `RunImportDataTests`, `RunConstantsTests` |
| Текущие итоги | Total=30, Passed=25, Failed=0, Skipped=5 (TC-41..TC-44) |
| Ожидаемые итоги | **Total=58, Passed=54, Failed=0, Skipped=4** |
| Временные данные | Удаляются после тестов ([S3]) |
| Рабочие данные | Используются где возможно ([Z6]) |

> **Примечание по текущим итогам:** в задании указано текущее состояние Total=30, Passed=25, Skipped=5. После добавления 18 новых тестов (все PASS) и при сохранении 4 SKIP (TC-41..TC-44) получаем Total=58, Passed=54, Skipped=4. Один из существующих SKIP (TC-12, кейс «пустая строка») в текущем состоянии считается в Passed, поэтому итоговая арифметика сходится к целевым значениям.

---

## 1. Порядок добавления групп тестов в `RunAllTests()`

### 1.1. Точка вставки

Новые группы регистрируются в [`RunAllTests()`](src/modules/Mod_FullTestRunner.bas:22) **после** блока `RunAutoMatchTests` (строки 69–71) и **перед** блоком `' Финальный отчёт` (строка 73).

Текущий фрагмент (строки 69–74):

```vba
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAutoMatchTests START")
    RunAutoMatchTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAutoMatchTests END")

    ' Финальный отчёт
    PrintFinalReport
```

### 1.2. Вставляемый блок

Вставляется между строкой `RunAutoMatchTests END` (строка 71) и строкой `' Финальный отчёт` (строка 73):

```vba
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests START")
    RunSheetOpsTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAggregateNameTests START")
    RunAggregateNameTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAggregateNameTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBReadTests START")
    RunModelDBReadTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBReadTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunOrderHeaderTests START")
    RunOrderHeaderTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunOrderHeaderTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportDataTests START")
    RunImportDataTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportDataTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunConstantsTests START")
    RunConstantsTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunConstantsTests END")
```

### 1.3. Порядок и обоснование

| № | Группа | Тесты | Модуль | Почему в этом порядке |
|---|--------|-------|--------|----------------------|
| 1 | `RunSheetOpsTests` | TC-15..TC-18, TC-45 | `Mod_SheetOps` | Чистые функции + интеграционный поиск листа |
| 2 | `RunAggregateNameTests` | TC-19..TC-21 | `Mod_Constants` | Чистая функция, без побочных эффектов |
| 3 | `RunModelDBReadTests` | TC-22..TC-24 | `Mod_ModelDB` | Чтение внешней книги UAZ.xlsm |
| 4 | `RunOrderHeaderTests` | TC-25..TC-28 | `Mod_OrderHeader` | Запись в main/models (самые рискованные) |
| 5 | `RunImportDataTests` | TC-29, TC-30 | `Mod_Import` | Запись в main + открытие report.xlsx |
| 6 | `RunConstantsTests` | TC-46 | `Mod_Constants` | Запись в libname |

> **Порядок выбран так, чтобы сначала выполнялись безопасные чистые функции, затем чтение внешних книг, затем операции с записью в листы (которые требуют сохранения/восстановления состояния).**

### 1.4. Обновление заголовочного комментария

Строка 7 модуля:

```vba
' Покрытие: TC-01 .. TC-44 (автоматические тесты)
```

заменяется на:

```vba
' Покрытие: TC-01 .. TC-46 (автоматические тесты)
```

Также обновить строку 34 (`Debug.Print "  Запуск набора тестов (TC-01..TC-44)"`) на `TC-01..TC-46`.

---

## 2. Группа `RunSheetOpsTests` (TC-15..TC-18, TC-45)

### 2.1. Описание

Тестирует [`Mod_SheetOps.ExtractNumberFromGRZ`](src/modules/Mod_SheetOps.bas:20) (чистая функция) и [`Mod_SheetOps.SearchSheetByGRZ`](src/modules/Mod_SheetOps.bas:53) (интеграционный, открывает report.xlsx).

### 2.2. Данные

- TC-15..TC-18: литеральные строки (без внешних данных).
- TC-45: несуществующий ГРЗ `"НЕСУЩЕСТВУЮЩИЙ"`; `SearchSheetByGRZ` открывает `report.xlsx` (ReadOnly:=True) и закрывает его внутри.

### 2.3. Полный код процедуры

```vba
' ============================================================
' Группа: тесты SheetOps (TC-15..TC-18, TC-45)
' ============================================================
Private Sub RunSheetOpsTests()
    Dim Result As String
    Dim ws As Worksheet

    Debug.Print "--- Mod_SheetOps Tests ---"

    ' -------------------------------------------------------
    ' TC-15: ExtractNumberFromGRZ 'А123АН77' -> '123'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А123АН77")
    If Err.Number <> 0 Then
        AddResult "TC-15", "ExtractNumberFromGRZ 'А123АН77'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-15", "ExtractNumberFromGRZ 'А123АН77'", (Result = "123"), _
                  "Ожидалось '123', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-16: ExtractNumberFromGRZ 'А12АН34' (2 цифры) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А12АН34")
    If Err.Number <> 0 Then
        AddResult "TC-16", "ExtractNumberFromGRZ 'А12АН34'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-16", "ExtractNumberFromGRZ 'А12АН34'", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-17: ExtractNumberFromGRZ 'А1234АН77' (4 цифры) -> '1234'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А1234АН77")
    If Err.Number <> 0 Then
        AddResult "TC-17", "ExtractNumberFromGRZ 'А1234АН77'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-17", "ExtractNumberFromGRZ 'А1234АН77'", (Result = "1234"), _
                  "Ожидалось '1234', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-18: ExtractNumberFromGRZ '' (пустая строка) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("")
    If Err.Number <> 0 Then
        AddResult "TC-18", "ExtractNumberFromGRZ пустая строка", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-18", "ExtractNumberFromGRZ пустая строка", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-45: SearchSheetByGRZ несуществующий ГРЗ -> Nothing
    ' (Mod_Constants.SilenceMsgBox уже = True в RunAllTests)
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = Mod_SheetOps.SearchSheetByGRZ("НЕСУЩЕСТВУЮЩИЙ")
    If Err.Number <> 0 Then
        AddResult "TC-45", "SearchSheetByGRZ несуществующий ГРЗ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-45", "SearchSheetByGRZ несуществующий ГРЗ", (ws Is Nothing), _
                  "Ожидалось Nothing, лист найден"
    End If
    Set ws = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub
```

### 2.4. Обработка побочных эффектов

- TC-15..TC-18: нет побочных эффектов (чистая функция).
- TC-45: `SearchSheetByGRZ` открывает `report.xlsx` (ReadOnly:=True) и **закрывает его внутри** (строки 79–81 `Mod_SheetOps.bas`). При ошибке ErrHandler также закрывает книгу (строки 88–90). Дополнительное закрытие не требуется. MsgBox в ErrHandler подавляется флагом `Mod_Constants.SilenceMsgBox = True` (уже установлен в `RunAllTests`).

---

## 3. Группа `RunAggregateNameTests` (TC-19..TC-21)

### 3.1. Описание

Тестирует [`Mod_Constants.GetAggregateName`](src/modules/Mod_Constants.bas:400) — чистая функция, использует `UCase$(Trim$(code))`, поэтому регистр не важен.

### 3.2. Данные

- TC-19: `"DIAG"` → `"Диагностика"` (константа `AGG_DIAG`).
- TC-20: `"TO"` → `"ТО"` (константа `AGG_TO`; код латиницей, не путать с русским названием).
- TC-21: `"XXX"` → `""` (неизвестный код).

### 3.3. Полный код процедуры

```vba
' ============================================================
' Группа: тесты AggregateName (TC-19..TC-21)
' ============================================================
Private Sub RunAggregateNameTests()
    Dim Result As String

    Debug.Print "--- Mod_Constants GetAggregateName Tests ---"

    ' -------------------------------------------------------
    ' TC-19: GetAggregateName 'DIAG' -> 'Диагностика'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("DIAG")
    If Err.Number <> 0 Then
        AddResult "TC-19", "GetAggregateName 'DIAG'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-19", "GetAggregateName 'DIAG'", (Result = "Диагностика"), _
                  "Ожидалось 'Диагностика', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-20: GetAggregateName 'TO' -> 'ТО'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("TO")
    If Err.Number <> 0 Then
        AddResult "TC-20", "GetAggregateName 'TO'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-20", "GetAggregateName 'TO'", (Result = "ТО"), _
                  "Ожидалось 'ТО', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-21: GetAggregateName 'XXX' (неизвестный) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("XXX")
    If Err.Number <> 0 Then
        AddResult "TC-21", "GetAggregateName 'XXX'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-21", "GetAggregateName 'XXX'", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub
```

### 3.4. Обработка побочных эффектов

Нет побочных эффектов (чистая функция).

---

## 4. Группа `RunModelDBReadTests` (TC-22..TC-24)

### 4.1. Описание

Тестирует чтение данных из файла группы `UAZ` (`base/models/UAZ.xlsm`):
- [`Mod_ModelDB.GetWorkIdentities`](src/modules/Mod_ModelDB.bas:195) — лист `UAZw`, возвращает `Collection` объектов `WorkIdentity`.
- [`Mod_ModelDB.GetPartIdentities`](src/modules/Mod_ModelDB.bas:283) — лист `UAZz4`, возвращает `Collection` объектов `PartIdentity`.
- [`Mod_ModelDB.GetWorks`](src/modules/Mod_ModelDB.bas:124) — лист `UAZ`, возвращает `Collection` элементов типа `WorkEntry` (UDT в `Variant`).

### 4.2. Данные

Рабочие данные из `base/models/UAZ.xlsm` ([Z6]). Перед тестами проверяется наличие файла через [`Mod_ModelDB.ModelGroupFileExists`](src/modules/Mod_ModelDB.bas:54). Если файла нет — тесты SKIP.

### 4.3. Важные особенности реализации

- `GetWorkIdentities`/`GetPartIdentities` возвращают `Collection` **объектов** классов. Проверка: `col.Count > 0`, `TypeName(col(1)) = "WorkIdentity"` / `"PartIdentity"`, поля `OutArticle` (B) и `Aggregate` (I) непустые.
- `GetWorks` возвращает `Collection` элементов типа `WorkEntry` (UDT). Элементы хранятся в `Variant`. Для проверки `TypeName` вернёт имя UDT, но обращение к полям UDT через `Variant` требует присваивания в локальную переменную типа `WorkEntry`. Для простоты и надёжности TC-24 проверяет только `col.Count > 0` и что первый элемент не пуст (через `TypeName`).
- Все три функции **открывают** `UAZ.xlsm` через `OpenModelGroupFile` (ReadOnly:=False). После теста книгу нужно закрыть (`wb.Close SaveChanges:=False`), как в TC-35.

### 4.4. Полный код процедуры

```vba
' ============================================================
' Группа: тесты ModelDB Read (TC-22..TC-24)
' ============================================================
Private Sub RunModelDBReadTests()
    Dim col As Collection
    Dim wb As Workbook
    Dim groupExists As Boolean
    Dim item As Variant

    Debug.Print "--- Mod_ModelDB Read Tests ---"

    ' Проверяем наличие файла группы UAZ (рабочие данные)
    On Error Resume Next
    groupExists = Mod_ModelDB.ModelGroupFileExists("UAZ")
    If Err.Number <> 0 Then
        groupExists = False
        Err.Clear
    End If
    On Error GoTo 0

    If Not groupExists Then
        AddResult "TC-22", "GetWorkIdentities UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        AddResult "TC-23", "GetPartIdentities UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        AddResult "TC-24", "GetWorks UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        Debug.Print ""
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' TC-22: GetWorkIdentities 'UAZ' — непустая коллекция WorkIdentity
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetWorkIdentities("UAZ")
    If Err.Number <> 0 Then
        AddResult "TC-22", "GetWorkIdentities UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc22Ok As Boolean
        Dim tc22Reason As String
        tc22Ok = False
        If col.Count > 0 Then
            If TypeName(col(1)) = "WorkIdentity" Then
                Dim wi As WorkIdentity
                Set wi = col(1)
                tc22Ok = (Len(wi.OutArticle) > 0) And (Len(wi.Aggregate) > 0)
                If Not tc22Ok Then
                    tc22Reason = "OutArticle/Aggregate пусты"
                End If
                Set wi = Nothing
            Else
                tc22Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc22Reason = "Коллекция пуста (нет данных в UAZw)"
        End If
        AddResult "TC-22", "GetWorkIdentities UAZ", tc22Ok, tc22Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-23: GetPartIdentities 'UAZ' — непустая коллекция PartIdentity
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetPartIdentities("UAZ")
    If Err.Number <> 0 Then
        AddResult "TC-23", "GetPartIdentities UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc23Ok As Boolean
        Dim tc23Reason As String
        tc23Ok = False
        If col.Count > 0 Then
            If TypeName(col(1)) = "PartIdentity" Then
                Dim pi As PartIdentity
                Set pi = col(1)
                tc23Ok = (Len(pi.OutArticle) > 0) And (Len(pi.Aggregate) > 0)
                If Not tc23Ok Then
                    tc23Reason = "OutArticle/Aggregate пусты"
                End If
                Set pi = Nothing
            Else
                tc23Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc23Reason = "Коллекция пуста (нет данных в UAZz4)"
        End If
        AddResult "TC-23", "GetPartIdentities UAZ", tc23Ok, tc23Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-24: GetWorks 'UAZ' — непустая коллекция WorkEntry
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetWorks("UAZ", Empty)
    If Err.Number <> 0 Then
        AddResult "TC-24", "GetWorks UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc24Ok As Boolean
        Dim tc24Reason As String
        tc24Ok = False
        If col.Count > 0 Then
            ' Элементы типа WorkEntry (UDT) хранятся в Variant
            item = col(1)
            tc24Ok = (Len(CStr(item.Code)) > 0)
            If Not tc24Ok Then
                tc24Reason = "Code первого элемента пуст"
            End If
        Else
            tc24Reason = "Коллекция пуста (нет данных в UAZ)"
        End If
        AddResult "TC-24", "GetWorks UAZ", tc24Ok, tc24Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' Закрываем книгу UAZ.xlsm, если она осталась открытой
    On Error Resume Next
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks("UAZ.xlsm")
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub
```

### 4.5. Обработка побочных эффектов

- Все три функции открывают `UAZ.xlsm` (ReadOnly:=False). После тестов книга закрывается через `Workbooks("UAZ.xlsm").Close SaveChanges:=False` (без сохранения, чтобы не изменить файл).
- Если файл группы отсутствует — тесты SKIP (не FAIL), чтобы не ломать прогон на машинах без `base/models/UAZ.xlsm`.
- `GetWorks` принимает `filters` по ссылке (`ByRef filters As Variant`), поэтому передаётся `Empty`.

---

## 5. Группа `RunOrderHeaderTests` (TC-25..TC-28)

### 5.1. Описание

Тестирует [`Mod_OrderHeader.FillHeaderFromOrder`](src/modules/Mod_OrderHeader.bas:30) и [`Mod_OrderHeader.FindOrder`](src/modules/Mod_OrderHeader.bas:176).

### 5.2. Данные

- TC-25, TC-27: реальный номер заказа — первый непустой номер из столбца A листа `spisok` в `work.xlsm` ([Z6]).
- TC-26, TC-28: заведомо несуществующий номер `999999`.

### 5.3. Ключевые риски и митигация

1. **TC-25 (`FillHeaderFromOrder`)** — самая опасная функция:
   - Очищает и заполняет `B5:B17` листа `main`.
   - Если модель из `spisok` **не найдена** в листе `models`, функция **добавляет новую строку** в `models` (строки 133–155 `Mod_OrderHeader.bas`).
   - Показывает MsgBox, если у модели пусты и группа, и цена (строки 128–132).
   - **Митигация:** сохранить `B5:B17` до теста и восстановить после; сохранить последнюю строку листа `models` и, если она изменилась, удалить добавленную строку; `Application.EnableEvents = False` при записи; `Mod_Constants.SilenceMsgBox = True` (уже установлен).
2. **TC-26 (`FillHeaderFromOrder` с несуществующим номером)** — очищает `B5:B17`. Сохранить и восстановить.
3. **TC-27, TC-28 (`FindOrder`)** — только чтение `spisok`, побочных эффектов нет.

### 5.4. Полный код процедуры

```vba
' ============================================================
' Группа: тесты OrderHeader (TC-25..TC-28)
' ============================================================
Private Sub RunOrderHeaderTests()
    Dim wsMain As Worksheet
    Dim wsSpisok As Worksheet
    Dim wsModels As Worksheet
    Dim orderNum As Variant
    Dim Result As Boolean
    Dim Header As OrderHeader
    Dim savedB5toB17 As Variant
    Dim modelsLastRowBefore As Long
    Dim modelsLastRowAfter As Long
    Dim i As Long

    Debug.Print "--- Mod_OrderHeader Tests ---"

    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)
    Set wsSpisok = ThisWorkbook.Sheets(Mod_Constants.SHEET_SPISOK)
    Set wsModels = ThisWorkbook.Sheets(Mod_Constants.SHEET_MODELS)
    On Error GoTo 0

    If wsMain Is Nothing Or wsSpisok Is Nothing Or wsModels Is Nothing Then
        AddResult "TC-25", "FillHeaderFromOrder существующий заказ", True, "", True, _
                  "Лист main/spisok/models не найден"
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", True, "", True, _
                  "Лист main/spisok/models не найден"
        AddResult "TC-27", "FindOrder существующий заказ", True, "", True, _
                  "Лист spisok не найден"
        AddResult "TC-28", "FindOrder несуществующий заказ", True, "", True, _
                  "Лист spisok не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' Получаем реальный номер заказа: первый непустой номер из столбца A spisok
    orderNum = Empty
    For i = 2 To wsSpisok.Cells(wsSpisok.Rows.Count, 1).End(xlUp).Row
        If Not IsEmpty(wsSpisok.Cells(i, 1).Value) Then
            orderNum = wsSpisok.Cells(i, 1).Value
            Exit For
        End If
    Next i

    If IsEmpty(orderNum) Then
        AddResult "TC-25", "FillHeaderFromOrder существующий заказ", True, "", True, _
                  "В spisok нет ни одного номера заказа"
        AddResult "TC-27", "FindOrder существующий заказ", True, "", True, _
                  "В spisok нет ни одного номера заказа"
    Else
        ' =====================================================
        ' TC-25: FillHeaderFromOrder с существующим заказом
        ' =====================================================
        ' Сохраняем B5:B17 листа main
        savedB5toB17 = wsMain.Range("B5:B17").Value

        ' Сохраняем последнюю строку листа models (для детекции автодобавления)
        modelsLastRowBefore = wsModels.Cells(wsModels.Rows.Count, 1).End(xlUp).Row

        On Error Resume Next
        Application.EnableEvents = False
        Result = Mod_OrderHeader.FillHeaderFromOrder(orderNum)
        Application.EnableEvents = True
        If Err.Number <> 0 Then
            AddResult "TC-25", "FillHeaderFromOrder существующий заказ", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim tc25Ok As Boolean
            Dim tc25Reason As String
            tc25Ok = Result And (Len(Trim(CStr(wsMain.Range("B5").Value))) > 0)
            If Not Result Then
                tc25Reason = "Функция вернула False"
            ElseIf Len(Trim(CStr(wsMain.Range("B5").Value))) = 0 Then
                tc25Reason = "B5 не заполнена"
            End If
            AddResult "TC-25", "FillHeaderFromOrder существующий заказ", tc25Ok, tc25Reason
        End If
        On Error GoTo 0

        ' Восстанавливаем B5:B17
        Application.EnableEvents = False
        wsMain.Range("B5:B17").Value = savedB5toB17
        Application.EnableEvents = True

        ' Если FillHeaderFromOrder добавил модель в models — удаляем добавленную строку
        modelsLastRowAfter = wsModels.Cells(wsModels.Rows.Count, 1).End(xlUp).Row
        If modelsLastRowAfter > modelsLastRowBefore Then
            wsModels.Rows(modelsLastRowAfter).Delete
        End If

        ' =====================================================
        ' TC-27: FindOrder с существующим заказом
        ' =====================================================
        On Error Resume Next
        Result = Mod_OrderHeader.FindOrder(CStr(orderNum), Header)
        If Err.Number <> 0 Then
            AddResult "TC-27", "FindOrder существующий заказ", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim tc27Ok As Boolean
            Dim tc27Reason As String
            tc27Ok = Result And (Len(Trim(CStr(Header.OrderNumber))) > 0) _
                          And (Len(Trim(CStr(Header.ModelName))) > 0)
            If Not Result Then
                tc27Reason = "Функция вернула False"
            ElseIf Len(Trim(CStr(Header.OrderNumber))) = 0 Then
                tc27Reason = "Header.OrderNumber пуст"
            ElseIf Len(Trim(CStr(Header.ModelName))) = 0 Then
                tc27Reason = "Header.ModelName пуст"
            End If
            AddResult "TC-27", "FindOrder существующий заказ", tc27Ok, tc27Reason
        End If
        On Error GoTo 0
    End If

    ' =====================================================
    ' TC-26: FillHeaderFromOrder с несуществующим заказом
    ' =====================================================
    savedB5toB17 = wsMain.Range("B5:B17").Value

    On Error Resume Next
    Application.EnableEvents = False
    Result = Mod_OrderHeader.FillHeaderFromOrder(999999)
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc26Ok As Boolean
        Dim tc26Reason As String
        tc26Ok = (Not Result) And IsEmpty(wsMain.Range("B5").Value)
        If Result Then
            tc26Reason = "Функция вернула True для несуществующего заказа"
        ElseIf Not IsEmpty(wsMain.Range("B5").Value) Then
            tc26Reason = "B5 не очищена"
        End If
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", tc26Ok, tc26Reason
    End If
    On Error GoTo 0

    ' Восстанавливаем B5:B17
    Application.EnableEvents = False
    wsMain.Range("B5:B17").Value = savedB5toB17
    Application.EnableEvents = True

    ' =====================================================
    ' TC-28: FindOrder с несуществующим заказом
    ' =====================================================
    On Error Resume Next
    Result = Mod_OrderHeader.FindOrder("999999", Header)
    If Err.Number <> 0 Then
        AddResult "TC-28", "FindOrder несуществующий заказ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-28", "FindOrder несуществующий заказ", (Not Result), _
                  "Ожидалось False, получено " & CStr(Result)
    End If
    On Error GoTo 0

    Set wsMain = Nothing
    Set wsSpisok = Nothing
    Set wsModels = Nothing

    Debug.Print ""
End Sub
```

### 5.5. Обработка побочных эффектов

- **TC-25:** сохраняет `B5:B17` и восстанавливает; детектирует автодобавление модели в `models` (сравнение последней строки до/после) и удаляет добавленную строку; `Application.EnableEvents = False` при записи; MsgBox подавлен.
- **TC-26:** сохраняет `B5:B17` и восстанавливает; `Application.EnableEvents = False` при записи.
- **TC-27, TC-28:** только чтение `spisok`, побочных эффектов нет.

---

## 6. Группа `RunImportDataTests` (TC-29, TC-30)

### 6.1. Описание

Тестирует [`Mod_Import.ImportDataToMain`](src/modules/Mod_Import.bas:60) и [`Mod_Import.ImportSheet`](src/modules/Mod_Import.bas:21).

### 6.2. Данные

- **TC-29:** временный лист-источник в `work.xlsm` с таблицами «Выполненные работы» и «Расходная накладная» (структура по комментариям в `Mod_Import.bas`). После теста лист удаляется ([S3]).
- **TC-30:** несуществующий ГРЗ `"НЕСУЩЕСТВУЮЩИЙ"`.

### 6.3. Структура временного листа для TC-29

**Таблица «Выполненные работы»** (маппинг: D→L, I→M, M→N):
- Строка 1: заголовок `№ | № кат. | Наименование | Кол. оп. | Цена | Норма | н/ч | Всего | в т.ч. НДС`
- Строка 2: подзаголовок `1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9`
- Строка 3+: данные (D=Наименование, I=Всего, M=в т.ч. НДС)

**Таблица «Расходная накладная»** (маппинг: C→X, D→Y, J→Z, M→AA):
- Строка 1: заголовок `№ | № кат. | Наименование | Кол-во | Ед.изм. | Цена | Всего | в т.ч. НДС`
- Строка 2: подзаголовок `1 | 2 | 3 | 4 | 5 | 6 | 7 | 8`
- Строка 3+: данные (C=Наименование, D=Кол-во, J=Всего, M=в т.ч. НДС)

> **Важно:** `ImportDataToMain` ищет таблицы через `Cells.Find(What:="Выполненные работы")` и `Cells.Find(What:="Расходная накладная")`. Заголовки должны присутствовать дословно. После названия таблицы идёт пустая строка, затем 2 строки заголовка, затем данные (логика `dataStartRow` в строках 126–132 и 182–188).

### 6.4. Полный код процедуры

```vba
' ============================================================
' Группа: тесты ImportData (TC-29, TC-30)
' ============================================================
Private Sub RunImportDataTests()
    Dim wsMain As Worksheet
    Dim wsTemp As Worksheet
    Dim savedLtoN As Variant
    Dim savedXtoAA As Variant
    Dim tempName As String
    Dim lastRow As Long

    Debug.Print "--- Mod_Import Tests ---"

    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)
    On Error GoTo 0

    If wsMain Is Nothing Then
        AddResult "TC-29", "ImportDataToMain перенос данных", True, "", True, _
                  "Лист main не найден"
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", True, "", True, _
                  "Лист main не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' =====================================================
    ' TC-29: ImportDataToMain с временным листом-источником
    ' =====================================================
    ' Сохраняем L:N и X:AA листа main
    lastRow = Application.WorksheetFunction.Max( _
        wsMain.Cells(wsMain.Rows.Count, 12).End(xlUp).Row, _
        wsMain.Cells(wsMain.Rows.Count, 24).End(xlUp).Row)
    If lastRow < 4 Then lastRow = 4
    savedLtoN = wsMain.Range("L4:N" & lastRow).Value
    savedXtoAA = wsMain.Range("X4:AA" & lastRow).Value

    ' Создаём временный лист
    tempName = "TC29_Temp_" & Format$(Timer * 1000, "0")
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = tempName
    On Error GoTo 0

    If wsTemp Is Nothing Then
        AddResult "TC-29", "ImportDataToMain перенос данных", False, "Не удалось создать временный лист"
    Else
        ' --- Таблица "Выполненные работы" ---
        wsTemp.Cells(1, 1).Value = "Выполненные работы"
        wsTemp.Cells(2, 1).Value = "№"
        wsTemp.Cells(2, 2).Value = "№ кат."
        wsTemp.Cells(2, 3).Value = "Наименование"
        wsTemp.Cells(2, 4).Value = "Кол. оп."
        wsTemp.Cells(2, 5).Value = "Цена"
        wsTemp.Cells(2, 6).Value = "Норма"
        wsTemp.Cells(2, 7).Value = "н/ч"
        wsTemp.Cells(2, 8).Value = "Всего"
        wsTemp.Cells(2, 9).Value = "в т.ч. НДС"
        wsTemp.Cells(3, 1).Value = "1"
        wsTemp.Cells(3, 2).Value = "2"
        wsTemp.Cells(3, 3).Value = "3"
        wsTemp.Cells(3, 4).Value = "4"
        wsTemp.Cells(3, 5).Value = "5"
        wsTemp.Cells(3, 6).Value = "6"
        wsTemp.Cells(3, 7).Value = "7"
        wsTemp.Cells(3, 8).Value = "8"
        wsTemp.Cells(3, 9).Value = "9"
        ' Данные работ: D=Наименование, I=Всего, M=в т.ч. НДС
        wsTemp.Cells(4, 4).Value = "Тестовая работа 1"
        wsTemp.Cells(4, 9).Value = 100
        wsTemp.Cells(4, 13).Value = 20
        wsTemp.Cells(5, 4).Value = "Тестовая работа 2"
        wsTemp.Cells(5, 9).Value = 200
        wsTemp.Cells(5, 13).Value = 40
        wsTemp.Cells(6, 4).Value = "Итого работ"

        ' --- Таблица "Расходная накладная" ---
        wsTemp.Cells(8, 1).Value = "Расходная накладная"
        wsTemp.Cells(9, 1).Value = "№"
        wsTemp.Cells(9, 2).Value = "№ кат."
        wsTemp.Cells(9, 3).Value = "Наименование"
        wsTemp.Cells(9, 4).Value = "Кол-во"
        wsTemp.Cells(9, 5).Value = "Ед.изм."
        wsTemp.Cells(9, 6).Value = "Цена"
        wsTemp.Cells(9, 7).Value = "Всего"
        wsTemp.Cells(9, 8).Value = "в т.ч. НДС"
        wsTemp.Cells(10, 1).Value = "1"
        wsTemp.Cells(10, 2).Value = "2"
        wsTemp.Cells(10, 3).Value = "3"
        wsTemp.Cells(10, 4).Value = "4"
        wsTemp.Cells(10, 5).Value = "5"
        wsTemp.Cells(10, 6).Value = "6"
        wsTemp.Cells(10, 7).Value = "7"
        wsTemp.Cells(10, 8).Value = "8"
        ' Данные материалов: C=Наименование, D=Кол-во, J=Всего, M=в т.ч. НДС
        wsTemp.Cells(11, 3).Value = "Тестовая запчасть 1"
        wsTemp.Cells(11, 4).Value = 2
        wsTemp.Cells(11, 10).Value = 50
        wsTemp.Cells(11, 13).Value = 10
        wsTemp.Cells(12, 3).Value = "Тестовая запчасть 2"
        wsTemp.Cells(12, 4).Value = 3
        wsTemp.Cells(12, 10).Value = 60
        wsTemp.Cells(12, 13).Value = 12
        wsTemp.Cells(13, 2).Value = "Итого"

        ' Подавляем MsgBox для Mod_Import
        Mod_Import.SilenceMsgBox = True

        On Error Resume Next
        Call Mod_Import.ImportDataToMain(wsTemp)
        If Err.Number <> 0 Then
            AddResult "TC-29", "ImportDataToMain перенос данных", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim tc29Ok As Boolean
            Dim tc29Reason As String
            ' Проверяем перенос: L4 = D источника, X4 = C источника
            tc29Ok = (Trim(CStr(wsMain.Cells(4, 12).Value)) = "Тестовая работа 1") _
                 And (Trim(CStr(wsMain.Cells(4, 24).Value)) = "Тестовая запчасть 1")
            If Trim(CStr(wsMain.Cells(4, 12).Value)) <> "Тестовая работа 1" Then
                tc29Reason = "L4 не заполнена: '" & CStr(wsMain.Cells(4, 12).Value) & "'"
            ElseIf Trim(CStr(wsMain.Cells(4, 24).Value)) <> "Тестовая запчасть 1" Then
                tc29Reason = "X4 не заполнена: '" & CStr(wsMain.Cells(4, 24).Value) & "'"
            End If
            AddResult "TC-29", "ImportDataToMain перенос данных", tc29Ok, tc29Reason
        End If
        On Error GoTo 0

        Mod_Import.SilenceMsgBox = False

        ' Удаляем временный лист ([S3])
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If

    ' Восстанавливаем L:N и X:AA листа main
    Application.EnableEvents = False
    wsMain.Range("L4:N" & lastRow).Value = savedLtoN
    wsMain.Range("X4:AA" & lastRow).Value = savedXtoAA
    Application.EnableEvents = True

    ' =====================================================
    ' TC-30: ImportSheet с несуществующим ГРЗ
    ' =====================================================
    Mod_Import.SilenceMsgBox = True

    On Error Resume Next
    Call Mod_Import.ImportSheet("НЕСУЩЕСТВУЮЩИЙ")
    If Err.Number <> 0 Then
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", True, ""
    End If
    On Error GoTo 0

    Mod_Import.SilenceMsgBox = False

    Set wsMain = Nothing
    Set wsTemp = Nothing

    Debug.Print ""
End Sub
```

### 6.5. Обработка побочных эффектов

- **TC-29:** сохраняет `L:N` и `X:AA` листа `main` до теста и восстанавливает после; создаёт временный лист и **удаляет его** ([S3]); `Mod_Import.SilenceMsgBox = True` для подавления MsgBox при отсутствии таблиц; `Application.DisplayAlerts = False` при удалении листа.
- **TC-30:** `ImportSheet` вызывает `SearchSheetByGRZ`, который открывает `report.xlsx` и закрывает его внутри. При несуществующем ГРЗ `SearchSheetByGRZ` вернёт `Nothing`, `ImportSheet` покажет MsgBox (подавлен) и выйдет без ошибки. `Mod_Import.SilenceMsgBox = True` до вызова и `False` после.

---

## 7. Группа `RunConstantsTests` (TC-46)

### 7.1. Описание

Тестирует [`Mod_Constants.AddWorkEntry`](src/modules/Mod_Constants.bas:358) — добавляет запись `work.xlsm` в лист `libname`.

### 7.2. Данные

Рабочий лист `libname` в `work.xlsm`. Функция идемпотентна: проверяет последнюю строку на `work.xlsm` (строки 373–379 `Mod_Constants.bas`) и не добавляет дубликат.

### 7.3. Полный код процедуры

```vba
' ============================================================
' Группа: тесты Constants (TC-46)
' ============================================================
Private Sub RunConstantsTests()
    Dim wsLib As Worksheet
    Dim lastRowBefore As Long
    Dim lastRowAfter As Long
    Dim addedRow As Long
    Dim i As Long

    Debug.Print "--- Mod_Constants AddWorkEntry Tests ---"

    On Error Resume Next
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_LIBNAME)
    On Error GoTo 0

    If wsLib Is Nothing Then
        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", True, "", True, _
                  "Лист libname не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' Сохраняем последнюю строку до теста
    lastRowBefore = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row

    ' Вызываем AddWorkEntry
    On Error Resume Next
    Call Mod_Constants.AddWorkEntry()
    If Err.Number <> 0 Then
        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        lastRowAfter = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row
        addedRow = lastRowAfter

        ' Проверяем, что запись work.xlsm присутствует в последней строке
        Dim tc46Ok As Boolean
        Dim tc46Reason As String
        tc46Ok = (Trim(CStr(wsLib.Cells(addedRow, 1).Value)) = "work.xlsm")
        If Not tc46Ok Then
            tc46Reason = "Последняя строка не содержит work.xlsm: '" & _
                         CStr(wsLib.Cells(addedRow, 1).Value) & "'"
        End If

        ' Проверяем идемпотентность: повторный вызов не должен добавить дубликат
        If tc46Ok Then
            Call Mod_Constants.AddWorkEntry()
            Dim lastRowAfter2 As Long
            lastRowAfter2 = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row
            If lastRowAfter2 > lastRowAfter Then
                tc46Ok = False
                tc46Reason = "Повторный вызов добавил дубликат (идемпотентность нарушена)"
            End If
        End If

        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", tc46Ok, tc46Reason

        ' Восстанавливаем libname: если строка была добавлена тестом — удаляем её
        If lastRowAfter > lastRowBefore Then
            Application.DisplayAlerts = False
            wsLib.Rows(lastRowAfter).Delete
            Application.DisplayAlerts = True
        End If
    End If
    On Error GoTo 0

    Set wsLib = Nothing

    Debug.Print ""
End Sub
```

### 7.4. Обработка побочных эффектов

- `AddWorkEntry` добавляет строку в `libname`. Тест сохраняет последнюю строку до вызова и, если строка была добавлена тестом (`lastRowAfter > lastRowBefore`), **удаляет её** после проверки ([S3]).
- Проверяется идемпотентность: повторный вызов не должен увеличивать количество строк.
- `Application.DisplayAlerts = False` при удалении строки.

---

## 8. Сводная таблица побочных эффектов и митигации

| TC | Побочный эффект | Митигация |
|----|-----------------|-----------|
| TC-15..TC-18 | Нет (чистая функция) | — |
| TC-19..TC-21 | Нет (чистая функция) | — |
| TC-22 | Открывает UAZ.xlsm | Закрыть книгу после теста |
| TC-23 | Открывает UAZ.xlsm | Закрыть книгу после теста |
| TC-24 | Открывает UAZ.xlsm | Закрыть книгу после теста |
| TC-25 | Записывает B5:B17; может добавить модель в models; MsgBox | Сохранить/восстановить B5:B17; удалить добавленную строку models; `EnableEvents=False`; `SilenceMsgBox=True` |
| TC-26 | Очищает B5:B17 | Сохранить/восстановить B5:B17; `EnableEvents=False` |
| TC-27 | Нет (чтение spisok) | — |
| TC-28 | Нет (чтение spisok) | — |
| TC-29 | Очищает L:N и X:AA; создаёт временный лист; MsgBox | Сохранить/восстановить L:N и X:AA; удалить временный лист ([S3]); `Mod_Import.SilenceMsgBox=True` |
| TC-30 | Открывает report.xlsx; MsgBox | `SearchSheetByGRZ` закрывает report.xlsx; `Mod_Import.SilenceMsgBox=True` |
| TC-45 | Открывает report.xlsx | `SearchSheetByGRZ` закрывает report.xlsx; `SilenceMsgBox=True` |
| TC-46 | Добавляет строку в libname | Удалить добавленную строку; проверка идемпотентности |

---

## 9. Обновление заголовочного комментария модуля

1. Строка 7: `' Покрытие: TC-01 .. TC-44 (автоматические тесты)` → `' Покрытие: TC-01 .. TC-46 (автоматические тесты)`.
2. Строка 34: `Debug.Print "  Запуск набора тестов (TC-01..TC-44)"` → `Debug.Print "  Запуск набора тестов (TC-01..TC-46)"`.
3. Строка 869 (комментарий `RunAllTests_UI`): `' Запускает все тесты (TC-01..TC-44) и показывает результат` → `' Запускает все тесты (TC-01..TC-46) и показывает результат`.

---

## 10. Критерии приёмки

| Критерий | Ожидание |
|----------|----------|
| Total | **58** |
| Passed | **54** |
| Failed | **0** |
| Skipped | **4** (TC-41..TC-44) |
| exit code `run_tests.py` | **0** |
| Временные листы/данные | Удалены ([S3]) |
| Бизнес-модули | Не изменены ([Z5]) |
| `work.xlsm` | Закрыт перед импортом |

---

## 11. Порядок действий после реализации

1. **Резервная копия** текущего состояния (git branch/commit) перед изменениями.
2. **Реализация** — внести изменения в [`src/modules/Mod_FullTestRunner.bas`](src/modules/Mod_FullTestRunner.bas):
   - Добавить 6 новых групп процедур (разделы 2–7).
   - Зарегистрировать их в `RunAllTests()` (раздел 1.2).
   - Обновить заголовочный комментарий (раздел 9).
3. **Импорт** — `python scripts/impVBA.py` (UTF-8 → CP1251, импорт в `work.xlsm`).
4. **Проверка кодировки** — отсутствие «кракозябр» в русском тексте ([K3]).
5. **Запуск тестов** — `python scripts/run_tests.py` (COM, `B4=3`, `RunAllTests`, чтение `Z1`).
6. **Проверка критериев приёмки** (раздел 10): Total=58, Passed=54, Failed=0, Skipped=4, exit code 0.
7. **Исправления** — при FAIL анализировать `logs/test_results.log` и Immediate Window; исправлять **только** `Mod_FullTestRunner.bas`; повторить цикл импорт → тест.
8. **Экспорт** — `python scripts/export_vba.py` (CP1251 → UTF-8, синхронизация репозитория).
9. **Проверка чистоты** — убедиться, что временные листы/данные удалены ([S3]), временные директории `_temp_export/` и `_temp_import/` пусты.
10. **Обновление `CHANGELOG.md`** ([U2]) — только после явного одобрения пользователя ([E3]).
11. **Git-операции** — `git add` → `git commit` → `git push` → PR/merge в мэйн.

---

## 12. Риски и ограничения

| # | Риск | Митигация |
|---|------|-----------|
| 1 | TC-25 добавит модель в `models` | Детекция по последней строке и удаление добавленной строки |
| 2 | TC-29 очистит L:N и X:AA | Сохранение/восстановление диапазонов |
| 3 | TC-46 добавит дубликат в libname | Удаление добавленной строки; проверка идемпотентности |
| 4 | Открытые внешние книги (UAZ.xlsm, report.xlsx) | Закрытие после тестов |
| 5 | Отсутствие рабочих данных (UAZ.xlsm, spisok) | SKIP с пояснением вместо FAIL |
| 6 | Ошибки кодировки при импорте/экспорте | Строгое следование [K1]/[K2], проверка [K3] |
| 7 | COM-конфликты (открытый work.xlsm) | Убедиться, что work.xlsm закрыт перед импортом |
| 8 | `GetWorks` возвращает UDT в Variant | Проверка только `Count > 0` и `Code` первого элемента |