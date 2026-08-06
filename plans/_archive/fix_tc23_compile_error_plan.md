# План исправления ошибки VBA `0x800A03EE` в TC-23 (GetPartIdentities)

**Файл:** `src/modules/Mod_ModelDB.bas`
**Симптом:** при запуске `python scripts/run_tests.py` (макрос `RunAllTests`) возникает COM-ошибка `(-2147352567, 'Ошибка.', (0, None, None, None, 0, -2146788248), None)` = `0x800A03EE` (Application-defined or object-defined error). TC-22 проходит, TC-23 падает; лог обрывается после `OpenModelGroupFile: Файл UAZ.xlsm уже открыт`.

**Диагноз (подтверждён чтением кода):**
- `GetWorkIdentities` (строки 196–275) защищает каждую операцию с листом через `On Error Resume Next` + проверку `Err.Number` (строки 233–239, 248–265).
- `GetPartIdentities` (строки 284–346) **не имеет** такой защиты: строки 317 и 324–338 выполняются под активным `On Error GoTo ErrHandler` (строка 285). Любая ошибка Excel в данных листа `UAZz4` (столбец F/Price, либо `lastRow` на строке 317) уходит в `ErrHandler` и возвращает пустую коллекцию, а при COM-автоматизации — всплывает как `0x800A03EE`.

---

## Приоритет 1 (критический): защита `GetPartIdentities` по аналогии с `GetWorkIdentities`

Файл: `src/modules/Mod_ModelDB.bas`, функция `GetPartIdentities` (строки 284–346).

### Шаг 1.1 — Защитить вычисление `lastRow` (строка 317)

**Текущий код (строки 316–321):**
```vba
    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        Set GetPartIdentities = result
        Exit Function
    End If
```

**Изменённый код:**
```vba
    ' Определяем последнюю строку (данные с 4-й строки)
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If Err.Number <> 0 Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetPartIdentities: Error getting lastRow: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo ErrHandler
    If lastRow < 4 Then
        Set GetPartIdentities = result
        Exit Function
    End If
```

### Шаг 1.2 — Защитить цикл чтения данных (строки 324–338)

**Текущий код (строки 323–338):**
```vba
    ' Читаем данные
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then ' I — АГРЕГАТ
                Set identity = New PartIdentity
                identity.OutArticle = CStr(ws.Cells(i, 2).Value)  ' B
                identity.OutName = CStr(ws.Cells(i, 3).Value)     ' C
                identity.QtyZN = Val(ws.Cells(i, 7).Value)        ' G
                identity.Price = Val(ws.Cells(i, 6).Value)        ' F
                identity.Aggregate = CStr(ws.Cells(i, 9).Value)   ' I
                identity.InCatNum = CStr(ws.Cells(i, 10).Value)   ' J
                identity.InName = CStr(ws.Cells(i, 11).Value)     ' K
                result.Add identity
            End If
        End If
    Next i
```

**Изменённый код:**
```vba
    ' Читаем данные
    For i = 4 To lastRow
        On Error Resume Next
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then ' I — АГРЕГАТ
                Set identity = New PartIdentity
                identity.OutArticle = CStr(ws.Cells(i, 2).Value)  ' B
                identity.OutName = CStr(ws.Cells(i, 3).Value)     ' C
                identity.QtyZN = Val(ws.Cells(i, 7).Value)        ' G
                identity.Price = Val(ws.Cells(i, 6).Value)        ' F
                identity.Aggregate = CStr(ws.Cells(i, 9).Value)   ' I
                identity.InCatNum = CStr(ws.Cells(i, 10).Value)   ' J
                identity.InName = CStr(ws.Cells(i, 11).Value)     ' K
                result.Add identity
            End If
        End If
        If Err.Number <> 0 Then
            Call Mod_Logger.WriteLog("Mod_ModelDB", "GetPartIdentities: Error at row " & CStr(i) & ": " & Err.Description)
            Err.Clear
        End If
        On Error GoTo ErrHandler
    Next i
```

### Шаг 1.3 — Защитить присваивание `identity.Price` (строка 331)

Строка `identity.Price = Val(ws.Cells(i, 6).Value)` уже попадает под `On Error Resume Next` из Шага 1.2, поэтому отдельная обёртка не требуется. Однако для устойчивости к ошибкам Excel в столбце F (например, `#VALUE!` в ячейке) рекомендуется заменить прямое чтение на безопасное:

**Текущая строка 331:**
```vba
                identity.Price = Val(ws.Cells(i, 6).Value)        ' F
```

**Изменённая строка:**
```vba
                identity.Price = Val(Mod_Utils.SafeCellValue(ws.Cells(i, 6)))  ' F
```

> **Примечание:** `Mod_Utils.SafeCellValue` — вспомогательная функция, которая должна существовать в `Mod_Utils.bas`. Если её нет, либо добавить её (см. Шаг 1.4), либо оставить `Val(ws.Cells(i, 6).Value)` — в этом случае защиту обеспечит `On Error Resume Next` из Шага 1.2. **Не добавлять** `On Error Resume Next` внутри уже защищённого блока.

### Шаг 1.4 (опционально) — Вспомогательная функция `SafeCellValue` в `Mod_Utils.bas`

Если решено использовать `Mod_Utils.SafeCellValue`, добавить в `src/modules/Mod_Utils.bas`:

```vba
' Возвращает значение ячейки, подавляя ошибки Excel (#N/A, #VALUE! и т.п.)
Public Function SafeCellValue(ByVal rng As Range) As Variant
    On Error Resume Next
    SafeCellValue = rng.Value
    If Err.Number <> 0 Then
        Err.Clear
        SafeCellValue = ""
    End If
End Function
```

---

## Приоритет 2: проверка данных листа `UAZz4` в `base/models/UAZ.xlsm`

Ошибка `0x800A03EE` может быть вызвана ошибками Excel (`#N/A`, `#VALUE!`, `#REF!`) в ячейках столбцов B, F, I листа `UAZz4`. Проверить двумя способами:

### Способ A — скрипт Python (рекомендуется)

Создать временный скрипт (например, `scripts/check_uaz_z4.py`) или выполнить инлайн-проверку через `win32com`:

```python
import win32com.client as win32

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(r"l:/PROject/SysW/base/models/UAZ.xlsm")
    ws = wb.Sheets("UAZz4")
    last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row  # xlUp = -4162
    print(f"UAZz4 lastRow = {last_row}")
    for col in (2, 6, 9):  # B, F, I
        for r in range(4, last_row + 1):
            v = ws.Cells(r, col).Value
            if isinstance(v, str) and v.startswith("#"):
                print(f"ОШИБКА Excel: {ws.Cells(r, col).Address} = {v}")
    wb.Close(False)
finally:
    excel.Quit()
```

**Критерий:** скрипт не должен выводить строки `ОШИБКА Excel`. Если такие строки есть — исправить данные в `UAZ.xlsm` (заменить ошибочные формулы на значения) либо добавить защиту из Приоритета 1.

### Способ B — ручная проверка

1. Открыть `base/models/UAZ.xlsm` в Excel.
2. Перейти на лист `UAZz4`.
3. Выделить столбцы B, F, I (строки 4 и ниже до конца данных).
4. Нажать `Ctrl+F` → «Найти» → ввести `#` → «Найти все».
5. Убедиться, что нет ячеек со значениями `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`.
6. Особое внимание — столбец F (Price): проверить, что все ячейки содержат числа или пустые, а не формулы с ошибками.

---

## Приоритет 3 (гигиена кода, отдельный рефакторинг): перенос `Dim` в начало процедур

**Файл:** `src/modules/Mod_FullTestRunner.bas`

**Важно:** это **отдельный шаг**, не смешивать с исправлением ошибки (Приоритет 1). `Dim` внутри блоков `If...Else`/`For` — легальный VBA-код (scope переменной — вся процедура) и **не является** причиной ошибки `0x800A03EE`. Рефакторинг выполняется только для чистоты кода и единообразия стиля.

### Объёмы работ

В файле ~40 объявлений `Dim`, часть из которых находится внутри блоков `If...Else`/`For` (не в начале процедуры). Ключевые места:

| Процедура | Строки `Dim` внутри блоков |
|---|---|
| `RunUtilsTests` | 196, 225, 277–278 |
| `RunLoggerTests` | 316–319, 338, 340, 375, 398–399, 416 |
| `RunLibNameTests` | 502–507, 520, 529, 532 |
| `RunModelDBTests` | 621–622, 640–641, 689 |
| `RunModelDBReadTests` | 1020–1021, 1025, 1052–1053, 1057, 1084–1085 |
| `RunOrderHeaderTests` | 1184–1185, 1216–1217, 1245–1246 |
| `RunImportDataTests` | 1397–1398 |
| `RunConstantsTests` | 1485–1486, 1496 |

### Правило рефакторинга

Для каждой процедуры:
1. Собрать **все** `Dim`-объявления процедуры (включая те, что сейчас внутри блоков).
2. Перенести их в начало процедуры, сразу после строки `Private Sub ...` / `Public Sub ...`, сгруппировав по типу.
3. Удалить `Dim` из исходных мест (оставить только присваивания).
4. Не менять логику, порядок присваиваний и обработку ошибок.

**Пример для `RunUtilsTests` (строки 156–287):**

Текущее начало (строки 156–160):
```vba
Private Sub RunUtilsTests()
    Dim LogPath As String
    Dim Result As Boolean
    Dim PathResult As String
    Dim UserResult As String
```

После рефакторинга (добавить объявления из строк 196, 225, 277–278):
```vba
Private Sub RunUtilsTests()
    Dim LogPath As String
    Dim Result As Boolean
    Dim PathResult As String
    Dim UserResult As String
    Dim FmtResult As String
    Dim ws As Worksheet
    Dim PathOk As Boolean
    Dim UserOk As Boolean
```

И удалить строки `Dim FmtResult As String` (196), `Dim ws As Worksheet` (225), `Dim PathOk As Boolean` / `Dim UserOk As Boolean` (277–278) из их текущих позиций.

### Критерий завершения рефакторинга

- В файле `Mod_FullTestRunner.bas` нет ни одного `Dim` внутри блоков `If...Else`/`For` — все объявления находятся в начале своих процедур.
- Повторный запуск `python scripts/run_tests.py` даёт тот же результат, что и до рефакторинга (без изменения поведения).

---

## Порядок выполнения

1. **Приоритет 1** — внести изменения в `GetPartIdentities` (Шаги 1.1, 1.2, при необходимости 1.3/1.4). Это устраняет саму ошибку.
2. **Приоритет 2** — проверить данные листа `UAZz4` (Способ A или B). Исправить найденные ошибки Excel в данных.
3. **Приоритет 3** — отдельным коммитом выполнить рефакторинг `Dim` в `Mod_FullTestRunner.bas`.

> Рекомендуется выполнять Приоритет 1 и Приоритет 3 **разными коммитами**, чтобы история изменений была чистой и рефакторинг не смешивался с исправлением ошибки.

---

## Критерии верификации

1. **Повторный запуск:** `python scripts/run_tests.py` завершается без COM-ошибки `0x800A03EE`.
2. **Лог:** в `logs/` появляется запись `GetPartIdentities: END count=N` (N > 0) для TC-23, а не обрыв после `OpenModelGroupFile: Файл UAZ.xlsm уже открыт`.
3. **TC-22 и TC-23:** оба проходят (в отчёте `Total/Passed/Failed` нет новых `Failed`).
4. **Приоритет 2:** скрипт проверки `UAZz4` не находит ошибок Excel в столбцах B, F, I.
5. **Приоритет 3:** в `Mod_FullTestRunner.bas` все `Dim` перенесены в начало процедур; повторный запуск тестов даёт идентичный результат.