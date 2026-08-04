# Диагностика ошибки DISP_E_EXCEPTION (0x80020009) при выполнении VBA-тестов

## 1. Контекст

- Тесты запускаются через [`run_tests.py`](scripts/run_tests.py) → COM → `Excel.Run("RunAllTests")`
- Компиляция проходит успешно
- Через ~3 минуты возникает `DISP_E_EXCEPTION` (0x80020009)
- `MsgBox` подавлены через `Mod_Constants.SilenceMsgBox = True`
- `Application.DisplayAlerts = False` установлен в Python-скрипте

## 2. Порядок выполнения тестов (из [`Mod_FullTestRunner.RunAllTests`](src/modules/Mod_FullTestRunner.bas:22))

```
RunUtilsTests       (TC-01..TC-08)  — быстрые, без открытия файлов
RunLoggerTests      (TC-09..TC-11)  — файловые операции с логом
RunUtilsEdgeTests   (TC-12)         — быстрые
RunLibNameTests     (TC-13)         — работа с листом libname
RunImportVHTests    (TC-14)         — вызов ImportFromB2_UI (открывает report.xlsx)
RunModelDBTests     (TC-31..TC-35)  — открывает UAZ.xlsm (TC-35)
RunPickWorkTests    (TC-36..TC-38)  — открывает UAZ.xlsm + Activate/Select (TC-38)
RunAutoMatchTests   (TC-39..TC-44)  — открывает UAZ.xlsm, читает UAZw/UAZz4, пишет формулы (TC-39, TC-40)
```

**Время ~3 минуты** указывает на то, что сбой происходит в районе **TC-35 — TC-40**, т.е. в тестах, которые работают с файлами групп (`base/models/UAZ.xlsm`).

## 3. Анализ возможных источников ошибки

### 3.1. TC-35: `OpenModelGroupFile("UAZ")` — открытие файла группы

Файл: [`Mod_ModelDB.OpenModelGroupFile`](src/modules/Mod_ModelDB.bas:70)

```vba
Set wb = Workbooks.Open(filePath, ReadOnly:=False)
```

**Проблема:** Файл открывается **не ReadOnly** (`ReadOnly:=False`). Если:
- В `UAZ.xlsm` есть макрос `Workbook_Open`, который вызывает `MsgBox` или ошибку
- Файл уже открыт в другом процессе
- Файл повреждён

**Симптом:** `DISP_E_EXCEPTION` при попытке открыть файл через COM.

### 3.2. TC-38: `PickWork_UI` — активация листа в невидимом Excel

Файл: [`Mod_PickWork.PickWork_UI`](src/modules/Mod_PickWork.bas:58)

```vba
wsWork.Activate   ' строка 110
wsWork.Select     ' строка 111
```

**Проблема:** `Activate` и `Select` на листе книги, открытой через COM-автоматизацию (Excel невидим, `Visible = False`). Это может вызвать `DISP_E_EXCEPTION`, так как Excel не может активировать лист в невидимом режиме.

**Важно:** `PickWork_UI` не проверяет `Mod_Constants.SilenceMsgBox` для `Activate`/`Select` — эти операции выполняются безусловно.

### 3.3. TC-39: `AutoMatchWorks` — запись формул с неверным разделителем

Файл: [`Mod_AutoMatch.AutoMatchWorks`](src/modules/Mod_AutoMatch.bas:182)

```vba
wsMain.Cells(i, MAIN_W_SUM).Formula = _
    "=ROUND(G" & i & "*H" & i & "*I" & i & ";2)"
```

**ПроблемА:** В русской локали Excel разделитель аргументов в формулах — `;`. Однако при запуске через COM-автоматизацию из Python, Excel может использовать **английскую локаль**, где разделитель — `,`. Формула `=ROUND(G4*H4*I4;2)` будет воспринята как неверная.

**Это КЛАССИЧЕСКАЯ причина `DISP_E_EXCEPTION`** при работе с COM. VBA-код выполняется в контексте Excel, но свойство `Formula` ожидает формулу в **английском формате** (с `,`), если Excel запущен с английским языком интерфейса.

**Альтернатива:** Использовать `FormulaLocal` для русскоязычных формул или `Formula` с английским синтаксисом.

Аналогичная проблема в [`AutoMatchParts`](src/modules/Mod_AutoMatch.bas:312):

```vba
wsMain.Cells(i, MAIN_P_SUM).Formula = _
    "=ROUND(T" & i & "*U" & i & ";2)"
```

### 3.4. TC-14: `ImportFromB2_UI` — открытие report.xlsx

Файл: [`Mod_Import.ImportFromB2_UI`](src/modules/Mod_Import.bas:326)

**Проблема:** `ImportFromB2_UI` открывает `report.xlsx` через `Workbooks.Open(reportPath, ReadOnly:=True)`. Если `report.xlsx` не существует или повреждён — будет ошибка. Однако TC-14 очищает B4 перед вызовом, поэтому `ImportFromB2_UI` должна выйти раньше (проверка `grz = ""` на строке 348).

**Вывод:** TC-14, вероятно, не является источником ошибки, так как при пустом B4 процедура выходит через `GoTo CleanUp`.

### 3.5. `WriteResultsToSheet` — запись в Z1

Файл: [`Mod_FullTestRunner.WriteResultsToSheet`](src/modules/Mod_FullTestRunner.bas:62)

```vba
On Error Resume Next
Set ws = ThisWorkbook.Sheets("main")
If Not ws Is Nothing Then
    ws.Range("Z1").Value = ReportMsg
End If
```

**Проблема:** Если лист `main` защищён или ячейка Z1 объединена/заблокирована — запись вызовет ошибку. Однако `On Error Resume Next` должен перехватить её.

**Вывод:** Маловероятный источник, так как ошибка перехватывается.

### 3.6. `Mod_Logger.RotateLogIfNeeded` — файловые операции

Файл: [`Mod_Logger.RotateLogIfNeeded`](src/modules/Mod_Logger.bas:73)

```vba
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile LogPath, OldLogPath
```

**Проблема:** Если файл лога заблокирован антивирусом или другим процессом — `MoveFile` вызовет ошибку. Однако `On Error Resume Next` перехватывает её.

**Вывод:** Маловероятный источник, так как ошибка перехватывается.

### 3.7. `Mod_Import.ImportDataToMain` — поиск таблиц через `Find`

Файл: [`Mod_Import.ImportDataToMain`](src/modules/Mod_Import.bas:60)

```vba
Set cell = wsSource.Cells.Find(What:="Выполненные работы", ...)
```

**Проблема:** Если лист-источник пуст или не содержит искомых строк, `Find` вернёт `Nothing`. Код проверяет это (`If Not cell Is Nothing`), но внутри блока есть вложенный `Find`, который может не найти `"Итого"` — это обрабатывается через `On Error Resume Next`.

**Вывод:** Маловероятный прямой источник `DISP_E_EXCEPTION`, но может быть косвенной причиной.

## 4. Наиболее вероятные причины (ранжированные)

| № | Причина | Тест | Вероятность |
|---|--------|------|-------------|
| 1 | **Формула с `;` вместо `,`** в `AutoMatchWorks`/`AutoMatchParts` | TC-39/TC-40 | **Высокая** |
| 2 | **`Activate`/`Select` в невидимом Excel** в `PickWork_UI` | TC-38 | **Высокая** |
| 3 | **Открытие `UAZ.xlsm` с `ReadOnly:=False`** при наличии `Workbook_Open` | TC-35 | Средняя |
| 4 | **Конфликт при повторном открытии `UAZ.xlsm`** (открывается в TC-35, TC-38, TC-39, TC-40) | TC-35..TC-40 | Средняя |
| 5 | **`report.xlsx` не существует или повреждён** | TC-14 | Низкая |

## 5. Логирование для подтверждения диагноза

Для точного определения места сбоя предлагаю добавить временные логи в следующие точки:

### 5.1. В [`Mod_FullTestRunner.RunAllTests`](src/modules/Mod_FullTestRunner.bas:22)

Добавить `Debug.Print` с меткой времени перед каждой группой тестов:

```vba
Debug.Print "[TIMESTAMP] " & Now & " Starting RunModelDBTests..."
' ... существующий код RunModelDBTests
Debug.Print "[TIMESTAMP] " & Now & " Finished RunModelDBTests"
```

### 5.2. В [`Mod_AutoMatch.AutoMatchWorks`](src/modules/Mod_AutoMatch.bas:105)

Добавить логирование перед записью формулы:

```vba
Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: Запись формулы в строку " & i)
wsMain.Cells(i, MAIN_W_SUM).Formula = _
    "=ROUND(G" & i & "*H" & i & "*I" & i & ";2)"
```

### 5.3. В [`Mod_PickWork.PickWork_UI`](src/modules/Mod_PickWork.bas:58)

Добавить логирование перед Activate/Select:

```vba
Call Mod_Logger.WriteLog("Mod_PickWork", "PickWork_UI: Попытка Activate листа " & sheetName)
wsWork.Activate
wsWork.Select
```

## 6. Предлагаемые исправления

### 6.1. Исправление формул в `AutoMatchWorks` и `AutoMatchParts`

**Проблема:** Использование `;` в свойстве `Formula`. Нужно использовать `FormulaLocal` для русской локали или `,` для английской.

**Вариант А (рекомендуемый):** Использовать `FormulaLocal`:

```vba
wsMain.Cells(i, MAIN_W_SUM).FormulaLocal = _
    "=ОКРУГЛ(G" & i & "*H" & i & "*I" & i & ";2)"
```

**Вариант Б:** Использовать `,` в `Formula` (английский синтаксис):

```vba
wsMain.Cells(i, MAIN_W_SUM).Formula = _
    "=ROUND(G" & i & "*H" & i & "*I" & i & ",2)"
```

### 6.2. Исправление `PickWork_UI` для COM-режима

**Проблема:** `Activate`/`Select` в невидимом Excel.

**Решение:** Пропускать `Activate`/`Select`, если `SilenceMsgBox = True` (тестовый режим):

```vba
If Not Mod_Constants.SilenceMsgBox Then
    wsWork.Activate
    wsWork.Select
End If
```

### 6.3. Исправление `OpenModelGroupFile` — открытие ReadOnly

**Проблема:** Файл открывается на запись, что может вызвать конфликты.

**Решение:** Открывать `ReadOnly:=True` для тестового режима:

```vba
Dim openReadOnly As Boolean
openReadOnly = Mod_Constants.SilenceMsgBox  ' В тестовом режиме — только чтение
Set wb = Workbooks.Open(filePath, ReadOnly:=openReadOnly)
```

## 7. Вывод

Наиболее вероятная причина `DISP_E_EXCEPTION` — **запись формул с неверным разделителем аргументов** в [`AutoMatchWorks`](src/modules/Mod_AutoMatch.bas:182) и [`AutoMatchParts`](src/modules/Mod_AutoMatch.bas:312). При запуске через COM-автоматизацию Excel может использовать английскую локаль, где разделитель — `,`, а не `;`.

Вторая по вероятности причина — **`Activate`/`Select` в невидимом Excel** в [`PickWork_UI`](src/modules/Mod_PickWork.bas:110-111).

Рекомендуется:
1. Добавить временные логи (п. 5) для подтверждения диагноза
2. Применить исправления (п. 6) после подтверждения