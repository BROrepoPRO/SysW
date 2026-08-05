# Анализ тестового покрытия: TC-15..TC-30, TC-45, TC-46

**Дата:** 2026-08-04
**Статус:** Анализ (без изменений кода)
**Автор:** SourceCraft Code Assistant (Architect)
**Связанный план:** [`plans/_archive/update_test_coverage_plan.md`](plans/_archive/update_test_coverage_plan.md)

---

## 1. Сводная таблица тестов

> **Принцип:** изменяется ТОЛЬКО [`Mod_FullTestRunner.bas`](src/modules/Mod_FullTestRunner.bas) ([Z5]). Бизнес-модули не трогаем. Временные данные удаляются ([S3]). Используем рабочие данные где возможно ([Z6]).

| TC | Функция (сигнатура) | Сценарий | Ожидаемый результат | Способ проверки | Группа |
|----|---------------------|----------|----------------------|-----------------|--------|
| TC-15 | `Mod_SheetOps.ExtractNumberFromGRZ(grz As String) As String` | `"А123АН77"` | `"123"` | `Result = "123"` | RunSheetOpsTests |
| TC-16 | `Mod_SheetOps.ExtractNumberFromGRZ(grz As String) As String` | `"А12АН34"` (2 цифры) | `""` | `Result = ""` | RunSheetOpsTests |
| TC-17 | `Mod_SheetOps.ExtractNumberFromGRZ(grz As String) As String` | `"А1234АН77"` (4 цифры) | `"1234"` | `Result = "1234"` | RunSheetOpsTests |
| TC-18 | `Mod_SheetOps.ExtractNumberFromGRZ(grz As String) As String` | `""` (пустая строка) | `""` | `Result = ""` | RunSheetOpsTests |
| TC-19 | `Mod_Constants.GetAggregateName(code As String) As String` | `"DIAG"` | `"Диагностика"` | `Result = "Диагностика"` | RunAggregateNameTests |
| TC-20 | `Mod_Constants.GetAggregateName(code As String) As String` | `"TO"` | `"ТО"` | `Result = "ТО"` | RunAggregateNameTests |
| TC-21 | `Mod_Constants.GetAggregateName(code As String) As String` | `"XXX"` (неизвестный) | `""` | `Result = ""` | RunAggregateNameTests |
| TC-22 | `Mod_ModelDB.GetWorkIdentities(groupName As String) As Collection` | `"UAZ"` | коллекция непустая, элементы типа `WorkIdentity` | `Count > 0`, `TypeName(item) = "WorkIdentity"`, поля непустые | RunModelDBReadTests |
| TC-23 | `Mod_ModelDB.GetPartIdentities(groupName As String) As Collection` | `"UAZ"` | коллекция непустая, элементы типа `PartIdentity` | `Count > 0`, `TypeName(item) = "PartIdentity"`, поля непустые | RunModelDBReadTests |
| TC-24 | `Mod_ModelDB.GetWorks(groupName As String, filters As Variant) As Collection` | `"UAZ"`, `filters = Empty` | коллекция непустая, элементы типа `WorkEntry` | `Count > 0`, `TypeName(item) = "WorkEntry"`, поля непустые | RunModelDBReadTests |
| TC-25 | `Mod_OrderHeader.FillHeaderFromOrder(orderNum As Variant) As Boolean` | существующий номер заказа | `True`, B5:B17 заполнены | `Result = True`, `Len(B5) > 0` | RunOrderHeaderTests |
| TC-26 | `Mod_OrderHeader.FillHeaderFromOrder(orderNum As Variant) As Boolean` | несуществующий номер (`999999`) | `False`, B5:B17 очищены | `Result = False`, `IsEmpty(B5)` | RunOrderHeaderTests |
| TC-27 | `Mod_OrderHeader.FindOrder(orderNum As String, Header As OrderHeader) As Boolean` | существующий заказ | `True`, структура заполнена | `Result = True`, `Header.OrderNumber <> ""` | RunOrderHeaderTests |
| TC-28 | `Mod_OrderHeader.FindOrder(orderNum As String, Header As OrderHeader) As Boolean` | несуществующий заказ | `False` | `Result = False` | RunOrderHeaderTests |
| TC-29 | `Mod_Import.ImportDataToMain(wsSource As Worksheet)` | временный лист с таблицами | данные перенесены в L:N и X:AA | проверка ячеек L4:N, X4:AA | RunImportDataTests |
| TC-30 | `Mod_Import.ImportSheet(grz As String)` | несуществующий ГРЗ | выход без ошибки (MsgBox подавлен) | `Err.Number = 0` | RunImportDataTests |
| TC-45 | `Mod_SheetOps.SearchSheetByGRZ(grz As String) As Worksheet` | несуществующий ГРЗ | `Nothing` | `ws Is Nothing` | RunSheetOpsTests |
| TC-46 | `Mod_Constants.AddWorkEntry()` | добавление записи work.xlsm | запись присутствует (идемпотентность) | проверка последней строки libname | RunConstantsTests |

---

## 2. Анализ данных

### 2.1. Файл `base/models/UAZ.xlsm`

Структура листов (по [`table.md`](table.md) и [`docs/ARCHITECTURE_SQLITE.md`](docs/ARCHITECTURE_SQLITE.md)):

| Лист | Назначение | Данные с | Ключевые столбцы |
|------|-----------|----------|-------------------|
| `UAZ` | Все работы группы | строка 4 | A=Code, B=Name, C=Unit, D=NormHours, E=Price, F=Note |
| `UAZw` | Модельные работы с аннотациями (тождества) | строка 4 | B=Артикул, C=Наименование, D=н/ч, G=Кол-во ЗН, I=Агрегат, J=Наим-ние (входящее) |
| `z4` | Все запчасти группы | строка 4 | — |
| `UAZz4` | Модельные запчасти с аннотациями (тождества) | строка 4 | B=Артикул, C=Наименование, F=Цена, G=Кол-во ЗН, I=АГРЕГАТ, J=№ кат., K=Наим-ние |

**Важно:** столбец A на листах UAZw/UAZz4 — информационный (№ п/п), данные начинаются с 4-й строки. Столбец I — код агрегата (ENG, HVAC и т.д.). Пустая строка или отсутствие агрегата в I = разделитель/пропуск.

По [`table.md`](table.md) колонка «Заполнено?» = «+» для ключевых полей (B, C, D, G, I, J), что подтверждает наличие реальных данных в UAZw и UAZz4.

### 2.2. Листы `spisok` и `models` в `work.xlsm`

**Лист `spisok`** (по [`Mod_OrderHeader.bas`](src/modules/Mod_OrderHeader.bas) и [`Mod_Constants.bas`](src/modules/Mod_Constants.bas)):

| Столбец | Константа | Назначение |
|---------|-----------|------------|
| A | `SPISOK_COL_NUM` (1) | № п/п (номер заказа) |
| B | `SPISOK_COL_MODEL` (2) | Модель |
| C | `SPISOK_COL_GRZ` (3) | ГРЗ |
| D | `SPISOK_COL_VIN` (4) | VIN |
| E | `SPISOK_COL_GARAGE` (5) | Гараж. № |
| F | `SPISOK_COL_YEAR` (6) | Год вып. |
| G | `SPISOK_COL_MILEAGE` (7) | Пробег |
| H | `SPISOK_COL_DATE` (8) | Дата |
| I | `SPISOK_COL_GROUP` (9) | Группа |
| J | `SPISOK_COL_NOTE` (10) | РЕЗЕРВ |

**Лист `models`** (по [`Mod_Constants.bas`](src/modules/Mod_Constants.bas)):

| Столбец | Константа | Назначение |
|---------|-----------|------------|
| A | `MODELS_COL_MODEL` (1) | Модель (данные с A3) |
| B | `MODELS_COL_GROUP` (2) | Группа |
| C | `MODELS_COL_PRICE` (3) | Цена н/ч |

### 2.3. Реальные данные для тестов

| Тест | Источник данных | Конкретное значение |
|------|-----------------|---------------------|
| TC-22 | `base/models/UAZ.xlsm` → лист `UAZw` | Реальные тождества работ (непустая коллекция) |
| TC-23 | `base/models/UAZ.xlsm` → лист `UAZz4` | Реальные тождества запчастей (непустая коллекция) |
| TC-24 | `base/models/UAZ.xlsm` → лист `UAZ` | Реальные работы (непустая коллекция) |
| TC-25 | `work.xlsm` → лист `spisok` | Первый непустой номер заказа из столбца A |
| TC-26 | — | Заведомо несуществующий номер `999999` |
| TC-27 | `work.xlsm` → лист `spisok` | Первый непустой номер заказа из столбца A |
| TC-28 | — | Заведомо несуществующий номер `999999` |
| TC-29 | Временный лист в `work.xlsm` | Контролируемые таблицы «Выполненные работы» и «Расходная накладная» |
| TC-30 | — | Несуществующий ГРЗ `"НЕСУЩЕСТВУЮЩИЙ"` |
| TC-45 | — | Несуществующий ГРЗ `"НЕСУЩЕСТВУЮЩИЙ"` |
| TC-46 | `work.xlsm` → лист `libname` | Запись `work.xlsm` |

---

## 3. Анализ рисков и побочных эффектов

| TC | Побочные эффекты | Риск | Митигация |
|----|------------------|------|-----------|
| TC-15..TC-18 | Нет (чистая функция) | Низкий | — |
| TC-19..TC-21 | Нет (чистая функция) | Низкий | — |
| TC-22 | Открывает `UAZ.xlsm` через `OpenModelGroupFile` (ReadOnly:=False) | Средний: файл остаётся открытым | Закрыть книгу после теста (`wb.Close SaveChanges:=False`), как в TC-35 |
| TC-23 | Открывает `UAZ.xlsm` | Средний: файл остаётся открытым | Закрыть книгу после теста |
| TC-24 | Открывает `UAZ.xlsm` | Средний: файл остаётся открытым | Закрыть книгу после теста |
| TC-25 | **Записывает в B5:B17 листа main**; может добавить модель в `models` (если не найдена); MsgBox при пустой группе/цене | **Высокий**: изменение данных main и models | Сохранить B5:B17 до теста, восстановить после; `Application.EnableEvents = False` при записи; `Mod_Constants.SilenceMsgBox = True` |
| TC-26 | Очищает B5:B17 листа main | Средний | Сохранить B5:B17 до теста, восстановить после |
| TC-27 | Нет (только чтение spisok) | Низкий | — |
| TC-28 | Нет (только чтение spisok) | Низкий | — |
| TC-29 | **Очищает L:N и X:AA листа main**; записывает данные; MsgBox при отсутствии таблиц | **Высокий**: изменение данных main | Сохранить L:N и X:AA до теста, восстановить после; `Mod_Import.SilenceMsgBox = True`; создать/удалить временный лист ([S3]) |
| TC-30 | Открывает `report.xlsx` (через `SearchSheetByGRZ`); MsgBox при ненайденном листе | Средний: открытие файла, MsgBox | `Mod_Import.SilenceMsgBox = True`; `SearchSheetByGRZ` сам закрывает report.xlsx |
| TC-45 | Открывает `report.xlsx` (ReadOnly:=True), закрывает его | Средний: открытие файла | `SearchSheetByGRZ` сам закрывает report.xlsx; при ошибке — MsgBox (подавить через `Mod_Constants.SilenceMsgBox`) |
| TC-46 | **Записывает в лист libname** (добавляет строку) | Средний: изменение libname | Проверить идемпотентность (повторный вызов не дублирует); при необходимости восстановить libname |

### 3.1. Ключевые риски

1. **TC-25 (FillHeaderFromOrder)** — самая опасная функция: при ненайденной модели **добавляет новую строку в лист `models`** (строки 134-155 `Mod_OrderHeader.bas`). Для теста нужно использовать модель, которая **уже существует** в `models`, чтобы избежать записи. Либо сохранить/восстановить лист `models`.
2. **TC-29 (ImportDataToMain)** — очищает L:N и X:AA на main. Обязательно сохранить исходные значения и восстановить.
3. **TC-46 (AddWorkEntry)** — добавляет строку в libname. Функция идемпотентна (проверяет последнюю строку на `work.xlsm`), но при повторном запуске теста может добавить дубликат, если между запусками libname изменился. Рекомендуется восстановить libname после теста.
4. **Открытие файлов** — `OpenModelGroupFile` (TC-22..24) и `SearchSheetByGRZ` (TC-30, TC-45) открывают внешние книги. Всегда закрывать их после теста, чтобы избежать утечки объектов и COM-конфликтов.

---

## 4. Рекомендации по структуре групп тестов

### 4.1. Новые группы процедур

| Процедура | Тесты | Модуль |
|-----------|-------|--------|
| `RunSheetOpsTests` | TC-15..TC-18, TC-45 | `Mod_SheetOps` |
| `RunAggregateNameTests` | TC-19..TC-21 | `Mod_Constants` |
| `RunModelDBReadTests` | TC-22..TC-24 | `Mod_ModelDB` |
| `RunOrderHeaderTests` | TC-25..TC-28 | `Mod_OrderHeader` |
| `RunImportDataTests` | TC-29, TC-30 | `Mod_Import` |
| `RunConstantsTests` | TC-46 | `Mod_Constants` |

### 4.2. Регистрация в `RunAllTests()`

Новые группы добавляются в [`RunAllTests()`](src/modules/Mod_FullTestRunner.bas:22) **после** существующих групп (после `RunAutoMatchTests`, строка 71) и **перед** `PrintFinalReport` (строка 74). Каждая группа оборачивается в логирование по образцу существующих:

```vba
Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests START")
RunSheetOpsTests
Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests END")
```

### 4.3. Стиль тестов (по образцу существующих)

- Каждый тест оборачивается в `On Error Resume Next` / `On Error GoTo 0`.
- При ошибке — `AddResult testId, testName, False, "Ошибка: " & Err.Description` + `Err.Clear`.
- При успехе — `AddResult testId, testName, (условие), "причина при FAIL"`.
- Для SKIP — `AddResult testId, testName, True, "", True, "причина"`.
- Сохранение/восстановление состояния листов (как в TC-14, TC-36).
- Подавление MsgBox: `Mod_Constants.SilenceMsgBox = True` (уже установлен в `RunAllTests`), для `Mod_Import` — дополнительно `Mod_Import.SilenceMsgBox = True`.

### 4.4. Сигнатура `AddResult`

```vba
Private Sub AddResult(testId As String, testName As String, _
                      passed As Boolean, Optional failReason As String = "", _
                      Optional skipped As Boolean = False, Optional skipReason As String = "")
```

---

## 5. Конкретные планы проверки по каждому тесту

### 5.1. TC-15..TC-18 — `ExtractNumberFromGRZ` (RunSheetOpsTests)

**Вызов:** `Mod_SheetOps.ExtractNumberFromGRZ(grz)`.

| TC | Вход | Ожидание | Обработка ошибки |
|----|------|----------|------------------|
| TC-15 | `"А123АН77"` | `"123"` | `On Error Resume Next`; при ошибке — FAIL |
| TC-16 | `"А12АН34"` | `""` | то же |
| TC-17 | `"А1234АН77"` | `"1234"` | то же |
| TC-18 | `""` | `""` | то же |

**Проверка:** `AddResult "TC-15", "ExtractNumberFromGRZ 'А123АН77'", (Result = "123"), "Ожидалось '123', получено '" & Result & "'"`.

### 5.2. TC-19..TC-21 — `GetAggregateName` (RunAggregateNameTests)

**Вызов:** `Mod_Constants.GetAggregateName(code)`.

| TC | Вход | Ожидание |
|----|------|----------|
| TC-19 | `"DIAG"` | `"Диагностика"` |
| TC-20 | `"TO"` | `"ТО"` |
| TC-21 | `"XXX"` | `""` |

**Примечание:** функция использует `UCase$(Trim$(code))`, поэтому регистр не важен. Для TC-20 важно, что `AGG_TO = "TO"` (не `"ТО"` — это русское название, а код латиницей).

### 5.3. TC-22..TC-24 — чтение ModelDB (RunModelDBReadTests)

**Общий шаблон:**
```vba
On Error Resume Next
Set wb = Mod_ModelDB.OpenModelGroupFile("UAZ")
' ... вызов функции ...
If Not wb Is Nothing Then wb.Close SaveChanges:=False
Set wb = Nothing
```

**TC-22 (GetWorkIdentities):**
```vba
Set col = Mod_ModelDB.GetWorkIdentities("UAZ")
' Проверка: col.Count > 0
' Проверка: TypeName(col(1)) = "WorkIdentity"
' Проверка: col(1).OutArticle <> "" (поле B)
' Проверка: col(1).Aggregate <> "" (поле I)
```

**TC-23 (GetPartIdentities):**
```vba
Set col = Mod_ModelDB.GetPartIdentities("UAZ")
' Проверка: col.Count > 0
' Проверка: TypeName(col(1)) = "PartIdentity"
' Проверка: col(1).OutArticle <> "" (поле B)
' Проверка: col(1).Aggregate <> "" (поле I)
```

**TC-24 (GetWorks):**
```vba
Set col = Mod_ModelDB.GetWorks("UAZ", Empty)
' Проверка: col.Count > 0
' Проверка: TypeName(col(1)) = "WorkEntry"
' Проверка: col(1).Code <> "" (поле A)
```

**Обработка ошибки:** если `col.Count = 0` — FAIL с пояснением «коллекция пуста (возможно, нет данных в UAZ.xlsm)». Рекомендуется SKIP, если файл группы не существует (проверка `Mod_ModelDB.ModelGroupFileExists("UAZ")`).

### 5.4. TC-25..TC-28 — OrderHeader (RunOrderHeaderTests)

**TC-25 (FillHeaderFromOrder, существующий заказ):**
1. Получить реальный номер заказа: первый непустой номер из столбца A листа `spisok`.
2. Сохранить B5:B17 листа main.
3. `Application.EnableEvents = False` (чтобы не сработал Worksheet_Change).
4. `Result = Mod_OrderHeader.FillHeaderFromOrder(orderNum)`.
5. Проверить `Result = True` и `Len(wsMain.Range("B5").Value) > 0`.
6. Восстановить B5:B17, `Application.EnableEvents = True`.

**TC-26 (FillHeaderFromOrder, несуществующий заказ):**
1. Сохранить B5:B17.
2. `Result = Mod_OrderHeader.FillHeaderFromOrder(999999)`.
3. Проверить `Result = False` и `IsEmpty(wsMain.Range("B5").Value)`.
4. Восстановить B5:B17.

**TC-27 (FindOrder, существующий заказ):**
```vba
Dim Header As OrderHeader
Result = Mod_OrderHeader.FindOrder(orderNum, Header)
' Проверка: Result = True
' Проверка: Header.OrderNumber <> ""
' Проверка: Header.ModelName <> ""
```

**TC-28 (FindOrder, несуществующий заказ):**
```vba
Result = Mod_OrderHeader.FindOrder("999999", Header)
' Проверка: Result = False
```

**Риск TC-25:** если модель из spisok не найдена в `models`, функция добавит новую строку в `models`. Для теста использовать модель, которая гарантированно есть в `models`, либо сохранить/восстановить лист `models`.

### 5.5. TC-29 — `ImportDataToMain` (RunImportDataTests)

1. Создать временный лист `wsTemp` в `work.xlsm`.
2. Заполнить таблицу «Выполненные работы» (структура по комментариям в [`Mod_Import.bas`](src/modules/Mod_Import.bas:92-99)):
   - Строка 1: заголовок «№ | № кат. | Наименование | Кол. оп. | Цена | Норма | н/ч | Всего | в т.ч. НДС»
   - Строка 2: подзаголовок «1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9»
   - Строка 3+: данные (D=Наименование, I=Всего, M=в т.ч. НДС)
3. Заполнить таблицу «Расходная накладная» (структура по [`Mod_Import.bas`](src/modules/Mod_Import.bas:156-163)):
   - Строка 1: заголовок «№ | № кат. | Наименование | Кол-во | Ед.изм. | Цена | Всего | в т.ч. НДС»
   - Строка 2: подзаголовок «1 | 2 | 3 | 4 | 5 | 6 | 7 | 8»
   - Строка 3+: данные (C=Наименование, D=Кол-во, J=Всего, M=в т.ч. НДС)
4. Сохранить L:N и X:AA листа main.
5. `Mod_Import.SilenceMsgBox = True`.
6. `Call Mod_Import.ImportDataToMain(wsTemp)`.
7. Проверить перенос: `wsMain.Cells(4, 12).Value` (L4) = D из источника, `wsMain.Cells(4, 24).Value` (X4) = C из источника.
8. Восстановить L:N и X:AA.
9. Удалить временный лист ([S3]).

**Риск:** функция очищает L:N и X:AA на main. Обязательно восстановить.

### 5.6. TC-30 — `ImportSheet` (RunImportDataTests)

1. `Mod_Import.SilenceMsgBox = True`.
2. `Call Mod_Import.ImportSheet("НЕСУЩЕСТВУЮЩИЙ")`.
3. Проверить `Err.Number = 0` (выход без ошибки).
4. `Mod_Import.SilenceMsgBox = False`.

**Примечание:** `ImportSheet` вызывает `SearchSheetByGRZ`, который открывает `report.xlsx`. При несуществующем ГРЗ `SearchSheetByGRZ` вернёт `Nothing`, и `ImportSheet` покажет MsgBox (подавлен) и выйдет. `report.xlsx` закрывается внутри `SearchSheetByGRZ`.

### 5.7. TC-45 — `SearchSheetByGRZ` (RunSheetOpsTests)

1. `Mod_Constants.SilenceMsgBox = True` (уже установлен в RunAllTests).
2. `Set ws = Mod_SheetOps.SearchSheetByGRZ("НЕСУЩЕСТВУЮЩИЙ")`.
3. Проверить `ws Is Nothing`.
4. `Set ws = Nothing`.

**Примечание:** `SearchSheetByGRZ` открывает `report.xlsx` (ReadOnly:=True) и закрывает его. При несуществующем ГРЗ вернёт `Nothing`. Если `report.xlsx` отсутствует — будет ошибка (ErrHandler покажет MsgBox, подавлен), вернёт `Nothing`.

### 5.8. TC-46 — `AddWorkEntry` (RunConstantsTests)

1. Сохранить последнюю строку листа `libname`.
2. `Call Mod_Constants.AddWorkEntry()`.
3. Проверить, что последняя строка содержит `work.xlsm` в столбце A.
4. Вызвать повторно — проверить идемпотентность (не добавился дубликат).
5. Восстановить libname (удалить добавленную строку, если она была добавлена тестом).

**Риск:** функция добавляет строку в libname. Идемпотентна (проверяет последнюю строку на `work.xlsm`), но при изменении libname между запусками может добавить дубликат. Рекомендуется восстановить libname после теста.

---

## 6. Итоговые рекомендации

1. **Порядок добавления групп** в `RunAllTests()`: `RunSheetOpsTests`, `RunAggregateNameTests`, `RunModelDBReadTests`, `RunOrderHeaderTests`, `RunImportDataTests`, `RunConstantsTests`.
2. **Подавление MsgBox:** `Mod_Constants.SilenceMsgBox = True` уже установлен в `RunAllTests`. Для `Mod_Import`-функций (TC-29, TC-30) дополнительно установить `Mod_Import.SilenceMsgBox = True` и сбросить после.
3. **Сохранение/восстановление состояния:** для TC-25, TC-26 (B5:B17), TC-29 (L:N, X:AA), TC-46 (libname) обязательно сохранять и восстанавливать исходные значения.
4. **Закрытие внешних книг:** для TC-22..TC-24 (UAZ.xlsm) и TC-30, TC-45 (report.xlsx) — закрывать книги после теста.
5. **Зависимость от рабочих данных:** TC-22..TC-24 зависят от наличия данных в `UAZ.xlsm`; TC-25, TC-27 — от данных в `spisok`/`models`. При отсутствии данных — SKIP с пояснением.
6. **Ожидаемые итоги:** Total = 40 (существующие) + 18 (новые) = 58; Passed = 54; Failed = 0; Skipped = 4 (TC-41..TC-44).