# Анализ ошибки импорта данных в Mod_Import.bas

## Диагноз

**Сдвиг данных таблицы "Выполненные работы" на 1 строку вверх** — данные услуг (колонки L:N на листе main) начинаются со строки 3 вместо строки 2, а в строку 2 попадают номера столбцов из подзаголовка таблицы-источника ("1 | 2 | 3 | ...").

**Проявление в ошибочном выводе (`sheet_main_er.md`):**
- `L2=3`, `M2=4,00` — это номера столбцов из подзаголовка (D=3, E=4), записанные как данные
- `L3=Диагностика кондиционера`, `M3=1,00` — реальные данные, но со сдвигом на 1 строку вниз
- Аналогичный сдвиг наблюдается для таблицы материалов (X:AA), но он маскируется тем, что `srcLastRow` уже инициализирован к моменту обработки этой таблицы

## Причина

**Неинициализированная переменная `srcLastRow`** в процедуре `ImportDataToMain`.

### Детали

В VBA переменная `Dim srcLastRow As Long` (строка 62) инициализируется значением `0`.

В строках 100-106 цикл пропуска пустых строк использует `srcLastRow` как верхнюю границу:

```vba
' строка 100
dataStartRow = cell.Row + 1
' строка 101 — srcLastRow = 0 на этот момент!
Do While dataStartRow <= srcLastRow
    If Trim(wsSource.Cells(dataStartRow, 4).Value) <> "" Then Exit Do
    dataStartRow = dataStartRow + 1
Loop
' строка 106
dataStartRow = dataStartRow + 2
```

Поскольку `dataStartRow = cell.Row + 1 > 0`, условие `dataStartRow <= 0` ложно, и цикл **никогда не выполняется**. Пустые строки между заголовком "Выполненные работы" и шапкой таблицы не пропускаются.

Корректное значение `srcLastRow` присваивается только на строке 110:
```vba
srcLastRow = wsSource.Cells(wsSource.Rows.count, 4).End(xlUp).Row
```

**Для таблицы "Расходная накладная"** эта же ошибка не проявляется, потому что к моменту её обработки (строки 144-150) `srcLastRow` уже содержит корректное значение от предыдущей таблицы (строка 110 или 115).

### Сценарий возникновения

Если лист-источник имеет структуру:

```
Строка N:   "Выполненные работы"
Строка N+1: (пустая строка)
Строка N+2: "№ | № кат. | Наименование | Кол. оп. | Цена | Норма | н/ч | Всего | в т.ч. НДС"
Строка N+3: "1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9"
Строка N+4: данные
```

Ожидаемое поведение:
- `dataStartRow = N + 1` (строка после "Выполненные работы")
- Цикл пропускает пустую строку N+1 → `dataStartRow = N + 2`
- `dataStartRow = dataStartRow + 2 = N + 4` (пропуск 2 строк заголовка)
- Данные читаются со строки N+4 ✓

Фактическое поведение (с ошибкой):
- `dataStartRow = N + 1` (строка после "Выполненные работы")
- Цикл НЕ выполняется (srcLastRow = 0)
- `dataStartRow = dataStartRow + 2 = N + 3` (пропущена только 1 строка заголовка!)
- Данные читаются со строки N+3 — это строка подзаголовка "1 | 2 | 3 | ..." ✗

## Исправление

Переместить присвоение `srcLastRow` (строка 110) перед циклом пропуска пустых строк (строки 101-104).

### Файл: `src/modules/Mod_Import.bas`

#### Изменение 1: Таблица "Выполненные работы" (блок строк 97-127)

**Было (строки 97-127):**
```vba
If Not cell Is Nothing Then
    foundWorks = True
    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow
        If Trim(wsSource.Cells(dataStartRow, 4).Value) <> "" Then Exit Do
        dataStartRow = dataStartRow + 1
    Loop
    ' Пропускаем 2 строки заголовка (заголовки колонок + номера столбцов)
    dataStartRow = dataStartRow + 2

    ' Определяем последнюю строку таблицы работ:
    ' ищем первую пустую строку после dataStartRow в столбце D (Наименование)
    srcLastRow = wsSource.Cells(wsSource.Rows.count, 4).End(xlUp).Row
```

**Стало:**
```vba
If Not cell Is Nothing Then
    foundWorks = True

    ' Определяем последнюю строку таблицы работ
    srcLastRow = wsSource.Cells(wsSource.Rows.count, 4).End(xlUp).Row

    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow
        If Trim(wsSource.Cells(dataStartRow, 4).Value) <> "" Then Exit Do
        dataStartRow = dataStartRow + 1
    Loop
    ' Пропускаем 2 строки заголовка (заголовки колонок + номера столбцов)
    dataStartRow = dataStartRow + 2
```

#### Изменение 2: Таблица "Расходная накладная" (блок строк 141-171)

Аналогичное исправление для консистентности, хотя ошибка там не проявляется из-за того, что `srcLastRow` уже инициализирован.

**Было (строки 141-159):**
```vba
If Not cell Is Nothing Then
    foundMaterials = True
    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow
        If Trim(wsSource.Cells(dataStartRow, 2).Value) <> "" Then Exit Do
        dataStartRow = dataStartRow + 1
    Loop
    ' Пропускаем 2 строки заголовка (заголовки колонок + номера столбцов)
    dataStartRow = dataStartRow + 2

    ' Определяем последнюю строку таблицы материалов
    srcLastRow = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).Row
```

**Стало:**
```vba
If Not cell Is Nothing Then
    foundMaterials = True

    ' Определяем последнюю строку таблицы материалов
    srcLastRow = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).Row

    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow
        If Trim(wsSource.Cells(dataStartRow, 2).Value) <> "" Then Exit Do
        dataStartRow = dataStartRow + 1
    Loop
    ' Пропускаем 2 строки заголовка (заголовки колонок + номера столбцов)
    dataStartRow = dataStartRow + 2
```

## Дополнительная рекомендация

Проверить соответствие маппинга колонок для таблицы "Расходная накладная". В комментариях (строка 132) указана структура:
```
A=№, B=№ кат., C=Наименование, D=Кол-во, E=Ед.изм., F=Цена, G=Всего, H=в т.ч. НДС
```

Однако в эталонном выводе (`sheet_main.md`) заголовок таблицы материалов:
```
№ п/п | Артикул | Наименование | Ед. изм. | Кол-во | Цена | Сумма
```

Это может означать, что:
- B = Артикул (не № кат.)
- D = Ед. изм. (не Кол-во)
- E = Кол-во (не Ед.изм.)
- G = Сумма (не Всего)

Текущий маппинг `D(4)→Z(26)` может быть неверным — должно быть `E(5)→Z(26)`. Однако это не является причиной текущей ошибки сдвига, и требует отдельной верификации на реальных данных.