# План исправления Mod_Import.bas

## Диагноз

**Ошибка:** Неинициализированная переменная `srcLastRow` в процедуре `ImportDataToMain`.

Переменная `Dim srcLastRow As Long` (строка 62) инициализируется значением `0`. Цикл пропуска пустых строк (строки 101-104) использует `srcLastRow` как верхнюю границу, но на момент выполнения цикла `srcLastRow = 0`. Поскольку `dataStartRow = cell.Row + 1 > 0`, условие `dataStartRow <= 0` ложно, и цикл никогда не выполняется.

**Следствие:** Пустые строки между заголовком "Выполненные работы" и шапкой таблицы не пропускаются. Данные читаются со строки подзаголовка ("1 | 2 | 3 | ...") вместо строки данных, что приводит к сдвигу на 1 строку вверх.

## Изменения

### Файл: `src/modules/Mod_Import.bas`

#### Изменение 1: Таблица "Выполненные работы" (блок строк 96-148)

Переместить присвоение `srcLastRow` перед циклом пропуска пустых строк.

**Было:**
```vba
If Not cell Is Nothing Then
    foundWorks = True
    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow   ' srcLastRow = 0 → цикл не выполняется!
        ...
    Loop
    ' Пропускаем две строки заголовка
    dataStartRow = dataStartRow + 2
    ' Определяем последнюю строку таблицы работ
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
        ...
    Loop
    ' Пропускаем две строки заголовка
    dataStartRow = dataStartRow + 2
```

#### Изменение 2: Таблица "Расходная накладная" (блок строк 160-205)

Аналогичное исправление для консистентности (ошибка там не проявляется, т.к. `srcLastRow` уже инициализирован к моменту обработки этой таблицы).

**Было:**
```vba
If Not cell Is Nothing Then
    foundMaterials = True
    ' Пропускаем пустые строки после названия таблицы
    dataStartRow = cell.Row + 1
    Do While dataStartRow <= srcLastRow
        ...
    Loop
    ' Пропускаем две строки заголовка
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
        ...
    Loop
    ' Пропускаем две строки заголовка
    dataStartRow = dataStartRow + 2
```

## Проверка

1. После исправления — импортировать модули в `work.xlsm`: `python scripts/impVBA.py`
2. Запустить тесты: `python scripts/run_tests.py`
3. Сравнить результат импорта с эталоном `base/exam-temp/sheet_main.md`