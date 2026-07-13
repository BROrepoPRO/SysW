# План исправления маппинга данных OrderHeader и ошибок импорта VBA

## Проблема 1: Неправильный маппинг данных OrderHeader

В коде `Mod_OrderHeader.bas` маппинг данных с листа `"spisok"` в тип `OrderHeader` и ячейки B3:B15 листа `main` **не соответствует** реальной структуре столбцов листа `"spisok"`.

### Текущее (неправильное) состояние

```vba
' Поиск идёт по колонке A (№ п/п) — ищет номер заказа, которого там нет
Set FoundCell = wsSpisok.Columns("A").Find(What:=OrderNum, ...)

' Маппинг не соответствует заголовкам:
Header.OrderNumber = FoundCell.Value              ' A = № п/п, а не номер заказа
Header.ClientName   = FoundCell.Offset(0, 1).Value ' B = Модель, а не клиент
Header.CarModel     = FoundCell.Offset(0, 2).Value ' C = ГРЗ, а не модель
Header.CarPlate     = FoundCell.Offset(0, 3).Value ' D = VIN, а не госномер
Header.Mileage      = FoundCell.Offset(0, 4).Value ' E = гараж. №, а не пробег
Header.DateIn       = FoundCell.Offset(0, 5).Value ' F = год вып., а не дата заезда
Header.DateOut      = FoundCell.Offset(0, 6).Value ' G = пробег, а не дата выезда
Header.Status       = FoundCell.Offset(0, 7).Value ' H = дата, а не статус
```

### Реальная структура листа `"spisok"`

| Колонка | Заголовок | Пример |
|---|---|---|
| A | № п/п | 1, 2, 3... |
| B | Модель | УАЗ UAZ PATRIOT 3163 |
| C | ГРЗ | О201МК126 |
| D | VIN | XTT316300R1000091 |
| E | гараж. № | 106-81 |
| F | год вып. | 2023 |
| G | пробег | 4116 |
| H | дата | 01.06.2026 |
| I | (код модели) | ? |

## Проблема 2: Ошибка "Attribute VB_Name - name syntax error" при импорте

Файлы `.bas` и `.cls` содержат строки `Attribute VB_Name = "..."`, которые генерируются VBA-редактором при экспорте. При импорте обратно через COM-скрипт `import_all_vba.py` эти строки **удаляются** (функция `strip_all_attribute_lines`). Но если импортировать файлы **вручную** через VBA-редактор (File → Import File) — возникает синтаксическая ошибка.

**Дополнительная проблема:** В файле `Sheet1_main.cls` есть и другие `Attribute`-строки (`VB_GlobalNameSpace`, `VB_Creatable`, `VB_PredeclaredId`, `VB_Exposed`, `VB_Description`, `VB_HelpID`, `VB_TemplateDerived`, `VB_Customizable`). Скрипт удаляет **все** `Attribute`-строки, что может привести к потере важных настроек листа (например, `VB_PredeclaredId = True` необходим для правильной работы событий `Worksheet_Change`).

## Что нужно исправить

### 1. Изменить поиск — искать по № п/п (колонка A)

Пользователь вводит в B2 листа `main` число (например, `1`), которое соответствует **№ п/п** на листе `spisok`. Поиск по колонке A **корректен**, но нужно понимать, что ищется не номер заказа, а порядковый номер записи.

### 2. Исправить маппинг полей OrderHeader

| Колонка | Заголовок | Поле OrderHeader | Ячейка main |
|---|---|---|---|
| A | № п/п | `OrderNumber` (будет содержать № п/п) | B3 — Номер заказ-наряда (формируется как `"00"&B2&"-20"`) |
| B | Модель | `ClientName` → переименовать в `ModelName` | B4 |
| C | ГРЗ | `CarModel` → переименовать в `GRZ` | B5 |
| D | VIN | `CarPlate` → переименовать в `VIN` | B6 |
| E | гараж. № | `Mileage` → переименовать в `GarageNumber` | B7 |
| F | год вып. | `DateIn` → переименовать в `YearMade` | B8 |
| G | пробег | `DateOut` → переименовать в `MileageValue` | B9 |
| H | дата | `Status` → переименовать в `DateValue` | B10 |

### 3. Номер заказ-наряда формировать, а не брать из spisok

Номер заказ-наряда для ячейки B3 формируется по формуле:
```
"00" & main!B2 & "-20"
```
То есть: `"00" & значение_ячейки_B2 & "-20"`

### 4. Переименовать поля OrderHeader

Текущие имена полей вводят в заблуждение. Предлагается переименовать:

```vba
Public Type OrderHeader
    OrderNumber As String    ' было: OrderNumber (оставить, но теперь это № п/п)
    ModelName As String      ' было: ClientName
    GRZ As String            ' было: CarModel
    VIN As String            ' было: CarPlate
    GarageNumber As String   ' было: Mileage (Long → String, т.к. "106-81")
    YearMade As Integer      ' было: DateIn (Date → Integer)
    MileageValue As Long     ' было: DateOut (Date → Long)
    DateValue As Date        ' было: Status (String → Date)
End Type
```

**Важно:** Изменение типа `OrderHeader` затронет все места, где он используется (FindOrder, тесты).

### 5. Исправить скрипт импорта import_all_vba.py

Сейчас скрипт удаляет **все** `Attribute`-строки. Для `.bas`-файлов это нормально (там только `VB_Name`). Для `.cls`-файлов (особенно `Sheet1_main.cls`) нужно **сохранять** некоторые `Attribute`-строки, критически важные для работы листа:

- `Attribute VB_Name = "Sheet1"` — имя класса
- `Attribute VB_PredeclaredId = True` — **обязательно** для работы событий листа
- `Attribute VB_Exposed = True` — доступность из проекта

Остальные (`VB_GlobalNameSpace`, `VB_Creatable`, `VB_Description`, `VB_HelpID`, `VB_TemplateDerived`, `VB_Customizable`) — опциональны.

## План действий

### Шаг 1: Исправить `Mod_Utils.bas` — тип OrderHeader

- [ ] Переименовать поля в соответствии с реальной структурой
- [ ] Скорректировать типы полей (Mileage → GarageNumber As String, DateIn → YearMade As Integer, DateOut → MileageValue As Long, Status → DateValue As Date)

### Шаг 2: Исправить `Mod_OrderHeader.bas` — функцию FindOrder

- [ ] Исправить маппинг полей в `FindOrder` (строки 71-78)
- [ ] Номер заказа (`OrderNumber`) теперь содержит № п/п из колонки A

### Шаг 3: Исправить `Mod_OrderHeader.bas` — процедуру FillHeaderFromOrder

- [ ] Исправить маппинг в `FillHeaderFromOrder` (строки 29-36)
- [ ] В ячейку B3 записывать сформированный номер: `"00" & wsMain.Range("B2").Value & "-20"`
- [ ] Остальные ячейки B4:B10 заполнять из правильных колонок

### Шаг 4: Исправить `Mod_MinimalTestRunner.bas` — тесты

- [ ] TC-08: изменить тест `FindOrder` — искать по № п/п (например, `"1"`), а не по `"ЗЗ-001"`
- [ ] TC-09: оставить как есть (поиск несуществующего)

### Шаг 5: Исправить `import_all_vba.py` — обработка Attribute-строк

- [ ] Изменить функцию `strip_all_attribute_lines` — для `.cls`-файлов сохранять `VB_Name`, `VB_PredeclaredId`, `VB_Exposed`
- [ ] Для `.bas`-файлов удалять все `Attribute`-строки как и раньше
- [ ] Либо альтернатива: удалять `Attribute`-строки только для `.bas`, а для `.cls` не трогать вообще

### Шаг 6: Проверить `Mod_Import.bas` на зависимость от OrderHeader

- [ ] Убедиться, что `Mod_Import.bas` не использует `OrderHeader` напрямую (судя по коду — не использует)

## Схема изменений

```mermaid
flowchart LR
    subgraph "Лист spisok"
        A[Колонка A: № п/п]
        B[Колонка B: Модель]
        C[Колонка C: ГРЗ]
        D[Колонка D: VIN]
        E[Колонка E: гараж. №]
        F[Колонка F: год вып.]
        G[Колонка G: пробег]
        H[Колонка H: дата]
    end

    subgraph "Лист main"
        B2[Ячейка B2: ввод пользователя]
        B3[Ячейка B3: номер ЗН]
        B4[Ячейка B4: Модель]
        B5[Ячейка B5: ГРЗ]
        B6[Ячейка B6: VIN]
        B7[Ячейка B7: гараж. №]
        B8[Ячейка B8: год вып.]
        B9[Ячейка B9: пробег]
        B10[Ячейка B10: дата]
    end

    B2 -->|"00\"&B2&\"-20"| B3
    A -->|OrderNumber| B3
    B -->|ModelName| B4
    C -->|GRZ| B5
    D -->|VIN| B6
    E -->|GarageNumber| B7
    F -->|YearMade| B8
    G -->|MileageValue| B9
    H -->|DateValue| B10
```

## Проверка после исправления

1. Ввести `1` в B2 → B3 должно стать `001-20`, B4 = "УАЗ UAZ PATRIOT 3163", B5 = "О201МК126" и т.д.
2. Ввести `2` в B2 → B3 должно стать `002-20`, B4 = "LADA VESTA GFL110" и т.д.
3. Тест TC-08 должен проходить при поиске `"1"` (первая запись)
4. Тест TC-09 должен проходить при поиске `"999"` (несуществующая запись)
5. `import_all_vba.py` должен импортировать все модули без ошибок
6. После импорта через скрипт, события `Worksheet_Change` на листе main должны работать корректно