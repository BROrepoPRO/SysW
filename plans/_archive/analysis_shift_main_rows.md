# Анализ жёстких ссылок на строки листа main

> **Дата:** 2026-07-21
> **Задача:** Сдвиг структуры листа `main` на 2 строки вниз (строки 1–2 резервируются под будущую логику)
> **Принцип:** Все старые номера строк → старый номер + 2

---

## 1. Текущая структура листа main (до сдвига)

Из [`plans/_archive/structure_current_work.xlsm.md`](plans/_archive/structure_current_work.xlsm.md):

### Шапка заказа (столбец B, строки 2–15)

| Ячейка | Поле | Назначение |
|--------|------|-----------|
| B2 | Номер заказа | Ввод пользователя |
| B3 | № заказа (формат "00{num}-20") | Авто из B2 |
| B4 | Модель | Из spisok (кол. B) |
| B5 | ГРЗ | Из spisok (кол. C) |
| B6 | VIN | Из spisok (кол. D) |
| B7 | Гараж. № | Из spisok (кол. E) |
| B8 | Год выпуска | Из spisok (кол. F) |
| B9 | Пробег | Из spisok (кол. G) |
| B10 | Дата | Из spisok (кол. H) |
| B11 | Цена н/ч | Из model (кол. C) |
| B12 | Группа | Из model (кол. B) |
| B13 | (заглушка) | — |
| B14 | (заглушка) | — |
| B15 | (заглушка) | — |

### Таблица работ (колонки L–N) — данные со строки 2
### Таблица материалов (колонки X–AA) — данные со строки 2

---

## 2. Новая структура листа main (после сдвига)

| Ячейка | Поле |
|--------|------|
| **B1** | **(пусто — резерв)** |
| **B2** | **(пусто — резерв)** |
| **B3** | **(бывш. B2)** Номер заказа |
| **B4** | **(бывш. B3)** № заказа (формат "00{num}-20") |
| **B5** | **(бывш. B4)** Модель |
| **B6** | **(бывш. B5)** ГРЗ |
| **B7** | **(бывш. B6)** VIN |
| **B8** | **(бывш. B7)** Гараж. № |
| **B9** | **(бывш. B8)** Год выпуска |
| **B10** | **(бывш. B9)** Пробег |
| **B11** | **(бывш. B10)** Дата |
| **B12** | **(бывш. B11)** Цена н/ч |
| **B13** | **(бывш. B12)** Группа |
| **B14** | **(бывш. B13)** заглушка |
| **B15** | **(бывш. B14)** заглушка |
| **B16** | **(бывш. B15)** заглушка |
| **L4:N...** | **(бывш. L2:N...)** Таблица работ |
| **X4:AA...** | **(бывш. X2:AA...)** Таблица материалов |

---

## 3. Полный перечень жёстких ссылок по файлам

### 3.1. [`src/modules/Mod_OrderHeader.bas`](src/modules/Mod_OrderHeader.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 27 | `' Заполняет B3:B15 на листе main данными` | B3:B15 | B5:B17 | Комментарий |
| 64 | `wsMain.Range("B3:B15").ClearContents` | B3:B15 | B5:B17 | Очистка шапки |
| 77 | `wsMain.Range("B3:B15").ClearContents` | B3:B15 | B5:B17 | Очистка при ошибке |
| 83 | `wsMain.Cells(3, 2).Value = ...` | строка 3 (B3) | Cells(5, 2) | Модель |
| 84 | `wsMain.Cells(4, 2).Value = ...` | строка 4 (B4) | Cells(6, 2) | ГРЗ |
| 85 | `wsMain.Cells(5, 2).Value = ...` | строка 5 (B5) | Cells(7, 2) | VIN |
| 86 | `wsMain.Cells(6, 2).Value = ...` | строка 6 (B6) | Cells(8, 2) | Гараж.№ |
| 87 | `wsMain.Cells(7, 2).Value = ...` | строка 7 (B7) | Cells(9, 2) | Год вып. |
| 88 | `wsMain.Cells(8, 2).Value = ...` | строка 8 (B8) | Cells(10, 2) | Пробег |
| 89 | `wsMain.Cells(9, 2).Value = ...` | строка 9 (B9) | Cells(11, 2) | Дата |
| 92 | `wsMain.Cells(10, 2).Value = "00" & ...` | строка 10 (B10) | Cells(12, 2) | № ЗН |
| 95 | `ModelCode = Trim(wsMain.Cells(3, 2).Value)` | строка 3 (B3) | Cells(5, 2) | Чтение модели |
| 116 | `wsMain.Cells(11, 2).Value = ...` | строка 11 (B11) | Cells(13, 2) | Цена н/ч |
| 117 | `wsMain.Cells(12, 2).Value = ...` | строка 12 (B12) | Cells(14, 2) | Группа |
| 203 | `orderNum = ...Range("B2").Value` | B2 | B4 | FillHeaderFromOrder_UI |
| 206 | `MsgBox ... "ячейку B2"` | B2 | B4 | Текст сообщения |

### 3.2. [`src/modules/Mod_Import.bas`](src/modules/Mod_Import.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 34 | `newName = Trim(wsMain.Range("B2").Value) & "M"` | B2 | B4 | ImportSheet |
| 78 | `If lastRow < 2 Then lastRow = 2` | 2 | 4 | Минимальная строка для очистки |
| 80 | `wsMain.Range("L2:N" & lastRow).ClearContents` | L2:N... | L4:N... | Очистка работ |
| 81 | `wsMain.Range("X2:AA" & lastRow).ClearContents` | X2:AA... | X4:AA... | Очистка материалов |
| 130 | `targetRow = 2` | 2 | 4 | Стартовая строка для работ |
| 186 | `targetRow = 2` | 2 | 4 | Стартовая строка для материалов |
| 236 | `Call ImportSheet(ThisWorkbook.Sheets("main").Range("B4").Value)` | B4 | B6 | ImportSheet_UI (ГРЗ) |
| 341 | `grz = Trim(CStr(wsMain.Range("B2").Value))` | B2 | B4 | ImportFromB2_UI |
| 345 | `MsgBox ... "Ячейка B2 на листе 'main' пуста"` | B2 | B4 | Текст сообщения |

### 3.3. [`src/modules/Mod_SheetOps.bas`](src/modules/Mod_SheetOps.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 182 | `If lastRow >= 2 Then` | 2 | 4 | ClearMainSheet_UI (silent) |
| 183 | `wsMain.Range("B2:ZZ" & lastRow).ClearContents` | B2:ZZ... | B4:ZZ... | Очистка данных |
| 189 | `If lastRow >= 2 Then` | 2 | 4 | ClearMainSheet_UI (с подтверждением) |
| 190 | `wsMain.Range("B2:ZZ" & lastRow).ClearContents` | B2:ZZ... | B4:ZZ... | Очистка данных |
| 208 | `' Очищает шапку заказа B3:B15 на листе main` | B3:B15 | B5:B17 | Комментарий |
| 216 | `wsMain.Range("B3:B15").ClearContents` | B3:B15 | B5:B17 | ClearHeader_UI |
| 218 | `MsgBox "Шапка заказа (B3:B15) очищена."` | B3:B15 | B5:B17 | Текст сообщения |

### 3.4. [`src/modules/Mod_MainButtons.bas`](src/modules/Mod_MainButtons.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 17 | `' Импорт из report.xlsx по ГРЗ из ячейки B4 листа main.` | B4 | B6 | Комментарий |

Прямых жёстких ссылок на строки в коде нет — только комментарий. Кнопки делегируют вызовы в `Mod_Import` и `Mod_SheetOps`.

### 3.5. [`src/modules/Mod_Constants.bas`](src/modules/Mod_Constants.bas)

Жёстких ссылок на строки листа `main` **нет**. Модуль содержит только константы для столбцов листов `spisok` и `models`, а также реестр `libname`.

### 3.6. [`src/modules/Mod_FullTestRunner.bas`](src/modules/Mod_FullTestRunner.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 472 | `oldB2 = Trim(CStr(wsMain.Range("B2").Value))` | B2 | B4 | TC-14: сохранение B2 |
| 475 | `wsMain.Range("B2").Value = ""` | B2 | B4 | TC-14: очистка B2 |
| 488 | `wsMain.Range("B2").Value = oldB2` | B2 | B4 | TC-14: восстановление B2 |

### 3.7. [`src/modules/Mod_Utils.bas`](src/modules/Mod_Utils.bas)

Жёстких ссылок на строки листа `main` **нет**. Модуль содержит только общие утилиты.

### 3.8. [`src/sheets/Лист2_main.cls`](src/sheets/Лист2_main.cls)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 22 | `If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub` | B2 | B4 | Проверка изменения B2 |
| 24 | `If Target.Address = "$B$2" Then` | $B$2 | $B$4 | Адрес ячейки |
| 26 | `If Target.Row >= 3 And Target.Row <= 15 Then Exit Sub` | 3..15 | 5..17 | Защита от рекурсии при очистке |
| 31 | `b2Value = Me.Range("B2").Value` | B2 | B4 | Чтение значения |

### 3.9. [`src/modules/Mod_ButtonDispatcher.bas`](src/modules/Mod_ButtonDispatcher.bas)

| Строка файла | Код | Жёсткая ссылка | Заменить на | Примечание |
|-------------|-----|----------------|-------------|-----------|
| 24 | `' Запускает импорт из отчёта по ГРЗ из ячейки B4` | B4 | B6 | Комментарий |
| 32 | `' Заполняет шапку заказа (B3:B15) по номеру из ячейки B2` | B3:B15, B2 | B5:B17, B4 | Комментарий |
| 40 | `' Очищает только шапку заказа (B3:B15) на листе main` | B3:B15 | B5:B17 | Комментарий |

Прямых жёстких ссылок в исполняемом коде нет — только комментарии. Все вызовы делегируются.

### 3.10. [`src/modules/Mod_SheetButtons.bas`](src/modules/Mod_SheetButtons.bas)

Жёстких ссылок на строки листа `main` **нет**. Модуль работает только с листами `z4` и `work`.

### 3.11. [`src/modules/Mod_Logger.bas`](src/modules/Mod_Logger.bas)

Жёстких ссылок на строки листа `main` **нет**. Модуль не взаимодействует с листом `main`.

---

## 4. Сводная статистика

| Файл | Кол-во изменений | Комментарии | Исполняемый код |
|------|-----------------|-------------|-----------------|
| `Mod_OrderHeader.bas` | 14 | 1 | 13 |
| `Mod_Import.bas` | 8 | 0 | 8 |
| `Mod_SheetOps.bas` | 7 | 1 | 6 |
| `Mod_MainButtons.bas` | 1 | 1 | 0 |
| `Mod_Constants.bas` | 0 | 0 | 0 |
| `Mod_FullTestRunner.bas` | 3 | 0 | 3 |
| `Mod_Utils.bas` | 0 | 0 | 0 |
| `Лист2_main.cls` | 5 | 0 | 5 |
| `Mod_ButtonDispatcher.bas` | 3 | 3 | 0 |
| `Mod_SheetButtons.bas` | 0 | 0 | 0 |
| `Mod_Logger.bas` | 0 | 0 | 0 |
| **Итого** | **41** | **6** | **35** |

---

## 5. Рекомендации по новым константам

**Категорически рекомендуется добавить следующие константы в [`src/modules/Mod_Constants.bas`](src/modules/Mod_Constants.bas):**

```vba
' ============================================================
' Константы строк листа main
' ============================================================
Public Const MAIN_HEADER_START_ROW As Long = 4   ' B4 — номер заказа (ввод пользователя)
Public Const MAIN_HEADER_END_ROW As Long = 17     ' B17 — последняя строка шапки (заглушка)
Public Const MAIN_DATA_START_ROW As Long = 4      ' Строка, с которой начинаются таблицы работ/материалов
Public Const MAIN_ORDER_NUM_CELL As String = "B4" ' Ячейка с номером заказа
Public Const MAIN_GRZ_CELL As String = "B6"       ' Ячейка с ГРЗ (для импорта)
```

**Пояснения:**
- `MAIN_HEADER_START_ROW = 4` — первая строка данных шапки (бывш. B2 → B4)
- `MAIN_HEADER_END_ROW = 17` — последняя строка шапки (бывш. B15 → B17)
- `MAIN_DATA_START_ROW = 4` — стартовая строка для таблиц L:N и X:AA (бывш. 2 → 4)
- `MAIN_ORDER_NUM_CELL = "B4"` — ячейка с номером заказа (бывш. B2)
- `MAIN_GRZ_CELL = "B6"` — ячейка с ГРЗ (бывш. B4)

**Опционально** (для полной константизации шапки):
```vba
Public Const MAIN_HEADER_ORDER_NUM_ROW As Long = 4   ' B4 — № заказа
Public Const MAIN_HEADER_MODEL_ROW As Long = 5       ' B5 — модель
Public Const MAIN_HEADER_GRZ_ROW As Long = 6         ' B6 — ГРЗ
Public Const MAIN_HEADER_VIN_ROW As Long = 7         ' B7 — VIN
Public Const MAIN_HEADER_GARAGE_ROW As Long = 8      ' B8 — Гараж.№
Public Const MAIN_HEADER_YEAR_ROW As Long = 9        ' B9 — Год вып.
Public Const MAIN_HEADER_MILEAGE_ROW As Long = 10    ' B10 — Пробег
Public Const MAIN_HEADER_DATE_ROW As Long = 11       ' B11 — Дата
Public Const MAIN_HEADER_ORDERNO_ROW As Long = 12    ' B12 — № ЗН (00{num}-20)
Public Const MAIN_HEADER_PRICE_ROW As Long = 13      ' B13 — Цена н/ч
Public Const MAIN_HEADER_GROUP_ROW As Long = 14      ' B14 — Группа
```

---

## 6. Риски и предупреждения

1. **Диапазон `B3:B15` встречается в 4 местах** — критически важно заменить все на `B5:B17`.
2. **`targetRow = 2` в `Mod_Import.bas`** — стартовая строка для таблиц работ и материалов. После сдвига должна стать `4`.
3. **`lastRow < 2` в `Mod_Import.bas`** — минимальная строка для очистки. После сдвига — `4`.
4. **`Лист2_main.cls` строка 26** — проверка `Target.Row >= 3 And Target.Row <= 15` защищает от рекурсии. После сдвига — `5..17`.
5. **Комментарии** — не влияют на исполнение, но для консистентности кода их тоже желательно обновить.
6. **Тест TC-14** в `Mod_FullTestRunner.bas` работает с B2 — после сдвига должен работать с B4.
7. **После внесения изменений** необходимо прогнать все тесты (TC-01..TC-14) для верификации.

---

## 7. Порядок внесения изменений (рекомендуемый)

1. Добавить константы в `Mod_Constants.bas`
2. `Mod_OrderHeader.bas` — заменить все `Cells(n, 2)` на `Cells(n+2, 2)` и `Range("B3:B15")` на `Range("B5:B17")`
3. `Mod_Import.bas` — заменить `B2`→`B4`, `B4`→`B6`, `targetRow = 2`→`4`, `L2:N`→`L4:N`, `X2:AA`→`X4:AA`
4. `Mod_SheetOps.bas` — заменить `B2:ZZ`→`B4:ZZ`, `B3:B15`→`B5:B17`
5. `Лист2_main.cls` — заменить `B2`→`B4`, `$B$2`→`$B$4`, диапазон `3..15`→`5..17`
6. `Mod_FullTestRunner.bas` — заменить `B2`→`B4` в TC-14
7. Обновить комментарии в `Mod_MainButtons.bas` и `Mod_ButtonDispatcher.bas`
8. Прогнать тесты