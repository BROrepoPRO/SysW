# План: Кнопочные обработчики для проекта SysW

## 1. Сводка по модулям

### 1.1. Mod_Utils.bas — Вспомогательные утилиты

| Процедура/Функция | Тип | Параметры | Описание |
|---|---|---|---|
| `GetSheetByName` | Function | `wb As Workbook, SheetName As String` → `Worksheet` | Получение листа по имени без ошибки |
| `GetWorkbookPath` | Function | нет → `String` | Путь к текущей книге |
| `FileExists` | Function | `FilePath As String` → `Boolean` | Проверка существования файла |
| `GetCurrentUser` | Function | нет → `String` | Имя пользователя Windows |
| `FormatDateSQL` | Function | `d As Date` → `String` | Форматирование даты в SQLite-формат |
| `WriteLog` | Sub | `Message As String` | Запись в лог-файл |

### 1.2. Mod_OrderHeader.bas — Заполнение шапки заказа

| Процедура/Функция | Тип | Параметры | Описание |
|---|---|---|---|
| `FillHeaderFromOrder` | Function | `OrderNum As Variant` → `Boolean` | Заполняет B3:B15 на листе main данными из spisok и model по номеру заказа |
| `FindOrder` | Function | `OrderNum As String, ByRef Header As OrderHeader` → `Boolean` | Поиск заказа по номеру, заполнение структуры OrderHeader |

### 1.3. Mod_Import.bas — Импорт данных

| Процедура/Функция | Тип | Параметры | Описание |
|---|---|---|---|
| `ExtractNumberFromGRZ` | Function | `GRZ As String` → `String` | Извлечение цифровой группы (3-4 цифры) из ГРЗ |
| `SearchSheetByGRZ` | Function | `GRZ As String` → `Worksheet` | Поиск листа в report.xlsx по номеру ГРЗ |
| `RenameSheetsByGRZ` | Sub | нет | Переименование листов в report.xlsx по ГРЗ |
| `ImportSheet` | Sub | `GRZ As String` | Импорт листа из report.xlsx по ГРЗ в текущую книгу |
| `ImportDataToMain` | Sub | `wsSource As Worksheet` | Перенос данных из листа-источника в лист main |

### 1.4. Mod_FullTestRunner.bas — Тестирование

| Процедура/Функция | Тип | Параметры | Описание |
|---|---|---|---|
| `RunAllTests` | Sub | нет | Запуск всех тестов (TC-01..TC-20) |

### 1.5. Mod_ButtonDispatcher.bas — Диспетчер кнопок (текущий)

| Процедура/Функция | Тип | Параметры | Описание |
|---|---|---|---|
| `Btn_main_Clear_Click` | Sub | нет | Очистка всех данных на листе main |
| `Btn_main_Import_Click` | Sub | нет | Импорт из отчёта по ГРЗ из ячейки B4 |

### 1.6. Sheet1_main.cls — Класс листа main

| Событие | Описание |
|---|---|
| `Worksheet_Change` | Отслеживает изменение ячейки B2, вызывает `FillHeaderFromOrder` |

---

## 2. Таблица существующих обработчиков

| Имя обработчика | Модуль | Вызываемая процедура | Лист | Статус |
|---|---|---|---|---|
| `Btn_main_Clear_Click` | `Mod_ButtonDispatcher` | Прямая реализация (очистка листа main) | main | ✅ Работает |
| `Btn_main_Import_Click` | `Mod_ButtonDispatcher` | `Mod_Import.ImportSheet(Me.Range("B4").Value)` | main | ✅ Работает |

**Всего существующих: 2**

---

## 3. Таблица требуемых обработчиков

### 3.1. Обработчики для листа main (основной лист)

| № | Имя обработчика | Вызываемая процедура | Параметры | Источник данных | Описание |
|---|---|---|---|---|---|
| 1 | `Btn_main_FillHeader_Click` | `Mod_OrderHeader.FillHeaderFromOrder` | `OrderNum` из `Me.Range("B2").Value` | Ячейка B2 | Заполнить шапку заказа по номеру из B2 |
| 2 | `Btn_main_ClearHeader_Click` | Прямая реализация | нет | — | Очистить только B3:B15 (шапку заказа) |
| 3 | `Btn_main_ImportByB4_Click` | `Mod_Import.ImportSheet` | `GRZ` из `Me.Range("B4").Value` | Ячейка B4 | Импорт по ГРЗ из ячейки B4 (уже есть как `Btn_main_Import_Click`) |
| 4 | `Btn_main_ImportByInput_Click` | `Mod_Import.ImportSheet` | `GRZ` из `InputBox` | Диалог ввода | Импорт по ГРЗ, введённому пользователем |
| 5 | `Btn_main_RunTests_Click` | `Mod_FullTestRunner.RunAllTests` | нет | — | Запустить все тесты |
| 6 | `Btn_main_WriteLog_Click` | `Mod_Utils.WriteLog` | `Message` из `InputBox` | Диалог ввода | Записать сообщение в лог |

### 3.2. Обработчики для листа import (если такой лист существует) или на main

| № | Имя обработчика | Вызываемая процедура | Параметры | Источник данных | Описание |
|---|---|---|---|---|---|
| 7 | `Btn_main_RenameSheets_Click` | `Mod_Import.RenameSheetsByGRZ` | нет | — | Переименовать листы в report.xlsx по ГРЗ |
| 8 | `Btn_main_ImportDataToMain_Click` | `Mod_Import.ImportDataToMain` | `wsSource` — активный лист | Активный лист | Импортировать данные с активного листа в main |

### 3.3. Обработчики для отладки/администрирования

| № | Имя обработчика | Вызываемая процедура | Параметры | Источник данных | Описание |
|---|---|---|---|---|---|
| 9 | `Btn_main_ShowWorkbookPath_Click` | `Mod_Utils.GetWorkbookPath` | нет | — | Показать путь к книге в MsgBox |
| 10 | `Btn_main_ShowCurrentUser_Click` | `Mod_Utils.GetCurrentUser` | нет | — | Показать имя текущего пользователя |
| 11 | `Btn_main_CheckFileExists_Click` | `Mod_Utils.FileExists` | `FilePath` из `InputBox` | Диалог ввода | Проверить существование файла |

### 3.4. Обработчики для поиска заказа

| № | Имя обработчика | Вызываемая процедура | Параметры | Источник данных | Описание |
|---|---|---|---|---|---|
| 12 | `Btn_main_FindOrder_Click` | `Mod_OrderHeader.FindOrder` | `OrderNum` из `InputBox` | Диалог ввода | Найти заказ по номеру и показать результат |

**Всего предлагаемых новых обработчиков: 12**

---

## 4. Архитектурные рекомендации

### 4.1. Структура модулей

**Рекомендация:** Все обработчики кнопок хранить в одном модуле `Mod_ButtonDispatcher.bas`.

**Обоснование:**
- Все обработчики имеют единый шаблон именования `Btn_<Лист>_<Действие>_Click`
- Единая точка входа для всех кнопок упрощает поддержку
- Модуль уже содержит 2 обработчика — расширение логично
- Разделение на несколько модулей диспетчеров создаст путаницу

**Альтернатива (если обработчиков станет > 20):** Создать отдельные модули по группам:
- `Mod_ButtonDispatcher_Import.bas` — для импорта
- `Mod_ButtonDispatcher_Order.bas` — для заказов
- `Mod_ButtonDispatcher_Utils.bas` — для утилит

### 4.2. Базовая обработка ошибок

Предлагается единый шаблон для всех обработчиков:

```vba
Public Sub Btn_main_<Действие>_Click()
    On Error GoTo ErrHandler

    ' ... логика обработчика ...

    Exit Sub

ErrHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Btn_main_<Действие>_Click: " & Err.Description)
End Sub
```

**Компоненты обработки ошибок:**
1. `On Error GoTo ErrHandler` — перехват ошибок
2. `MsgBox` — уведомление пользователя
3. `Mod_Utils.WriteLog` — запись в лог
4. `Exit Sub` перед меткой `ErrHandler` — предотвращение случайного выполнения

### 4.3. Группировка обработчиков

Внутри `Mod_ButtonDispatcher.bas` обработчики группируются по листам с помощью комментариев-разделителей:

```vba
' ============================================================
' ОБРАБОТЧИКИ ЛИСТА MAIN
' ============================================================

' ... все Btn_main_* ...

' ============================================================
' ОБРАБОТЧИКИ ЛИСТА IMPORT (если потребуется)
' ============================================================

' ... все Btn_import_* ...
```

### 4.4. Взаимодействие с листами

- Для получения значений с листа использовать `Me.Range("...").Value` (при вызове из класса листа) или `ThisWorkbook.Sheets("main").Range("...").Value`
- Для диалогового ввода — `InputBox` с проверкой на пустое значение (Cancel)
- Для подтверждения — `MsgBox` с `vbYesNo`

### 4.5. Расхождения в коде (важно!)

В ходе анализа обнаружены расхождения между модулями и тестами:

| Расхождение | Где | Описание |
|---|---|---|
| `FillHeaderFromOrder` сигнатура | `Mod_OrderHeader.bas` vs `Mod_FullTestRunner.bas` | В модуле: 1 параметр (`OrderNum`). В тестах: 4 параметра (`OrderNum, wsMain, wsSpisok, wsModel`). **Нужна синхронизация.** |
| `ImportFromReport` | `Mod_FullTestRunner.bas` | Тест TC-17 вызывает `ImportFromReport`, но такой процедуры нет ни в одном модуле. **Возможно, устаревшее название или缺失ющая процедура.** |

---

## 5. Предложение по обновлению .ycarules

Поскольку `.ycarules` — бинарный файл, его содержимое нечитаемо. Рекомендуется:

1. **Проверить текущие правила** через интерфейс Roo Code (если это файл конфигурации ассистента).
2. **Добавить правила для новых обработчиков**, если `.ycarules` управляет видимостью/доступом процедур.
3. **Если `.ycarules` не относится к VBA**, а является конфигурацией IDE/ассистента — создать текстовый файл `.ycarules.md` с пояснениями.

---

## 6. Итоговая таблица: все публичные процедуры и их обработчики

| Модуль | Процедура/Функция | Существует обработчик? | Предлагаемый обработчик |
|---|---|---|---|
| `Mod_Utils` | `GetSheetByName` | ❌ Нет (вспомогательная, не для кнопки) | — |
| `Mod_Utils` | `GetWorkbookPath` | ❌ Нет | `Btn_main_ShowWorkbookPath_Click` |
| `Mod_Utils` | `FileExists` | ❌ Нет | `Btn_main_CheckFileExists_Click` |
| `Mod_Utils` | `GetCurrentUser` | ❌ Нет | `Btn_main_ShowCurrentUser_Click` |
| `Mod_Utils` | `FormatDateSQL` | ❌ Нет (вспомогательная, не для кнопки) | — |
| `Mod_Utils` | `WriteLog` | ❌ Нет | `Btn_main_WriteLog_Click` |
| `Mod_OrderHeader` | `FillHeaderFromOrder` | ❌ Нет (есть автовызов из `Worksheet_Change`) | `Btn_main_FillHeader_Click` |
| `Mod_OrderHeader` | `FindOrder` | ❌ Нет | `Btn_main_FindOrder_Click` |
| `Mod_Import` | `ExtractNumberFromGRZ` | ❌ Нет (вспомогательная, не для кнопки) | — |
| `Mod_Import` | `SearchSheetByGRZ` | ❌ Нет (вспомогательная, не для кнопки) | — |
| `Mod_Import` | `RenameSheetsByGRZ` | ❌ Нет | `Btn_main_RenameSheets_Click` |
| `Mod_Import` | `ImportSheet` | ✅ `Btn_main_Import_Click` | — |
| `Mod_Import` | `ImportDataToMain` | ❌ Нет | `Btn_main_ImportDataToMain_Click` |
| `Mod_FullTestRunner` | `RunAllTests` | ❌ Нет | `Btn_main_RunTests_Click` |
| `Mod_ButtonDispatcher` | `Btn_main_Clear_Click` | ✅ Самостоятельный | — |
| `Mod_ButtonDispatcher` | `Btn_main_Import_Click` | ✅ Самостоятельный | — |

---

## 7. Очерёдность реализации

| Приоритет | Обработчик | Обоснование |
|---|---|---|
| 🔴 P0 | `Btn_main_FillHeader_Click` | Основная функция — заполнение шапки заказа (сейчас только автовызов) |
| 🔴 P0 | `Btn_main_RunTests_Click` | Необходим для тестирования |
| 🟡 P1 | `Btn_main_ClearHeader_Click` | Очистка только шапки (без очистки всего листа) |
| 🟡 P1 | `Btn_main_ImportByInput_Click` | Импорт по введённому ГРЗ |
| 🟡 P1 | `Btn_main_RenameSheets_Click` | Переименование листов в отчёте |
| 🟢 P2 | `Btn_main_ImportDataToMain_Click` | Импорт с активного листа |
| 🟢 P2 | `Btn_main_FindOrder_Click` | Поиск заказа |
| 🔵 P3 | `Btn_main_ShowWorkbookPath_Click` | Администрирование |
| 🔵 P3 | `Btn_main_ShowCurrentUser_Click` | Администрирование |
| 🔵 P3 | `Btn_main_CheckFileExists_Click` | Администрирование |
| 🔵 P3 | `Btn_main_WriteLog_Click` | Администрирование |

---

## 8. Пример реализации нового обработчика

```vba
' --------------------------------------------------------------------------
' Btn_main_FillHeader_Click
' Заполняет шапку заказа (B3:B15) по номеру из ячейки B2
' --------------------------------------------------------------------------
Public Sub Btn_main_FillHeader_Click()
    On Error GoTo ErrHandler

    Dim orderNum As Variant
    orderNum = ThisWorkbook.Sheets("main").Range("B2").Value

    If IsEmpty(orderNum) Or orderNum = "" Then
        MsgBox "Введите номер заказа в ячейку B2!", vbExclamation, "Предупреждение"
        Exit Sub
    End If

    If Not Mod_OrderHeader.FillHeaderFromOrder(orderNum) Then
        MsgBox "Не удалось заполнить шапку заказа.", vbExclamation, "Ошибка"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в Btn_main_FillHeader_Click: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Btn_main_FillHeader_Click: " & Err.Description)
End Sub