Attribute VB_Name = "Mod_OrderHeader"
Option Explicit

' ============================================================================
' Модуль: Mod_OrderHeader
' Назначение: Автоматическое заполнение шапки заказ-наряда по номеру заказа
' Автор: SourceCraft
' Дата: 2026-07-09
' ============================================================================

' --------------------------------------------------------------------------
' Главная функция FillHeaderFromOrder
' Заполняет ячейки B3:B15 на листе main из данных заказа (B2).
'
' Параметры:
'   orderNumber - номер заказа (значение из ячейки B2)
'   wsMain      - объект листа "main"
'   wsSpisok    - объект листа "spisok" (справочник заказов)
'   wsModel     - объект листа "model" (справочник моделей)
'
' Возврат:
'   True  - если выполнение прошло без ошибок VBA
'   False - если произошла runtime-ошибка
' --------------------------------------------------------------------------
Public Function FillHeaderFromOrder( _
    ByVal orderNumber As Variant, _
    ByVal wsMain As Worksheet, _
    ByVal wsSpisok As Worksheet, _
    ByVal wsModel As Worksheet _
) As Boolean

    ' Константы для листа "spisok"
    Const SPISOK_COL_ORDER As Long = 1   ' Колонка A - номер заказа
    Const SPISOK_COL_MODEL As Long = 2   ' Колонка B - модель
    Const SPISOK_COL_GRZ  As Long = 3   ' Колонка C - ГРЗ
    Const SPISOK_COL_VIN  As Long = 4   ' Колонка D - VIN
    Const SPISOK_COL_GARAGE As Long = 5 ' Колонка E - Гараж. N
    Const SPISOK_COL_YEAR As Long = 6   ' Колонка F - Год вып.
    Const SPISOK_COL_MILEAGE As Long = 7 ' Колонка G - Пробег
    Const SPISOK_COL_DATE As Long = 8   ' Колонка H - Дата
    Const SPISOK_HEADER_ROWS As Long = 1 ' Число строк заголовка

    ' Константы для листа "model"
    Const MODEL_COL_NAME As Long = 1     ' Колонка 1 - наименование модели
    Const MODEL_COL_WORK_ORIG As Long = 3 ' Колонка 3 - работа ориг
    Const MODEL_COL_WORK_MOD As Long = 4  ' Колонка 4 - работа мод
    Const MODEL_COL_PARTS_MOD As Long = 5 ' Колонка 5 - з/ч мод
    Const MODEL_COL_PRICE_NH As Long = 6  ' Колонка 6 - Цена з/ч

    ' Константы для листа "main" (целевые ячейки)
    Const MAIN_CELL_ORDER As String = "B2"   ' Номер заказа (вход)
    Const MAIN_CELL_MODEL As String = "B3"   ' Модель
    Const MAIN_CELL_GRZ As String = "B4"     ' ГРЗ
    Const MAIN_CELL_VIN As String = "B5"     ' VIN
    Const MAIN_CELL_GARAGE As String = "B6"  ' Гараж. N
    Const MAIN_CELL_YEAR As String = "B7"    ' Год вып.
    Const MAIN_CELL_MILEAGE As String = "B8" ' Пробег
    Const MAIN_CELL_DATE As String = "B9"    ' Дата
    Const MAIN_CELL_ORDER_FULL As String = "B10" ' "00" & B2 & "-20"
    Const MAIN_CELL_PRICE_NH As String = "B11"   ' Цена з/ч
    Const MAIN_CELL_WORK_ORIG As String = "B12"  ' Работа ориг
    Const MAIN_CELL_WORK_MOD As String = "B13"   ' Работа мод
    Const MAIN_CELL_PARTS_MOD As String = "B14"  ' з/ч мод
    Const MAIN_CELL_ROW_NUM As String = "B15"    ' Номер строки из spisok - 1

    ' Переменные
    Dim foundRowSpisok As Range   ' Найденная строка на листе spisok
    Dim spisokRowIndex As Long    ' Индекс строки на spisok (найденной)
    Dim orderNumStr As String     ' Строковое представление номера заказа
    Dim orderFullValue As String  ' Значение для B10

    ' Включить обработку ошибок
    On Error GoTo ErrHandler

    ' --- Очистка целевых ячеек перед заполнением ---
    wsMain.Range(MAIN_CELL_MODEL & ":" & MAIN_CELL_ROW_NUM).ClearContents

    ' --- Проверка входного значения ---
    ' Если номер заказа пустой или не число - выходим без ошибки
    If IsEmpty(orderNumber) Then
        FillHeaderFromOrder = True
        Exit Function
    End If

    If Not IsNumeric(orderNumber) Then
        MsgBox "Номер заказа должен быть числовым значением." & vbCrLf & _
               "Пожалуйста, введите корректное число.", _
               vbExclamation, "Неверный номер заказа"
        FillHeaderFromOrder = True
        Exit Function
    End If

    ' Проверка, что число целое и положительное
    ' Используем CDec для сохранения точности для номеров > 15 знаков
    Dim orderNumDec As Variant
    orderNumDec = CDec(orderNumber)

    If orderNumDec < 1 Or orderNumDec <> Int(orderNumDec) Then
        MsgBox "Номер заказа должен быть целым положительным числом." & vbCrLf & _
               "Пожалуйста, проверьте номер заказа.", _
               vbExclamation, "Неверный номер заказа"
        FillHeaderFromOrder = True
        Exit Function
    End If

    ' --- Поиск заказа на листе spisok ---
    ' Используем CStr для поиска по строке, чтобы сохранить ведущие нули
    Set foundRowSpisok = wsSpisok.Columns(SPISOK_COL_ORDER).Find( _
        What:=CStr(orderNumber), _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If foundRowSpisok Is Nothing Then
        MsgBox "Заказ с номером " & orderNumber & " не найден в справочнике заказов." & vbCrLf & _
               "Проверьте, введен ли номер заказа или добавьте новый заказ.", _
               vbInformation, "Заказ не найден"
        FillHeaderFromOrder = False
        GoTo CleanUp
    End If

    ' --- Заполнение данных из spisok ---
    spisokRowIndex = foundRowSpisok.Row

    wsMain.Range(MAIN_CELL_MODEL).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_MODEL).Value
    wsMain.Range(MAIN_CELL_GRZ).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_GRZ).Value
    wsMain.Range(MAIN_CELL_VIN).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_VIN).Value
    wsMain.Range(MAIN_CELL_GARAGE).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_GARAGE).Value
    wsMain.Range(MAIN_CELL_YEAR).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_YEAR).Value
    wsMain.Range(MAIN_CELL_MILEAGE).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_MILEAGE).Value
    wsMain.Range(MAIN_CELL_DATE).Value = _
        wsSpisok.Cells(spisokRowIndex, SPISOK_COL_DATE).Value

    ' --- Формирование значения B10: "00" & B2 & "-20" ---
    orderNumStr = CStr(orderNumber)
    orderFullValue = "00" & orderNumStr & "-20"
    wsMain.Range(MAIN_CELL_ORDER_FULL).Value = orderFullValue

    ' --- B15 = номер строки из spisok минус 1 (т.е. искомая строка - заголовок) ---
    wsMain.Range(MAIN_CELL_ROW_NUM).Value = spisokRowIndex - SPISOK_HEADER_ROWS

    ' --- Поиск модели на листе model ---
    Dim modelName As String
    modelName = Trim(CStr(wsMain.Range(MAIN_CELL_MODEL).Value))

    If Len(modelName) > 0 Then
        ' Ищем модель в списке моделей, чтобы найти расценки и нормо-часы.
        ' Проходим по всем строкам колонки 1 листа model, используя Trim и сравнение
        ' значений через StrComp (vbTextCompare - для без учета регистра).
        Dim modelRowIndex As Long
        Dim lastRowModel As Long
        Dim i As Long
        Dim cellVal As String
        modelRowIndex = 0

        lastRowModel = wsModel.Cells(wsModel.Rows.Count, MODEL_COL_NAME).End(xlUp).Row

        For i = 1 To lastRowModel
            cellVal = Trim(CStr(wsModel.Cells(i, MODEL_COL_NAME).Value))
            If StrComp(cellVal, modelName, vbTextCompare) = 0 Then
                modelRowIndex = i
                Exit For
            End If
        Next i

        If modelRowIndex = 0 Then
            MsgBox "Модель """ & modelName & """ не найдена в справочнике моделей." & vbCrLf & _
                   "Пожалуйста, добавьте модель на лист model в справочник.", _
                   vbInformation, "Модель не найдена"
            ' B11:B14 остаются пустыми - выходим без ошибки
            FillHeaderFromOrder = False
            GoTo CleanUp
        End If

        ' --- Заполнение данных из model ---
        wsMain.Range(MAIN_CELL_PRICE_NH).Value = _
            wsModel.Cells(modelRowIndex, MODEL_COL_PRICE_NH).Value
        wsMain.Range(MAIN_CELL_WORK_ORIG).Value = _
            wsModel.Cells(modelRowIndex, MODEL_COL_WORK_ORIG).Value
        wsMain.Range(MAIN_CELL_WORK_MOD).Value = _
            wsModel.Cells(modelRowIndex, MODEL_COL_WORK_MOD).Value
        wsMain.Range(MAIN_CELL_PARTS_MOD).Value = _
            wsModel.Cells(modelRowIndex, MODEL_COL_PARTS_MOD).Value
    End If

    ' --- Успешное завершение ---
    FillHeaderFromOrder = True
    GoTo CleanUp

CleanUp:
    ' Освобождение объекта Range
    If Not foundRowSpisok Is Nothing Then Set foundRowSpisok = Nothing

    ' Отключаем обработчик ошибок, чтобы не перехватить ошибку в точке выхода
    On Error GoTo 0

    Exit Function

ErrHandler:
    ' Обработка runtime-ошибки
    MsgBox "Возникла необработанная ошибка при заполнении шапки заказ-наряда:" & vbCrLf & _
           vbCrLf & _
           "Ошибка: " & Err.Description & vbCrLf & _
           "Номер: " & Err.Number, _
           vbCritical, "Ошибка выполнения"
    FillHeaderFromOrder = False
    Resume CleanUp

End Function