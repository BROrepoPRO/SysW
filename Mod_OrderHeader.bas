Attribute VB_Name = "Mod_OrderHeader"
Option Explicit

' ============================================================================
' Модуль: Mod_OrderHeader
' Назначение: Автоматическое заполнение шапки заказ-наряда по номеру заказа
' Автор: SourceCraft
' Дата: 2026-07-09
' ============================================================================

' --------------------------------------------------------------------------
' Публичная функция FillHeaderFromOrder
' Заполняет ячейки B3:B15 на листе main на основе номера заказа (B2).
'
' Параметры:
'   orderNumber - номер заказа (значение из ячейки B2)
'   wsMain      - рабочий лист "main"
'   wsSpisok    - рабочий лист "spisok" (справочник ТС)
'   wsModel     - рабочий лист "model" (справочник моделей)
'
' Возвращает:
'   True  - если выполнение прошло без критических ошибок VBA
'   False - если возникла runtime-ошибка
' --------------------------------------------------------------------------
Public Function FillHeaderFromOrder( _
    ByVal orderNumber As Variant, _
    ByVal wsMain As Worksheet, _
    ByVal wsSpisok As Worksheet, _
    ByVal wsModel As Worksheet _
) As Boolean

    ' Константы для листа "spisok"
    Const SPISOK_COL_ORDER As Long = 1   ' столбец A — номер заказа
    Const SPISOK_COL_MODEL As Long = 2   ' столбец B — Модель
    Const SPISOK_COL_GRZ  As Long = 3   ' столбец C — ГРЗ
    Const SPISOK_COL_VIN  As Long = 4   ' столбец D — VIN
    Const SPISOK_COL_GARAGE As Long = 5 ' столбец E — гараж. №
    Const SPISOK_COL_YEAR As Long = 6   ' столбец F — год вып.
    Const SPISOK_COL_MILEAGE As Long = 7 ' столбец G — пробег
    Const SPISOK_COL_DATE As Long = 8   ' столбец H — дата
    Const SPISOK_HEADER_ROWS As Long = 1 ' одна строка заголовка

    ' Константы для листа "model"
    Const MODEL_COL_NAME As Long = 1     ' столбец 1 — название модели
    Const MODEL_COL_WORK_ORIG As Long = 3 ' столбец 3 — Работы исх
    Const MODEL_COL_WORK_MOD As Long = 4  ' столбец 4 — Работы мод
    Const MODEL_COL_PARTS_MOD As Long = 5 ' столбец 5 — З/ч мод
    Const MODEL_COL_PRICE_NH As Long = 6  ' столбец 6 — цена н/ч

    ' Константы для листа "main" (целевые ячейки)
    Const MAIN_CELL_ORDER As String = "B2"   ' номер заказа (вход)
    Const MAIN_CELL_MODEL As String = "B3"   ' Модель
    Const MAIN_CELL_GRZ As String = "B4"     ' ГРЗ
    Const MAIN_CELL_VIN As String = "B5"     ' VIN
    Const MAIN_CELL_GARAGE As String = "B6"  ' гараж. №
    Const MAIN_CELL_YEAR As String = "B7"    ' год вып.
    Const MAIN_CELL_MILEAGE As String = "B8" ' пробег
    Const MAIN_CELL_DATE As String = "B9"    ' дата
    Const MAIN_CELL_ORDER_FULL As String = "B10" ' "00" & B2 & "-20"
    Const MAIN_CELL_PRICE_NH As String = "B11"   ' цена н/ч
    Const MAIN_CELL_WORK_ORIG As String = "B12"  ' Работы исх
    Const MAIN_CELL_WORK_MOD As String = "B13"   ' Работы мод
    Const MAIN_CELL_PARTS_MOD As String = "B14"  ' З/ч мод
    Const MAIN_CELL_ROW_NUM As String = "B15"    ' номер строки на spisok - 1

    ' Переменные
    Dim foundRowSpisok As Range   ' найденная строка на листе spisok
    Dim spisokRowIndex As Long    ' номер строки на spisok (абсолютный)
    Dim orderNumStr As String     ' строковое представление номера заказа
    Dim orderFullValue As String  ' значение для B10

    ' Включаем обработку ошибок
    On Error GoTo ErrHandler

    ' --- Очищаем целевые ячейки перед заполнением ---
    wsMain.Range(MAIN_CELL_MODEL & ":" & MAIN_CELL_ROW_NUM).ClearContents

    ' --- Проверка входного значения ---
    ' Если номер заказа пустой или не число — выходим без сообщения
    If IsEmpty(orderNumber) Then
        FillHeaderFromOrder = True
        Exit Function
    End If

    If Not IsNumeric(orderNumber) Then
        MsgBox "Номер заказа должен быть числовым значением." & vbCrLf & _
               "Пожалуйста, введите целое положительное число.", _
               vbExclamation, "Некорректный номер заказа"
        FillHeaderFromOrder = True
        Exit Function
    End If

    ' Проверка, что число целое и положительное
    ' Используем CDbl для безопасного преобразования Variant в число,
    ' затем Int для проверки на целость (избегаем банковского округления CLng).
    Dim orderNumDouble As Double
    orderNumDouble = CDbl(orderNumber)

    If orderNumDouble < 1 Or orderNumDouble <> Int(orderNumDouble) Then
        MsgBox "Номер заказа должен быть целым положительным числом." & vbCrLf & _
               "Пожалуйста, введите корректный номер заказа.", _
               vbExclamation, "Некорректный номер заказа"
        FillHeaderFromOrder = True
        Exit Function
    End If

    ' --- Поиск заказа на листе spisok ---
    Set foundRowSpisok = wsSpisok.Columns(SPISOK_COL_ORDER).Find( _
        What:=orderNumber, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If foundRowSpisok Is Nothing Then
        MsgBox "Заказ с номером " & orderNumber & " не найден в справочнике ТС." & vbCrLf & _
               "Пожалуйста, проверьте номер заказа или заполните данные вручную.", _
               vbInformation, "Заказ не найден"
        FillHeaderFromOrder = True
        Exit Function
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

    ' --- B15 = номер строки на spisok минус 1 (т.к. первая строка — заголовок) ---
    wsMain.Range(MAIN_CELL_ROW_NUM).Value = spisokRowIndex - SPISOK_HEADER_ROWS

    ' --- Поиск модели на листе model ---
    Dim modelName As String
    modelName = Trim(CStr(wsMain.Range(MAIN_CELL_MODEL).Value))

    If Len(modelName) > 0 Then
        ' Ищем модель с точным совпадением, без учёта регистра и лишних пробелов.
        ' Проходим по всем строкам столбца 1 листа model, применяем Trim к каждому
        ' значению и сравниваем через StrComp (vbTextCompare — без учёта регистра).
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
                   "Пожалуйста, заполните данные по работам и запчастям вручную.", _
                   vbInformation, "Модель не найдена"
            ' B11:B14 остаются пустыми — очищены в начале
            FillHeaderFromOrder = True
            Exit Function
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
    ' Освобождаем объектные переменные
    If Not foundRowSpisok Is Nothing Then Set foundRowSpisok = Nothing

    ' Сбрасываем обработчик ошибок, чтобы не маскировать ошибки в вызывающем коде
    On Error GoTo 0

    FillHeaderFromOrder = True
    Exit Function

ErrHandler:
    ' Обработка runtime-ошибки
    MsgBox "Произошла непредвиденная ошибка при заполнении шапки заказ-наряда:" & vbCrLf & _
           vbCrLf & _
           "Ошибка: " & Err.Description & vbCrLf & _
           "Номер: " & Err.Number, _
           vbCritical, "Ошибка выполнения"
    FillHeaderFromOrder = False

End Function