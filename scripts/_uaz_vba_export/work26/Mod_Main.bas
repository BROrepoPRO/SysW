Attribute VB_Name = "Mod_Main"
'===========================================================
' МОДУЛЬ: Mod_Main
' НАЗНАЧЕНИЕ: Заполнение шапки заказ-наряда в новой структуре
' ВЕРСИЯ: Без изменения ширины столбцов
'===========================================================

Option Explicit

' Процедура для кнопки "Заполнить шапку"
Public Sub FillHeaderButton()
    FillHeaderTransposed
End Sub

' Процедура для кнопки "Очистить шапку"
Public Sub ClearHeaderButton()
    ClearHeader
End Sub

' Основная процедура заполнения шапки (транспонированная структура)
Public Sub FillHeaderTransposed()
    Dim wsMain As Worksheet, wsSpisok As Worksheet
    Dim rowNum As Long
    Dim modelName As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Получаем рабочие листы
    Set wsMain = ThisWorkbook.Sheets("main")
    Set wsSpisok = ThisWorkbook.Sheets("spisok")
    
    ' Проверяем, что в B2 есть значение
    If IsEmpty(wsMain.Range("B2").Value) Then
        MsgBox "Введите номер строки из листа 'spisok' в ячейку B2!", vbExclamation
        Exit Sub
    End If
    
    ' Проверяем, что это число
    If Not IsNumeric(wsMain.Range("B2").Value) Then
        MsgBox "В ячейке B2 должен быть номер строки (число)!", vbExclamation
        Exit Sub
    End If
    
    rowNum = CLng(wsMain.Range("B2").Value)
    
    ' Проверяем, что строка существует в spisok
    Dim lastRowSpisok As Long
    lastRowSpisok = wsSpisok.Cells(wsSpisok.Rows.count, "A").End(xlUp).Row
    
    If rowNum < 1 Or rowNum > lastRowSpisok - 1 Then
        MsgBox "Строка с номером " & rowNum & " не найдена в списке автомобилей!", vbExclamation
        Exit Sub
    End If
    
    ' Очищаем старые данные (A3:B15) - только значения, не форматы
    ClearDataOnly wsMain.Range("A3:B15")
    
    ' Заполняем данные в транспонированной структуре
    With wsMain
        ' Устанавливаем заголовки в столбец A (начиная с A3)
        .Range("A3").Value = "название ТС"
        .Range("A4").Value = "ГРЗ"
        .Range("A5").Value = "VIN"
        .Range("A6").Value = "гараж. №"
        .Range("A7").Value = "год вып."
        .Range("A8").Value = "пробег"
        .Range("A9").Value = "дата"
        .Range("A10").Value = "№ ЗН"
        .Range("A11").Value = "цена н/ч"
        .Range("A12").Value = "работы"
        .Range("A13").Value = "з/ч"
        .Range("A14").Value = "исх раб"
        .Range("A15").Value = "№ п/п"
        
        ' Заполняем данные из spisok в столбец B (начиная с B3)
        .Range("B3").Value = wsSpisok.Cells(rowNum + 1, "B").Value  ' Модель
        .Range("B4").Value = wsSpisok.Cells(rowNum + 1, "C").Value  ' ГРЗ
        .Range("B5").Value = wsSpisok.Cells(rowNum + 1, "D").Value  ' VIN
        .Range("B6").Value = wsSpisok.Cells(rowNum + 1, "E").Value  ' Гаражный №
        .Range("B7").Value = wsSpisok.Cells(rowNum + 1, "F").Value  ' Год выпуска
        .Range("B8").Value = wsSpisok.Cells(rowNum + 1, "G").Value  ' Пробег
        
        ' Обработка даты
        Dim dateValue As Variant
        dateValue = wsSpisok.Cells(rowNum + 1, "H").Value
        If IsDate(dateValue) Then
            .Range("B9").Value = dateValue
            .Range("B9").numberFormat = "dd.mm.yy"
        Else
            .Range("B9").Value = dateValue
        End If
        
        ' № ЗН: формат "00" + номер + "-10"
        .Range("B10").Value = "00" & rowNum & "-20"
        
        ' Поиск модели в листе model
        modelName = .Range("B3").Value
        Dim wsModel As Worksheet
        Dim foundRow As Long
        Set wsModel = ThisWorkbook.Sheets("model")
        foundRow = FindModelSimple(wsModel, modelName)
        
        If foundRow > 0 Then
            .Range("B11").Value = wsModel.Cells(foundRow, "B").Value  ' цена н/ч
            .Range("B12").Value = wsModel.Cells(foundRow, "C").Value  ' работы
            .Range("B13").Value = wsModel.Cells(foundRow, "D").Value  ' з/ч
            .Range("B14").Value = wsModel.Cells(foundRow, "E").Value  ' исх раб
        Else
            ' Оставляем пустым, если модель не найдена
            ClearDataOnly .Range("B11:B14")
        End If
        
        ' Сохраняем исходный номер из B2 в B15
        .Range("B15").Value = rowNum
    End With
    
    ' Форматирование (без изменения ширины столбцов)
    FormatTransposedHeaderNoResize wsMain
    
    ' Сохраняем настройки
    Call Mod_Settings.SaveMainLayout
    
    MsgBox "Шапка заказ-наряда успешно заполнена!", vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при заполнении шапки: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' Упрощенный поиск модели
Private Function FindModelSimple(wsModel As Worksheet, modelName As String) As Long
    Dim lastRow As Long, i As Long
    
    If wsModel Is Nothing Then
        FindModelSimple = 0
        Exit Function
    End If
    
    lastRow = wsModel.Cells(wsModel.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsModel.Cells(i, "A").Value) = Trim(modelName) Then
            FindModelSimple = i
            Exit Function
        End If
    Next i
    
    FindModelSimple = 0
End Function

' Очистка только данных (без форматов)
Private Sub ClearDataOnly(rng As Range)
    rng.ClearContents
End Sub

' Очистка шапки (только данные)
Public Sub ClearHeader()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Sheets("main")
    
    Application.EnableEvents = False
    
    ' Очищаем только содержимое, сохраняя форматы
    ClearDataOnly wsMain.Range("B2")
    ClearDataOnly wsMain.Range("A3:B15")
    
    Application.EnableEvents = True
    
    ' Сохраняем настройки
    Call Mod_Settings.SaveMainLayout
    
    MsgBox "Шапка заказ-наряда очищена!", vbInformation
End Sub

' Форматирование без изменения ширины столбцов
Private Sub FormatTransposedHeaderNoResize(ws As Worksheet)
    With ws
        ' Только устанавливаем выравнивание для новых данных
        .Range("A2").horizontalAlignment = xlRight
        .Range("B2").horizontalAlignment = xlLeft
        
        ' Форматирование для новых данных в столбце A
        .Range("A3:A15").horizontalAlignment = xlRight
        
        ' Форматирование для новых данных в столбце B
        .Range("B3:B15").horizontalAlignment = xlLeft
        
        ' Форматирование числа в B9 (дата)
        If IsDate(.Range("B9").Value) Then
            .Range("B9").numberFormat = "dd.mm.yy"
        End If
    End With
End Sub

' Инициализация листа main (упрощенная)
Public Sub InitializeMainSheet()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Sheets("main")
    
    ' Восстанавливаем сохраненные настройки
    Call Mod_Settings.RestoreMainLayout
    
    ' Создаем заголовки таблиц (только если их нет)
    Call Mod_Settings.CreateTableHeaders
    
    ' Устанавливаем базовую структуру
    wsMain.Range("A1").Value = "Поле"
    wsMain.Range("B1").Value = "Значение"
    wsMain.Range("A2").Value = "Введите номер:"
    
    ' Форматируем базовые заголовки (без изменения ширины)
    With wsMain.Range("A1:B1")
        .Font.Bold = True
        .horizontalAlignment = xlCenter
    End With
    
    With wsMain.Range("A2")
        .Font.Bold = True
        .horizontalAlignment = xlRight
    End With
End Sub
