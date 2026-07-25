Attribute VB_Name = "Mod_ZN"
'===========================================================
' МОДУЛЬ: Mod_ZN
' НАЗНАЧЕНИЕ: Обработка листа "ЗН" для формирования текстового представления
' ВЕРСИЯ: 3.0 - Исправление проблемы с форматированием и границами
'===========================================================

Option Explicit

' Основная процедура обработки листа ЗН
Public Sub toTEXT()
    Dim wsNZ As Worksheet
    
    On Error Resume Next
    Set wsNZ = ThisWorkbook.Sheets("ЗН")
    On Error GoTo 0
    
    If wsNZ Is Nothing Then
        MsgBox "Лист 'ЗН' не найден в книге.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Шаг 1: Выполняем существующий алгоритм
    HideRowsBasedOnValues wsNZ
    ProcessAmountToText wsNZ
    
    ' Шаг 2: Преобразуем все формулы в значения БЕЗ изменения форматирования
    ConvertFormulasToValuesPreservingFormat wsNZ
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Обработка листа 'ЗН' завершена:" & vbCrLf & _
           "? Строки отфильтрованы" & vbCrLf & _
           "? Сумма преобразована в текст" & vbCrLf & _
           "? Все формулы заменены на значения" & vbCrLf & _
           "? Форматы полностью сохранены", vbInformation, "Готово"
End Sub

' Процедура скрытия строк на основе значений в столбце B
Private Sub HideRowsBasedOnValues(ws As Worksheet)
    Dim i As Long
    
    With ws
        ' Первый диапазон: строки 17-203
        For i = 17 To 203
            If .Cells(i, "B").Value = "" Or .Cells(i, "B").Value = 0 Then
                .Rows(i).Hidden = True
            Else
                .Rows(i).Hidden = False
            End If
        Next i

        ' Второй диапазон: строки 209-308
        For i = 209 To 308
            If .Cells(i, "B").Value = "" Or .Cells(i, "B").Value = 0 Then
                .Rows(i).Hidden = True
            Else
                .Rows(i).Hidden = False
            End If
        Next i
    End With
End Sub

' Процедура обработки суммы и преобразования в текст
Private Sub ProcessAmountToText(ws As Worksheet)
    Dim amountK311 As Double
    
    ' Проверяем ячейку K311
    If Not IsEmpty(ws.Range("K311").Value) And IsNumeric(ws.Range("K311").Value) Then
        amountK311 = ws.Range("K311").Value
        
        ' Получаем текстовое представление суммы
        Dim rublesText As String
        rublesText = ConvertNumberToWords(Int(amountK311), False)
        
        ' Получаем копейки
        Dim kopiykas As Long
        kopiykas = Round((amountK311 - Int(amountK311)) * 100)
        Dim kopiykasText As String
        kopiykasText = Format(kopiykas, "00")
        
        ' Форма слова "рубль"
        Dim rublesForm As String
        rublesForm = ProperRublesForm(Int(amountK311))
        
        ' Форма слова "копейка"
        Dim kopiykasForm As String
        kopiykasForm = KopiykasToText(kopiykas)
        
        ' Формируем итоговую строку
        ws.Range("G313").Value = CapitalizeFirstLetter(rublesText) & " " & rublesForm & _
                                 " " & kopiykasText & " " & kopiykasForm & ". в т.ч. НДС 5%."
    Else
        ' Если K311 пустая или не число - очищаем G313
        ws.Range("G313").ClearContents
    End If
End Sub

' НОВАЯ УЛУЧШЕННАЯ ПРОЦЕДУРА: Преобразование формул в значения с полным сохранением форматирования
Private Sub ConvertFormulasToValuesPreservingFormat(ws As Worksheet)
    Dim formulaCells As Range
    Dim cell As Range
    Dim processedMergeAreas As Collection
    Dim i As Long
    
    ' Создаем коллекцию для отслеживания обработанных объединенных областей
    Set processedMergeAreas = New Collection
    
    ' Находим все ячейки с формулами
    On Error Resume Next
    Set formulaCells = ws.usedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then
        Exit Sub ' Нет формул на листе
    End If
    
    ' ШАГ 1: Обрабатываем ВСЕ объединенные ячейки с формулами (каждую область один раз)
    For Each cell In formulaCells
        If cell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = cell.mergeArea
            Dim topLeftCell As Range
            Set topLeftCell = mergeArea.Cells(1, 1)
            
            ' Проверяем, не обрабатывали ли мы уже эту объединенную область
            If Not IsRangeInCollection(processedMergeAreas, mergeArea) Then
                ' Сохраняем и восстанавливаем значение для объединенной ячейки
                PreserveMergedCellValue topLeftCell
                ' Добавляем область в коллекцию обработанных
                processedMergeAreas.Add mergeArea, mergeArea.Address
            End If
        End If
    Next cell
    
    ' ШАГ 2: Обрабатываем все НЕобъединенные ячейки с формулами
    For Each cell In formulaCells
        If Not cell.MergeCells Then
            ' Простое преобразование формулы в значение для обычной ячейки
            Dim cellValue As Variant
            Dim cellFormat As String
            
            cellValue = cell.Value
            cellFormat = cell.numberFormat
            
            ' Заменяем формулу на значение
            cell.Value = cellValue
            
            ' Восстанавливаем числовой формат
            cell.numberFormat = cellFormat
        End If
    Next cell
    
    ' Очищаем память
    Set processedMergeAreas = Nothing
    Set formulaCells = Nothing
End Sub

' Процедура сохранения и восстановления значения в объединенной ячейке
Private Sub PreserveMergedCellValue(cell As Range)
    ' Этот метод работает БЕЗ разъединения ячеек и БЕЗ изменения границ
    
    ' 1. Сохраняем все важные свойства ячейки
    Dim originalValue As Variant
    Dim numberFormat As String
    Dim horizontalAlignment As XlHAlign
    Dim verticalAlignment As XlVAlign
    Dim interiorColor As Long
    Dim interiorPattern As XlPattern
    Dim fontName As String
    Dim fontSize As Double
    Dim fontBold As Boolean
    Dim fontColor As Long
    
    On Error Resume Next
    
    ' Сохраняем основные свойства
    originalValue = cell.Value
    numberFormat = cell.numberFormat
    horizontalAlignment = cell.horizontalAlignment
    verticalAlignment = cell.verticalAlignment
    interiorColor = cell.Interior.Color
    interiorPattern = cell.Interior.Pattern
    fontName = cell.Font.Name
    fontSize = cell.Font.Size
    fontBold = cell.Font.Bold
    fontColor = cell.Font.Color
    
    ' 2. Сохраняем информацию о границах (только внешние границы)
    Dim borderInfo(1 To 4, 1 To 3) As Variant ' 4 стороны: левая, верхняя, правая, нижняя
    
    With cell
        ' Левая граница
        borderInfo(1, 1) = .Borders(xlEdgeLeft).LineStyle
        borderInfo(1, 2) = .Borders(xlEdgeLeft).Weight
        borderInfo(1, 3) = .Borders(xlEdgeLeft).Color
        
        ' Верхняя граница
        borderInfo(2, 1) = .Borders(xlEdgeTop).LineStyle
        borderInfo(2, 2) = .Borders(xlEdgeTop).Weight
        borderInfo(2, 3) = .Borders(xlEdgeTop).Color
        
        ' Правая граница
        borderInfo(3, 1) = .Borders(xlEdgeRight).LineStyle
        borderInfo(3, 2) = .Borders(xlEdgeRight).Weight
        borderInfo(3, 3) = .Borders(xlEdgeRight).Color
        
        ' Нижняя граница
        borderInfo(4, 1) = .Borders(xlEdgeBottom).LineStyle
        borderInfo(4, 2) = .Borders(xlEdgeBottom).Weight
        borderInfo(4, 3) = .Borders(xlEdgeBottom).Color
    End With
    
    ' 3. Специальный трюк для объединенных ячеек: используем Copy/Paste Special
    ' Но делаем это так, чтобы не нарушить форматирование
    
    ' Метод 1: Простое присваивание значения (самый безопасный способ)
    cell.Value = originalValue
    
    ' 4. Восстанавливаем все свойства форматирования
    cell.numberFormat = numberFormat
    cell.horizontalAlignment = horizontalAlignment
    cell.verticalAlignment = verticalAlignment
    cell.Interior.Color = interiorColor
    cell.Interior.Pattern = interiorPattern
    cell.Font.Name = fontName
    cell.Font.Size = fontSize
    cell.Font.Bold = fontBold
    cell.Font.Color = fontColor
    
    ' 5. Восстанавливаем границы (только если они были явно заданы)
    With cell
        If borderInfo(1, 1) <> xlLineStyleNone Then
            .Borders(xlEdgeLeft).LineStyle = borderInfo(1, 1)
            .Borders(xlEdgeLeft).Weight = borderInfo(1, 2)
            .Borders(xlEdgeLeft).Color = borderInfo(1, 3)
        End If
        
        If borderInfo(2, 1) <> xlLineStyleNone Then
            .Borders(xlEdgeTop).LineStyle = borderInfo(2, 1)
            .Borders(xlEdgeTop).Weight = borderInfo(2, 2)
            .Borders(xlEdgeTop).Color = borderInfo(2, 3)
        End If
        
        If borderInfo(3, 1) <> xlLineStyleNone Then
            .Borders(xlEdgeRight).LineStyle = borderInfo(3, 1)
            .Borders(xlEdgeRight).Weight = borderInfo(3, 2)
            .Borders(xlEdgeRight).Color = borderInfo(3, 3)
        End If
        
        If borderInfo(4, 1) <> xlLineStyleNone Then
            .Borders(xlEdgeBottom).LineStyle = borderInfo(4, 1)
            .Borders(xlEdgeBottom).Weight = borderInfo(4, 2)
            .Borders(xlEdgeBottom).Color = borderInfo(4, 3)
        End If
    End With
    
    On Error GoTo 0
End Sub

' Функция проверки наличия диапазона в коллекции
Private Function IsRangeInCollection(col As Collection, rng As Range) As Boolean
    Dim item As Range
    
    On Error GoTo NotInCollection
    
    For Each item In col
        If item.Address = rng.Address Then
            IsRangeInCollection = True
            Exit Function
        End If
    Next item
    
NotInCollection:
    IsRangeInCollection = False
End Function

' Альтернативный метод: Простая замена формул на значения БЕЗ сохранения границ (для тестирования)
Private Sub ConvertFormulasToValuesSimple(ws As Worksheet)
    Dim formulaCells As Range
    Dim cell As Range
    
    ' Находим все ячейки с формулами
    On Error Resume Next
    Set formulaCells = ws.usedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then
        Exit Sub
    End If
    
    ' Просто заменяем формулы на значения
    formulaCells.Value = formulaCells.Value
End Sub

' Функция подсчета видимых строк
Public Function CountVisibleRows(ws As Worksheet, startRow As Long, endRow As Long) As Long
    Dim count As Long
    Dim i As Long
    
    count = 0
    
    For i = startRow To endRow
        If Not ws.Rows(i).Hidden Then
            count = count + 1
        End If
    Next i
    
    CountVisibleRows = count
End Function

' Функция преобразования числа в текст с рублями и копейками
Public Function NumberToText(amount As Double) As String
    Dim rubles As Long
    Dim kopiykas As Long
    Dim rublesText As String
    Dim kopiykasText As String

    rubles = Int(amount)
    kopiykas = Round((amount - rubles) * 100)
    
    rublesText = ConvertNumberToWords(rubles, False) & " " & ProperRublesForm(rubles)
    kopiykasText = Format(kopiykas, "00") & " " & KopiykasToText(kopiykas)

    NumberToText = rublesText & " " & kopiykasText
End Function

' Основная функция преобразования числа в слова
Private Function ConvertNumberToWords(number As Long, Optional isFemale As Boolean = False) As String
    Dim ones() As String, tens() As String, hundreds() As String
    Dim feminineOnes() As String
    Dim teens() As String
    Dim words As String
    Dim thousands As Long
    Dim hundredIndex As Long, teenIndex As Long, tensIndex As Long
    
    ' Инициализация массивов
    ones = Split("ноль один два три четыре пять шесть семь восемь девять", " ")
    feminineOnes = Split("ноль одна две три четыре пять шесть семь восемь девять", " ")
    tens = Split("десять двадцать тридцать сорок пятьдесят шестьдесят семьдесят восемьдесят девяносто", " ")
    hundreds = Split("сто двести триста четыреста пятьсот шестьсот семьсот восемьсот девятьсот", " ")
    teens = Split("одиннадцать двенадцать тринадцать четырнадцать пятнадцать шестнадцать семнадцать восемнадцать девятнадцать", " ")

    If number = 0 Then
        ConvertNumberToWords = ones(0)
        Exit Function
    End If

    words = ""

    ' Обработка тысяч
    If number >= 1000 Then
        thousands = number \ 1000
        words = words & ConvertNumberToWords(thousands, True) & " " & ThousandsForm(thousands) & " "
        number = number Mod 1000
    End If

    ' Обработка сотен
    If number >= 100 Then
        hundredIndex = (number \ 100) - 1
        If hundredIndex >= 0 And hundredIndex <= UBound(hundreds) Then
            words = words & hundreds(hundredIndex) & " "
        End If
        number = number Mod 100
    End If

    ' Обработка чисел от 11 до 19
    If number >= 11 And number <= 19 Then
        teenIndex = number - 11
        If teenIndex >= 0 And teenIndex <= UBound(teens) Then
            words = words & teens(teenIndex) & " "
        End If
    Else
        ' Обработка десятков
        If number >= 20 Or number = 10 Then
            tensIndex = (number \ 10) - 1
            If tensIndex >= 0 And tensIndex <= UBound(tens) Then
                words = words & tens(tensIndex) & " "
            End If
            number = number Mod 10
        End If
        
        ' Обработка единиц
        If number > 0 Then
            If number >= 0 And number <= UBound(ones) Then
                If isFemale Then
                    words = words & feminineOnes(number) & " "
                Else
                    words = words & ones(number) & " "
                End If
            End If
        End If
    End If

    ConvertNumberToWords = Trim(words)
End Function

' Функция склонения слова "тысяча"
Private Function ThousandsForm(thousands As Long) As String
    Dim mod10 As Long, mod100 As Long
    mod10 = thousands Mod 10
    mod100 = thousands Mod 100

    If mod10 = 1 And mod100 <> 11 Then
        ThousandsForm = "тысяча"
    ElseIf (mod10 >= 2 And mod10 <= 4) And Not (mod100 >= 12 And mod100 <= 14) Then
        ThousandsForm = "тысячи"
    Else
        ThousandsForm = "тысяч"
    End If
End Function

' Функция склонения слова "рубль"
Private Function ProperRublesForm(rubles As Long) As String
    Dim mod10 As Long, mod100 As Long
    mod10 = rubles Mod 10
    mod100 = rubles Mod 100
    
    ' Особые случаи для чисел, оканчивающихся на 11-14
    If mod100 >= 11 And mod100 <= 14 Then
        ProperRublesForm = "рублей"
    Else
        Select Case mod10
            Case 1
                ProperRublesForm = "рубль"
            Case 2, 3, 4
                ProperRublesForm = "рубля"
            Case Else
                ProperRublesForm = "рублей"
        End Select
    End If
End Function

' Функция склонения слова "копейка"
Private Function KopiykasToText(kopiykas As Long) As String
    Dim mod10 As Long, mod100 As Long
    mod10 = kopiykas Mod 10
    mod100 = kopiykas Mod 100

    If mod10 = 1 And mod100 <> 11 Then
        KopiykasToText = "копейка"
    ElseIf (mod10 >= 2 And mod10 <= 4) And Not (mod100 >= 12 And mod100 <= 14) Then
        KopiykasToText = "копейки"
    Else
        KopiykasToText = "копеек"
    End If
End Function

' Функция капитализации первой буквы
Private Function CapitalizeFirstLetter(text As String) As String
    If Len(text) > 0 Then
        CapitalizeFirstLetter = UCase(Left(text, 1)) & Mid(text, 2)
    Else
        CapitalizeFirstLetter = text
    End If
End Function

' Процедура для кнопки "Преобразовать в текст" (с выбором метода)
Public Sub ConvertToTextButton()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Выберите метод преобразования:" & vbCrLf & vbCrLf & _
                     "Да - Сохранить все форматирование (рекомендуется)" & vbCrLf & _
                     "Нет - Простая замена формул на значения" & vbCrLf & _
                     "Отмена - Отменить операцию", _
                     vbYesNoCancel + vbQuestion, "Выбор метода")
    
    If response = vbYes Then
        ' Используем метод с сохранением форматирования
        Call toTEXT
    ElseIf response = vbNo Then
        ' Используем простой метод
        ConvertFormulasSimple
    End If
    ' При vbCancel ничего не делаем
End Sub

' Простая замена формул на значения (без сохранения форматирования)
Public Sub ConvertFormulasSimple()
    Dim wsNZ As Worksheet
    
    On Error Resume Next
    Set wsNZ = ThisWorkbook.Sheets("ЗН")
    On Error GoTo 0
    
    If wsNZ Is Nothing Then
        MsgBox "Лист 'ЗН' не найден в книге.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Выполняем существующий алгоритм
    HideRowsBasedOnValues wsNZ
    ProcessAmountToText wsNZ
    
    ' Простая замена формул на значения
    ConvertFormulasToValuesSimple wsNZ
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Простая замена формул завершена." & vbCrLf & _
           "Формулы заменены на значения.", vbInformation, "Готово"
End Sub

' Функция: Проверка наличия формул на листе
Public Function HasFormulas(ws As Worksheet) As Boolean
    Dim formulaCells As Range
    
    On Error Resume Next
    Set formulaCells = ws.usedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    HasFormulas = Not formulaCells Is Nothing
End Function

' Процедура: Информация о листе
Public Sub ShowSheetInfo()
    Dim ws As Worksheet
    Dim formulaCount As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim mergeCount As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ЗН")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Лист 'ЗН' не найден", vbExclamation
        Exit Sub
    End If
    
    ' Подсчет формул
    formulaCount = 0
    On Error Resume Next
    Dim formulaRange As Range
    Set formulaRange = ws.usedRange.SpecialCells(xlCellTypeFormulas)
    If Not formulaRange Is Nothing Then
        formulaCount = formulaRange.CountLarge
    End If
    On Error GoTo 0
    
    ' Подсчет объединенных областей
    mergeCount = CountMergeAreas(ws)
    
    ' Поиск последних заполненных строки и столбца
    On Error Resume Next
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    
    MsgBox "Информация о листе 'ЗН':" & vbCrLf & _
           "------------------------" & vbCrLf & _
           "Формул на листе: " & formulaCount & vbCrLf & _
           "Объединенных областей: " & mergeCount & vbCrLf & _
           "Последняя строка: " & lastRow & vbCrLf & _
           "Последний столбец: " & Split(ws.Cells(1, lastCol).Address, "$")(1) & vbCrLf & _
           "Используемый диапазон: A1:" & Split(ws.Cells(lastRow, lastCol).Address, "$")(1) & lastRow, _
           vbInformation, "Информация о листе"
End Sub

' Функция подсчета объединенных областей
Private Function CountMergeAreas(ws As Worksheet) As Long
    Dim mergeCount As Long
    Dim cell As Range
    Dim processedAreas As Collection
    
    Set processedAreas = New Collection
    mergeCount = 0
    
    For Each cell In ws.usedRange
        If cell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = cell.mergeArea
            
            ' Проверяем, не считали ли мы уже эту область
            On Error Resume Next
            processedAreas.Add mergeArea, mergeArea.Address
            If Err.number = 0 Then
                ' Если добавление успешно (не было дубликата)
                mergeCount = mergeCount + 1
            End If
            On Error GoTo 0
        End If
    Next cell
    
    CountMergeAreas = mergeCount
End Function

