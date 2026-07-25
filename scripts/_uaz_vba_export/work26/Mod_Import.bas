Attribute VB_Name = "Mod_Import"
'===========================================================
' МОДУЛЬ: Mod_Import
' НАЗНАЧЕНИЕ: Импорт данных с входящего листа [номер]M на лист main
' ВЕРСИЯ: 2.2 - Поиск "№" теперь в 3-м столбце (можно расширить)
'===========================================================

Option Explicit

' Основная процедура импорта
Public Sub ImportFromIncomingSheet()
    Dim wsMain As Worksheet, wsIncoming As Worksheet
    Dim incomingSheetName As String
    Dim importCountWorks As Long, importCountParts As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Получаем рабочие листы
    Set wsMain = ThisWorkbook.Sheets("main")
    
    ' Проверяем, заполнено ли B2
    If IsEmpty(wsMain.Range("B2").Value) Then
        MsgBox "Введите номер в ячейку B2 для определения входящего листа!", vbExclamation
        Exit Sub
    End If
    
    ' Формируем имя входящего листа
    incomingSheetName = CStr(wsMain.Range("B2").Value) & "M"
    
    ' Проверяем существование входящего листа
    If Not SheetExists(incomingSheetName) Then
        MsgBox "Входящий лист '" & incomingSheetName & "' не найден!" & vbCrLf & _
               "Убедитесь, что лист с таким именем существует.", vbExclamation
        Exit Sub
    End If
    
    Set wsIncoming = ThisWorkbook.Sheets(incomingSheetName)
    
    ' Очищаем старые данные
    ClearImportData wsMain
    
    ' Импорт работ
    importCountWorks = ImportWorksTable(wsMain, wsIncoming)
    
    ' Импорт запчастей
    importCountParts = ImportPartsTable(wsMain, wsIncoming)
    
    ' Форматирование числовых значений
    FormatNumericValues wsMain, importCountWorks, importCountParts
    
    ' Сохраняем настройки
    Call Mod_Settings.SaveMainLayout
    
    ' Показываем результат
    ShowImportResult importCountWorks, importCountParts
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при импорте данных: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' Импорт таблицы работ
Private Function ImportWorksTable(wsMain As Worksheet, wsIncoming As Worksheet) As Long
    Dim foundRow As Long, dataStartRow As Long, targetRow As Long
    Dim importCount As Long
    Dim cellValue As String
    Dim foundHeader As Boolean
    
    foundRow = FindTableStart(wsIncoming, "Выполненные работы")
    
    If foundRow = 0 Then
        ImportWorksTable = 0
        Exit Function
    End If
    
    ' Находим строку с заголовком "№" — теперь ищем в 3-м столбце
    dataStartRow = foundRow + 1
    foundHeader = False
    
    Do While dataStartRow <= wsIncoming.Rows.count
        ' Ищем "№" в 3-м столбце (C)
        If Trim(CStr(wsIncoming.Cells(dataStartRow, 3).Value)) = "№" Then
            foundHeader = True
            Exit Do
        End If
        dataStartRow = dataStartRow + 1
    Loop
    
    If Not foundHeader Then
        ' Если не нашли заголовок, используем запасной вариант смещения
        dataStartRow = foundRow + 3
    Else
        ' Пропускаем строку с заголовком и строку с номерами столбцов
        dataStartRow = dataStartRow + 2
    End If
    
    targetRow = 2
    importCount = 0
    
    ' Импортируем данные до строки с "Итого работ" или пустой строки
    Do While True
        cellValue = Trim(CStr(wsIncoming.Cells(dataStartRow, 1).Value))
        
        ' Проверяем окончание таблицы
        If cellValue = "" Or InStr(1, cellValue, "Итого работ", vbTextCompare) > 0 Then
            Exit Do
        End If
        
        ' Импортируем данные (столбцы C, D, H > L, M, N)
        wsMain.Cells(targetRow, "L").Value = Trim(wsIncoming.Cells(dataStartRow, "C").Value)  ' Наименование
        wsMain.Cells(targetRow, "M").Value = Trim(wsIncoming.Cells(dataStartRow, "D").Value)  ' Кол. оп.
        wsMain.Cells(targetRow, "N").Value = Trim(wsIncoming.Cells(dataStartRow, "H").Value)  ' Всего
        
        targetRow = targetRow + 1
        importCount = importCount + 1
        dataStartRow = dataStartRow + 1
    Loop
    
    ImportWorksTable = importCount
End Function

' Импорт таблицы запчастей
Private Function ImportPartsTable(wsMain As Worksheet, wsIncoming As Worksheet) As Long
    Dim foundRow As Long, dataStartRow As Long, targetRow As Long
    Dim importCount As Long
    Dim cellValue As String
    Dim foundHeader As Boolean
    
    foundRow = FindTableStart(wsIncoming, "Расходная накладная")
    
    If foundRow = 0 Then
        ImportPartsTable = 0
        Exit Function
    End If
    
    ' Находим строку с заголовком "№" — теперь ищем в 3-м столбце
    dataStartRow = foundRow + 1
    foundHeader = False
    
    Do While dataStartRow <= wsIncoming.Rows.count
        ' Ищем "№" в 3-м столбце (C)
        If Trim(CStr(wsIncoming.Cells(dataStartRow, 3).Value)) = "№" Then
            foundHeader = True
            Exit Do
        End If
        dataStartRow = dataStartRow + 1
    Loop
    
    If Not foundHeader Then
        ' Если не нашли заголовок, используем запасной вариант смещения
        dataStartRow = foundRow + 3
    Else
        ' Пропускаем строку с заголовком и строку с номерами столбцов
        dataStartRow = dataStartRow + 2
    End If
    
    targetRow = 2
    importCount = 0
    
    ' Импортируем данные до строки с "Итого материалов" или пустой строки
    Do While True
        cellValue = Trim(CStr(wsIncoming.Cells(dataStartRow, 1).Value))
        
        ' Проверяем окончание таблицы
        If cellValue = "" Or InStr(1, cellValue, "Итого материалов", vbTextCompare) > 0 Then
            Exit Do
        End If
        
        ' Импортируем 4 столбца: B, C, D, G > X, Y, Z, AA
        wsMain.Cells(targetRow, "X").Value = Trim(wsIncoming.Cells(dataStartRow, "B").Value)  ' № кат.
        wsMain.Cells(targetRow, "Y").Value = Trim(wsIncoming.Cells(dataStartRow, "C").Value)  ' Наименование
        wsMain.Cells(targetRow, "Z").Value = Trim(wsIncoming.Cells(dataStartRow, "D").Value)  ' Кол-во
        wsMain.Cells(targetRow, "AA").Value = Trim(wsIncoming.Cells(dataStartRow, "G").Value) ' Всего
        
        targetRow = targetRow + 1
        importCount = importCount + 1
        dataStartRow = dataStartRow + 1
    Loop
    
    ImportPartsTable = importCount
End Function

' Проверка, является ли строка строкой с номерами столбцов
Private Function IsColumnNumbersRow(ws As Worksheet, rowNum As Long) As Boolean
    Dim col As Long
    Dim cellValue As String
    Dim hasNumber As Boolean
    Dim hasNonEmptyCell As Boolean
    
    hasNumber = False
    hasNonEmptyCell = False
    
    ' Проверяем первые 8 столбцов
    For col = 1 To 8
        cellValue = Trim(CStr(ws.Cells(rowNum, col).Value))
        
        ' Если ячейка не пустая
        If cellValue <> "" Then
            hasNonEmptyCell = True
            
            ' Проверяем, является ли значение числом
            If IsNumeric(cellValue) Then
                hasNumber = True
            Else
                ' Если есть не-число, то это не строка с номерами столбцов
                IsColumnNumbersRow = False
                Exit Function
            End If
        End If
    Next col
    
    ' Если были только числа в непустых ячейках — это строка с номерами столбцов
    IsColumnNumbersRow = hasNonEmptyCell And hasNumber
End Function

' Поиск начала таблицы по тексту


