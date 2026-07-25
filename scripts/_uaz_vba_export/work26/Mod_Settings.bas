Attribute VB_Name = "Mod_Settings"
'===========================================================
' МОДУЛЬ: Mod_Settings
' НАЗНАЧЕНИЕ: Управление настройками листа main, сохранение макета
' ВЕРСИЯ: 1.1 - Исправленные заголовки
'===========================================================

Option Explicit

' Константы для хранения настроек
Private Const SETTINGS_SHEET As String = "_SETTINGS"
Private Const MAIN_SHEET As String = "main"

' Сохранение текущих настроек листа main
Public Sub SaveMainLayout()
    Dim wsMain As Worksheet, wsSettings As Worksheet
    Dim i As Long, col As Long
    Dim settingRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsMain = ThisWorkbook.Sheets(MAIN_SHEET)
    
    ' Создаем или получаем лист настроек
    If Not SheetExists(SETTINGS_SHEET) Then
        Set wsSettings = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsSettings.Name = SETTINGS_SHEET
        wsSettings.Visible = xlSheetVeryHidden
        InitializeSettingsSheet wsSettings
    Else
        Set wsSettings = ThisWorkbook.Sheets(SETTINGS_SHEET)
    End If
    
    ' Сохраняем ширину всех столбцов (A:AB)
    settingRow = 1
    For col = 1 To 28 ' A=1, AB=28
        wsSettings.Cells(settingRow, col).Value = wsMain.Columns(col).ColumnWidth
    Next col
    
    ' Сохраняем высоту строк 1-15
    settingRow = 2
    For i = 1 To 15
        wsSettings.Cells(settingRow, i).Value = wsMain.Rows(i).RowHeight
    Next i
    
    ' Сохраняем заголовки таблиц
    SaveTableHeaders wsMain, wsSettings
    
    ' Сохраняем форматы заголовков
    SaveHeaderFormats wsMain, wsSettings
    
    ' Сохраняем дату последнего изменения
    wsSettings.Cells(10, 1).Value = "Дата сохранения:"
    wsSettings.Cells(10, 2).Value = Now
    
ErrorHandler:
    If Err.number <> 0 Then
        Debug.Print "Ошибка при сохранении настроек: " & Err.Description
    End If
End Sub

' Восстановление настроек листа main
Public Sub RestoreMainLayout()
    Dim wsMain As Worksheet, wsSettings As Worksheet
    Dim i As Long, col As Long
    Dim settingRow As Long
    
    On Error Resume Next
    
    Set wsMain = ThisWorkbook.Sheets(MAIN_SHEET)
    
    If Not SheetExists(SETTINGS_SHEET) Then
        Exit Sub
    End If
    
    Set wsSettings = ThisWorkbook.Sheets(SETTINGS_SHEET)
    
    ' Восстанавливаем ширину всех столбцов (A:AB)
    settingRow = 1
    For col = 1 To 28
        If IsNumeric(wsSettings.Cells(settingRow, col).Value) Then
            wsMain.Columns(col).ColumnWidth = wsSettings.Cells(settingRow, col).Value
        End If
    Next col
    
    ' Восстанавливаем высоту строк 1-15
    settingRow = 2
    For i = 1 To 15
        If IsNumeric(wsSettings.Cells(settingRow, i).Value) Then
            wsMain.Rows(i).RowHeight = wsSettings.Cells(settingRow, i).Value
        End If
    Next i
    
    ' Восстанавливаем форматы заголовков
    RestoreHeaderFormats wsMain, wsSettings
    
    On Error GoTo 0
End Sub

' Инициализация листа настроек
Private Sub InitializeSettingsSheet(ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "Ширина столбцов A:AB"
        .Cells(2, 1).Value = "Высота строк 1-15"
        .Cells(3, 1).Value = "Заголовки таблиц"
        .Cells(4, 1).Value = "Форматы заголовков"
    End With
End Sub

' Сохранение заголовков таблиц
Private Sub SaveTableHeaders(wsMain As Worksheet, wsSettings As Worksheet)
    Dim headerRow As Long
    Dim col As Long
    
    headerRow = 1
    
    ' Сохраняем заголовки из строки 1
    wsSettings.Cells(3, 1).Value = "Заголовки строки 1:"
    For col = 1 To 28
        wsSettings.Cells(3, col + 1).Value = wsMain.Cells(headerRow, col).Value
    Next col
End Sub

' Сохранение форматов заголовков
Private Sub SaveHeaderFormats(wsMain As Worksheet, wsSettings As Worksheet)
    Dim col As Long
    
    wsSettings.Cells(4, 1).Value = "Форматы заголовков строки 1:"
    
    For col = 1 To 28
        With wsMain.Cells(1, col)
            ' Сохраняем свойства форматирования
            wsSettings.Cells(4, col + 1).Value = .Font.Name
            wsSettings.Cells(5, col + 1).Value = .Font.Size
            wsSettings.Cells(6, col + 1).Value = .Font.Bold
            wsSettings.Cells(7, col + 1).Value = .Font.Color
            wsSettings.Cells(8, col + 1).Value = .horizontalAlignment
            wsSettings.Cells(9, col + 1).Value = .Interior.Color
        End With
    Next col
End Sub

' Восстановление форматов заголовков
Private Sub RestoreHeaderFormats(wsMain As Worksheet, wsSettings As Worksheet)
    Dim col As Long
    
    For col = 1 To 28
        With wsMain.Cells(1, col)
            ' Восстанавливаем свойства форматирования
            If wsSettings.Cells(4, col + 1).Value <> "" Then
                .Font.Name = wsSettings.Cells(4, col + 1).Value
            End If
            
            If IsNumeric(wsSettings.Cells(5, col + 1).Value) Then
                .Font.Size = wsSettings.Cells(5, col + 1).Value
            End If
            
            If wsSettings.Cells(6, col + 1).Value <> "" Then
                .Font.Bold = wsSettings.Cells(6, col + 1).Value
            End If
            
            If IsNumeric(wsSettings.Cells(7, col + 1).Value) Then
                .Font.Color = wsSettings.Cells(7, col + 1).Value
            End If
            
            If IsNumeric(wsSettings.Cells(8, col + 1).Value) Then
                .horizontalAlignment = wsSettings.Cells(8, col + 1).Value
            End If
            
            If IsNumeric(wsSettings.Cells(9, col + 1).Value) Then
                .Interior.Color = wsSettings.Cells(9, col + 1).Value
            End If
        End With
    Next col
End Sub

' Проверка существования листа
Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' Создание заголовков таблиц (только если их нет)
Public Sub CreateTableHeaders()
    Dim wsMain As Worksheet
    Dim headersCreated As Boolean
    
    Set wsMain = ThisWorkbook.Sheets(MAIN_SHEET)
    
    ' Проверяем, есть ли уже заголовки
    If wsMain.Range("D1").Value = "" Then
        ' Таблица 1: Работы (для подбора) - D1:J1
        wsMain.Range("D1").Value = "№ п/п"
        wsMain.Range("E1").Value = "Артикул"
        wsMain.Range("F1").Value = "Наименование"
        wsMain.Range("G1").Value = "Кол-во н/ч"
        wsMain.Range("H1").Value = "Кол-во оп"
        wsMain.Range("I1").Value = "Цена н/ч"
        wsMain.Range("J1").Value = "Сумма"
        
        ' Форматирование заголовков таблицы 1
        FormatTableHeaders wsMain.Range("D1:J1")
    End If
    
    If wsMain.Range("L1").Value = "" Then
        ' Таблица 2: Работы (импорт) - L1:N1
        wsMain.Range("L1").Value = "Наименование"
        wsMain.Range("M1").Value = "Кол. оп."
        wsMain.Range("N1").Value = "Всего"
        
        ' Форматирование заголовков таблицы 2
        FormatTableHeaders wsMain.Range("L1:N1")
    End If
    
    If wsMain.Range("P1").Value = "" Then
        ' Таблица 3: Запчасти (для подбора) - P1:V1
        wsMain.Range("P1").Value = "№ п/п"
        wsMain.Range("Q1").Value = "Артикул"
        wsMain.Range("R1").Value = "Наименование"
        wsMain.Range("S1").Value = "Цена"
        wsMain.Range("T1").Value = "Кол-во"
        wsMain.Range("U1").Value = "Ед. изм."
        wsMain.Range("V1").Value = "Сумма"
        
        ' Форматирование заголовков таблицы 3
        FormatTableHeaders wsMain.Range("P1:V1")
    End If
    
    If wsMain.Range("X1").Value = "" Then
        ' Таблица 4: Запчасти (импорт) - X1:AA1 (4 столбца)
        wsMain.Range("X1").Value = "№ кат."
        wsMain.Range("Y1").Value = "Наименование"
        wsMain.Range("Z1").Value = "Кол-во"
        wsMain.Range("AA1").Value = "Всего"
        
        ' Форматирование заголовков таблицы 4
        FormatTableHeaders wsMain.Range("X1:AA1")
    End If
End Sub

' Форматирование заголовков таблиц (только при создании)
Private Sub FormatTableHeaders(rng As Range)
    With rng
        .Font.Bold = True
        .horizontalAlignment = xlCenter
        .verticalAlignment = xlCenter
        .WrapText = True
        
        ' Стандартный цвет фона
        .Interior.Color = RGB(240, 240, 240)
        
        ' Стандартные границы
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
End Sub
