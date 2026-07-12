Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================
' Модуль: Mod_Import
' Назначение: импорт данных из Excel в SQLite
' ============================================================

' Константы
Private Const START_ROW As Long = 2
Private Const COL_ORDER As Long = 2  ' B
Private Const COL_CLIENT As Long = 3 ' C
Private Const COL_CAR As Long = 4    ' D
Private Const COL_PLATE As Long = 5  ' E
Private Const COL_MILEAGE As Long = 6 ' F
Private Const COL_DATE_IN As Long = 7 ' G
Private Const COL_DATE_OUT As Long = 8 ' H
Private Const COL_STATUS As Long = 9  ' I

' Главная процедура импорта
Public Sub RunImport()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim ImportCount As Long
    
    Set ws = ThisWorkbook.Sheets("Лист1")
    LastRow = ws.Cells(ws.Rows.Count, COL_ORDER).End(xlUp).Row
    
    If LastRow < START_ROW Then
        MsgBox "Нет данных для импорта.", vbInformation, "Импорт"
        Exit Sub
    End If
    
    ImportCount = 0
    For i = START_ROW To LastRow
        If Trim(ws.Cells(i, COL_ORDER).Value) <> "" Then
            ' Здесь будет вызов SQLite-вставки
            ImportCount = ImportCount + 1
        End If
    Next i
    
    MsgBox "Импорт завершён. Обработано строк: " & ImportCount, vbInformation, "Импорт"
End Sub

' Процедура: проверка данных перед импортом
Public Function ValidateData(ByVal ws As Worksheet) As String
    Dim LastRow As Long
    Dim i As Long
    Dim Errors As String
    
    LastRow = ws.Cells(ws.Rows.Count, COL_ORDER).End(xlUp).Row
    
    For i = START_ROW To LastRow
        If Trim(ws.Cells(i, COL_ORDER).Value) <> "" Then
            If Trim(ws.Cells(i, COL_CLIENT).Value) = "" Then
                Errors = Errors & "Строка " & i & ": не указан клиент" & vbCrLf
            End If
            If Trim(ws.Cells(i, COL_CAR).Value) = "" Then
                Errors = Errors & "Строка " & i & ": не указан автомобиль" & vbCrLf
            End If
        End If
    Next i
    
    ValidateData = Errors
End Function