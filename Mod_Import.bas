Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================
' Модуль: Mod_Import
' Назначение: Импорт данных из Excel в SQLite и между листами
' ============================================================

' Извлекает цифры из госномера
Public Function ExtractNumberFromGRZ(ByVal GRZ As String) As String
    Dim i As Long
    Dim Result As String
    Result = ""
    For i = 1 To Len(GRZ)
        If Mid(GRZ, i, 1) Like "[0-9]" Then
            Result = Result & Mid(GRZ, i, 1)
        End If
    Next i
    ExtractNumberFromGRZ = Result
End Function

' Переименовывает листы по госномеру из ячейки B2
Public Sub RenameSheetsByGRZ()
    Dim ws As Worksheet
    Dim GRZ As String
    Dim SheetNum As String
    
    For Each ws In ThisWorkbook.Sheets
        GRZ = Trim(ws.Range("B2").Value)
        If GRZ <> "" And ws.Name <> "main" And ws.Name <> "spisok" And ws.Name <> "model" Then
            SheetNum = ExtractNumberFromGRZ(GRZ)
            If SheetNum <> "" Then
                ws.Name = "GRZ_" & SheetNum
            End If
        End If
    Next ws
End Sub

' Ищет лист по госномеру в имени
Public Function SearchSheetByGRZ(ByVal GRZ As String) As Worksheet
    Dim ws As Worksheet
    Dim GRZNum As String
    
    GRZNum = ExtractNumberFromGRZ(GRZ)
    If GRZNum = "" Then
        Set SearchSheetByGRZ = Nothing
        Exit Function
    End If
    
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, GRZNum, vbTextCompare) > 0 Then
            Set SearchSheetByGRZ = ws
            Exit Function
        End If
    Next ws
    
    Set SearchSheetByGRZ = Nothing
End Function

' Импорт данных с листа по госномеру
Public Sub ImportSheet(ByVal GRZ As String)
    Dim ws As Worksheet
    Set ws = SearchSheetByGRZ(GRZ)
    
    If ws Is Nothing Then
        MsgBox "Лист с госномером " & GRZ & " не найден.", vbExclamation, "Импорт"
        Exit Sub
    End If
    
    ImportDataToMain ws
End Sub

' Импорт входящего номера документа
Public Sub ImportIncomingDocNumber(ByVal GRZ As String)
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim wsMain As Worksheet
    
    Set ws = SearchSheetByGRZ(GRZ)
    If ws Is Nothing Then Exit Sub
    
    Set wsMain = ThisWorkbook.Sheets("main")
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To LastRow
        If Trim(ws.Cells(i, "A").Value) <> "" Then
            wsMain.Cells(wsMain.Rows.Count, "W").End(xlUp).Offset(1, 0).Value = ws.Cells(i, "A").Value
        End If
    Next i
End Sub

' Импорт данных в лист main с маппингом столбцов
Public Sub ImportDataToMain(ByVal wsSource As Worksheet)
    Dim wsMain As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim DestRow As Long
    
    Set wsMain = ThisWorkbook.Sheets("main")
    LastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    If LastRow < 2 Then Exit Sub
    
    For i = 2 To LastRow
        If Trim(wsSource.Cells(i, "B").Value) <> "" Then
            DestRow = wsMain.Cells(wsMain.Rows.Count, "L").End(xlUp).Row + 1
            If DestRow < 2 Then DestRow = 2
            
            ' Маппинг столбцов:
            ' C(source) -> L(main)
            ' D(source) -> M(main)
            ' H(source) -> N(main)
            ' B(source) -> X(main)
            ' C(source) -> Y(main)
            ' D(source) -> Z(main)
            ' G(source) -> AA(main)
            wsMain.Cells(DestRow, "L").Value = wsSource.Cells(i, "C").Value
            wsMain.Cells(DestRow, "M").Value = wsSource.Cells(i, "D").Value
            wsMain.Cells(DestRow, "N").Value = wsSource.Cells(i, "H").Value
            wsMain.Cells(DestRow, "X").Value = wsSource.Cells(i, "B").Value
            wsMain.Cells(DestRow, "Y").Value = wsSource.Cells(i, "C").Value
            wsMain.Cells(DestRow, "Z").Value = wsSource.Cells(i, "D").Value
            wsMain.Cells(DestRow, "AA").Value = wsSource.Cells(i, "G").Value
        End If
    Next i
    
    MsgBox "Импорт завершен. Обработано строк: " & (DestRow - 1), vbInformation, "Импорт"
End Sub

' Импорт из отчёта
Public Sub ImportFromReport()
    Dim wsReport As Worksheet
    Dim wsMain As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim DestRow As Long
    
    Set wsReport = GetSheetByName(ThisWorkbook, "report")
    If wsReport Is Nothing Then
        MsgBox "Лист 'report' не найден.", vbExclamation, "Импорт"
        Exit Sub
    End If
    
    Set wsMain = ThisWorkbook.Sheets("main")
    LastRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To LastRow
        If Trim(wsReport.Cells(i, "A").Value) <> "" Then
            DestRow = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).Row + 1
            wsMain.Cells(DestRow, "A").Value = wsReport.Cells(i, "A").Value
            wsMain.Cells(DestRow, "B").Value = wsReport.Cells(i, "B").Value
            wsMain.Cells(DestRow, "C").Value = wsReport.Cells(i, "C").Value
        End If
    Next i
End Sub
