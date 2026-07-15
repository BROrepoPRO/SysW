Attribute VB_Name = "Mod_SheetOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Модуль: Mod_SheetOps
' Назначение: Операции с листами Excel (поиск, переименование,
'             очистка, работа с ГРЗ)
' ============================================================

' --------------------------------------------------------------------------
' ExtractNumberFromGRZ
' Извлекает цифровую группу длиной 3 или 4 цифры из строки ГРЗ
' Пример: "А123АН77" -> "123", "А12АН34" -> "" (только 2 цифры)
' --------------------------------------------------------------------------
Public Function ExtractNumberFromGRZ(grz As String) As String
    Dim i As Long
    Dim currentDigits As String
    Dim Result As String

    Result = ""
    currentDigits = ""

    For i = 1 To Len(grz)
        If Mid(grz, i, 1) >= "0" And Mid(grz, i, 1) <= "9" Then
            currentDigits = currentDigits & Mid(grz, i, 1)
        Else
            If Len(currentDigits) = 3 Or Len(currentDigits) = 4 Then
                Result = currentDigits
                Exit For
            End If
            currentDigits = ""
        End If
    Next i

    ' Проверка в конце строки
    If Result = "" And (Len(currentDigits) = 3 Or Len(currentDigits) = 4) Then
        Result = currentDigits
    End If

    ExtractNumberFromGRZ = Result
End Function

' --------------------------------------------------------------------------
' SearchSheetByGRZ
' Открывает report.xlsx (если не открыт), ищет лист по номеру ГРЗ
' Возвращает найденный лист или Nothing, если не найден
' --------------------------------------------------------------------------
Public Function SearchSheetByGRZ(grz As String) As Worksheet
    Dim wbReport As Workbook
    Dim ws As Worksheet
    Dim grzNumber As String
    Dim wsResult As Worksheet

    Set wsResult = Nothing
    grzNumber = Mod_SheetOps.ExtractNumberFromGRZ(grz)

    If grzNumber = "" Then
        Set SearchSheetByGRZ = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set wbReport = Workbooks.Open(ThisWorkbook.path & "\report.xlsx", ReadOnly:=True)
    On Error GoTo 0

    If wbReport Is Nothing Then
        MsgBox "Не удалось открыть report.xlsx!", vbExclamation, "Ошибка"
        Set SearchSheetByGRZ = Nothing
        Exit Function
    End If

    On Error GoTo CleanUp

    For Each ws In wbReport.Sheets
        If InStr(1, ws.Name, grzNumber, vbTextCompare) > 0 Then
            Set wsResult = ws
            Exit For
        End If
    Next ws

CleanUp:
    If Not wbReport Is Nothing Then
        wbReport.Close SaveChanges:=False
    End If
    Set SearchSheetByGRZ = wsResult
End Function

' --------------------------------------------------------------------------
' RenameSheetsByGRZ
' Переименовывает листы в report.xlsx по номеру ГРЗ
' --------------------------------------------------------------------------
Public Sub RenameSheetsByGRZ()
    Dim wbReport As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim grzNumber As String
    Dim existingWs As Worksheet

    On Error Resume Next
    Set wbReport = Workbooks.Open(ThisWorkbook.path & "\report.xlsx", ReadOnly:=False)
    On Error GoTo 0

    If wbReport Is Nothing Then
        MsgBox "Не удалось открыть report.xlsx!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    For Each ws In wbReport.Sheets
        If ws.Visible = xlSheetVisible And ws.Name <> "report" And ws.Name <> "spisok" Then
            ' Ищем первую сверху ячейку со словом "автомобиль"
            Set cell = ws.Cells.Find(What:="автомобиль", LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            If Not cell Is Nothing Then
                grzNumber = Mod_SheetOps.ExtractNumberFromGRZ(cell.Value)
                If grzNumber <> "" Then
                    ' Если лист с таким именем уже существует, удаляем
                    On Error Resume Next
                    Set existingWs = wbReport.Sheets(grzNumber)
                    If Not existingWs Is Nothing Then
                        existingWs.Delete
                    End If
                    On Error GoTo 0

                    ws.Name = grzNumber
                End If
            End If
        End If
    Next ws

    wbReport.Save
    wbReport.Close
End Sub

' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' ClearMainSheet_UI
' Очищает все данные на листе main с подтверждением пользователя
' --------------------------------------------------------------------------
Public Sub ClearMainSheet_UI()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim response As VbMsgBoxResult

    Set wsMain = ThisWorkbook.Sheets("main")

    response = MsgBox("Очистить все данные на листе main?", vbYesNo + vbQuestion, "Подтверждение")

    If response = vbYes Then
        lastRow = wsMain.UsedRange.Rows.count
        If lastRow >= 2 Then
            wsMain.Range("B2:ZZ" & lastRow).ClearContents
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ClearMainSheet_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ClearMainSheet_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ClearHeader_UI
' Очищает шапку заказа B3:B15 на листе main
' --------------------------------------------------------------------------
Public Sub ClearHeader_UI()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Sheets("main")

    wsMain.Range("B3:B15").ClearContents

    MsgBox "Шапка заказа (B3:B15) очищена.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ClearHeader_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ClearHeader_UI: " & Err.Description)
End Sub