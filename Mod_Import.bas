Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================
' Модуль: Mod_Import
' Назначение: Импорт данных из Excel в SQLite и обратно
' ============================================================

' --------------------------------------------------------------------------
' ExtractNumberFromGRZ
' Извлекает цифровую группу длиной 3 или 4 цифры из строки ГРЗ
' Пример: "А123АН77" -> "123", "А12АН34" -> "" (только 2 цифры)
' --------------------------------------------------------------------------
Public Function ExtractNumberFromGRZ(GRZ As String) As String
    Dim i As Long
    Dim currentDigits As String
    Dim result As String

    result = ""
    currentDigits = ""

    For i = 1 To Len(GRZ)
        If Mid(GRZ, i, 1) >= "0" And Mid(GRZ, i, 1) <= "9" Then
            currentDigits = currentDigits & Mid(GRZ, i, 1)
        Else
            If Len(currentDigits) = 3 Or Len(currentDigits) = 4 Then
                result = currentDigits
                Exit For
            End If
            currentDigits = ""
        End If
    Next i

    ' Проверка в конце строки
    If result = "" And (Len(currentDigits) = 3 Or Len(currentDigits) = 4) Then
        result = currentDigits
    End If

    ExtractNumberFromGRZ = result
End Function

' --------------------------------------------------------------------------
' SearchSheetByGRZ
' Открывает report.xlsx (если не открыт), ищет лист по номеру ГРЗ
' Возвращает найденный лист или Nothing, если не найден
' --------------------------------------------------------------------------
Public Function SearchSheetByGRZ(GRZ As String) As Worksheet
    Dim wbReport As Workbook
    Dim ws As Worksheet
    Dim grzNumber As String
    Dim wsResult As Worksheet

    Set wsResult = Nothing
    grzNumber = ExtractNumberFromGRZ(GRZ)

    If grzNumber = "" Then
        Set SearchSheetByGRZ = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set wbReport = Workbooks.Open(ThisWorkbook.Path & "\report.xlsx", ReadOnly:=True)
    On Error GoTo 0

    If wbReport Is Nothing Then
        MsgBox "Не удалось открыть report.xlsx!", vbExclamation, "Ошибка"
        Set SearchSheetByGRZ = Nothing
        Exit Function
    End If

    For Each ws In wbReport.Sheets
        If InStr(1, ws.Name, grzNumber, vbTextCompare) > 0 Then
            Set wsResult = ws
            Exit For
        End If
    Next ws

    Set SearchSheetByGRZ = wsResult
    ' Книгу не закрываем
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
    Set wbReport = Workbooks.Open(ThisWorkbook.Path & "\report.xlsx", ReadOnly:=False)
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
                grzNumber = ExtractNumberFromGRZ(cell.Value)
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

' --------------------------------------------------------------------------
' ImportSheet
' Импортирует лист из report.xlsx по ГРЗ в текущую книгу
' --------------------------------------------------------------------------
Public Sub ImportSheet(GRZ As String)
    Dim wsSource As Worksheet
    Dim wsMain As Worksheet
    Dim newName As String

    Set wsMain = ThisWorkbook.Sheets("main")

    Set wsSource = SearchSheetByGRZ(GRZ)
    If wsSource Is Nothing Then
        MsgBox "Лист с ГРЗ " & GRZ & " не найден!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    newName = wsMain.Range("B2").Value & "M"
    On Error Resume Next
    ActiveSheet.Name = newName
    On Error GoTo 0

    Call ImportDataToMain(ActiveSheet)
End Sub

' --------------------------------------------------------------------------
' ImportDataToMain
' Переносит данные из листа-источника в лист main по столбцам
' Ищет таблицы "Выполненные работы" и "Расходная накладная" на листе
' --------------------------------------------------------------------------
Public Sub ImportDataToMain(wsSource As Worksheet)
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim srcLastRow As Long
    Dim foundWorks As Boolean
    Dim foundMaterials As Boolean
    Dim cell As Range
    Dim startRow As Long

    Set wsMain = ThisWorkbook.Sheets("main")

    ' Очистка диапазонов L:N и X:AA
    Dim lastRowL As Long, lastRowX As Long
    lastRowL = wsMain.Cells(wsMain.Rows.Count, 12).End(xlUp).Row
    lastRowX = wsMain.Cells(wsMain.Rows.Count, 24).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowL, lastRowX)
    If lastRow < 2 Then lastRow = 2

    wsMain.Range("L2:N" & lastRow).ClearContents
    wsMain.Range("X2:AA" & lastRow).ClearContents

    foundWorks = False
    foundMaterials = False

    ' Поиск таблицы "Выполненные работы"
    Set cell = wsSource.Cells.Find(What:="Выполненные работы", LookAt:=xlPart, SearchOrder:=xlByRows)
    If cell Is Nothing Then
        Set cell = wsSource.Cells.Find(What:="Наименование", LookAt:=xlPart, SearchOrder:=xlByRows)
    End If

    If Not cell Is Nothing Then
        foundWorks = True
        startRow = cell.Row + 1
        srcLastRow = wsSource.Cells(wsSource.Rows.Count, 3).End(xlUp).Row

        For i = startRow To srcLastRow
            If wsSource.Cells(i, 3).Value <> "" Then
                wsMain.Cells(i - startRow + 2, 12).Value = wsSource.Cells(i, 3).Value ' C -> L
                wsMain.Cells(i - startRow + 2, 13).Value = wsSource.Cells(i, 4).Value ' D -> M
                wsMain.Cells(i - startRow + 2, 14).Value = wsSource.Cells(i, 8).Value ' H -> N
            End If
        Next i
    End If

    ' Поиск таблицы "Расходная накладная"
    Set cell = wsSource.Cells.Find(What:="Расходная накладная", LookAt:=xlPart, SearchOrder:=xlByRows)

    If Not cell Is Nothing Then
        foundMaterials = True
        startRow = cell.Row + 1
        srcLastRow = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).Row

        For i = startRow To srcLastRow
            If wsSource.Cells(i, 2).Value <> "" Then
                wsMain.Cells(i - startRow + 2, 24).Value = wsSource.Cells(i, 2).Value ' B -> X
                wsMain.Cells(i - startRow + 2, 25).Value = wsSource.Cells(i, 3).Value ' C -> Y
                wsMain.Cells(i - startRow + 2, 26).Value = wsSource.Cells(i, 4).Value ' D -> Z
                wsMain.Cells(i - startRow + 2, 27).Value = wsSource.Cells(i, 7).Value ' G -> AA
            End If
        Next i
    End If

    If Not foundWorks Then
        MsgBox "Таблица 'Выполненные работы' не найдена!", vbExclamation, "Предупреждение"
    End If

    If Not foundMaterials Then
        MsgBox "Таблица 'Расходная накладная' не найдена!", vbExclamation, "Предупреждение"
    End If
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
        lastRow = wsMain.UsedRange.Rows.Count
        If lastRow >= 2 Then
            wsMain.Range("A2:XFD" & lastRow).ClearContents
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ClearMainSheet_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ClearMainSheet_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportSheet_UI
' Импортирует лист из report.xlsx по ГРЗ из ячейки B4 листа main
' --------------------------------------------------------------------------
Public Sub ImportSheet_UI()
    On Error GoTo ErrHandler

    Call ImportSheet(ThisWorkbook.Sheets("main").Range("B4").Value)

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportSheet_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportSheet_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportByInput_UI
' Запрашивает ГРЗ через InputBox, вызывает ImportSheet
' --------------------------------------------------------------------------
Public Sub ImportByInput_UI()
    On Error GoTo ErrHandler

    Dim grz As String
    grz = InputBox("Введите ГРЗ для импорта:", "Импорт по ГРЗ")

    If grz = "" Then
        Exit Sub
    End If

    Call ImportSheet(grz)

    MsgBox "Импорт по ГРЗ '" & grz & "' выполнен.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportByInput_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportByInput_UI: " & Err.Description)
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

' --------------------------------------------------------------------------
' RenameSheets_UI
' Переименовывает листы в report.xlsx по ГРЗ
' --------------------------------------------------------------------------
Public Sub RenameSheets_UI()
    On Error GoTo ErrHandler

    Call RenameSheetsByGRZ

    MsgBox "Переименование листов выполнено.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RenameSheets_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RenameSheets_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportDataToMain_UI
' Переносит данные с активного листа в лист main
' --------------------------------------------------------------------------
Public Sub ImportDataToMain_UI()
    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Set wsSource = ActiveSheet

    If wsSource Is Nothing Then
        MsgBox "Нет активного листа!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    If wsSource.Name = "main" Then
        MsgBox "Активный лист не может быть main. Выберите другой лист.", vbExclamation, "Предупреждение"
        Exit Sub
    End If

    Call ImportDataToMain(wsSource)

    MsgBox "Данные с листа '" & wsSource.Name & "' перенесены в main.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportDataToMain_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportDataToMain_UI: " & Err.Description)
End Sub
