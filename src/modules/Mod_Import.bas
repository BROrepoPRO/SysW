Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================
' Модуль: Mod_Import
' Назначение: Импорт данных из Excel в SQLite и обратно
' ============================================================

' ============================================================
' ОСНОВНЫЕ ФУНКЦИИ ИМПОРТА
' ============================================================

' --------------------------------------------------------------------------
' ImportSheet
' Импортирует лист из report.xlsx по ГРЗ в текущую книгу
' --------------------------------------------------------------------------
Public Sub ImportSheet(grz As String)
    Dim wsSource As Worksheet
    Dim wsMain As Worksheet
    Dim newName As String

    Set wsMain = ThisWorkbook.Sheets("main")

    Set wsSource = Mod_SheetOps.SearchSheetByGRZ(grz)
    If wsSource Is Nothing Then
        MsgBox "Лист с ГРЗ " & grz & " не найден!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

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
    lastRowL = wsMain.Cells(wsMain.Rows.count, 12).End(xlUp).Row
    lastRowX = wsMain.Cells(wsMain.Rows.count, 24).End(xlUp).Row
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
        srcLastRow = wsSource.Cells(wsSource.Rows.count, 3).End(xlUp).Row

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
        srcLastRow = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).Row

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
' RenameSheets_UI
' Переименовывает листы в report.xlsx по ГРЗ
' --------------------------------------------------------------------------
Public Sub RenameSheets_UI()
    On Error GoTo ErrHandler

    Call Mod_SheetOps.RenameSheetsByGRZ

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
