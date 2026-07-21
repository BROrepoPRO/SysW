Attribute VB_Name = "Mod_ModelDB"
Option Explicit

' ============================================================
' Модуль: Mod_ModelDB
' Назначение: Доступ к файлам модельных групп
'             (базовый слой абстракции для работы с base/models/)
' ============================================================

' ============================================================
' Константы
' ============================================================
Public Const MODELDB_BASE_PATH As String = "L:\PROject\SysW\base\models\"

' ============================================================
' Типы данных
' ============================================================

' --------------------------------------------------------------------------
' WorkEntry
' Структура записи работы из листа {GroupName} файла группы
' --------------------------------------------------------------------------
Public Type WorkEntry
    Code As String
    Name As String
    Unit As String
    NormHours As Double
    Price As Currency
    Note As String
End Type

' ============================================================
' Вспомогательные функции
' ============================================================

' --------------------------------------------------------------------------
' GetModelGroupFilePath
' Возвращает полный путь к файлу группы
' --------------------------------------------------------------------------
Public Function GetModelGroupFilePath(ByVal groupName As String) As String
    GetModelGroupFilePath = MODELDB_BASE_PATH & groupName & ".xlsx"
End Function

' --------------------------------------------------------------------------
' ModelGroupFileExists
' Проверяет существование файла группы
' --------------------------------------------------------------------------
Public Function ModelGroupFileExists(ByVal groupName As String) As Boolean
    Dim filePath As String
    filePath = GetModelGroupFilePath(groupName)
    ModelGroupFileExists = (Len(Dir(filePath)) > 0)
End Function

' ============================================================
' Основные функции
' ============================================================

' --------------------------------------------------------------------------
' OpenModelGroupFile
' Открывает файл группы {groupName}.xlsx (если ещё не открыт)
' и возвращает ссылку на Workbook.
' Если файл не найден — возвращает Nothing.
' --------------------------------------------------------------------------
Public Function OpenModelGroupFile(ByVal groupName As String) As Workbook
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim filePath As String
    Dim wbName As String

    wbName = groupName & ".xlsx"

    ' 1. Проверяем, не открыт ли уже файл
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo ErrHandler

    If Not wb Is Nothing Then
        ' Файл уже открыт — возвращаем ссылку
        Set OpenModelGroupFile = wb
        Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Файл " & wbName & " уже открыт")
        Exit Function
    End If

    ' 2. Проверяем существование файла
    filePath = GetModelGroupFilePath(groupName)
    If Not Mod_Utils.FileExists(filePath) Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Файл не найден — " & filePath)
        Set OpenModelGroupFile = Nothing
        Exit Function
    End If

    ' 3. Открываем файл
    Set wb = Workbooks.Open(filePath, ReadOnly:=False)
    Set OpenModelGroupFile = wb

    Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Открыт файл " & filePath)
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Ошибка — " & Err.Description)
    Set OpenModelGroupFile = Nothing
End Function

' --------------------------------------------------------------------------
' GetWorks
' Возвращает коллекцию работ из листа {groupName} файла группы
' с применением фильтров.
' На данном этапе — заглушка для будущего использования.
' --------------------------------------------------------------------------
Public Function GetWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim entry As WorkEntry

    Set result = New Collection

    ' Открываем файл группы
    Set wb = OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set GetWorks = result
        Exit Function
    End If

    ' Получаем лист {groupName}
    On Error Resume Next
    Set ws = wb.Sheets(groupName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorks: Лист " & groupName & " не найден в файле " & groupName & ".xlsx")
        Set GetWorks = result
        Exit Function
    End If

    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 4 Then
        Set GetWorks = result
        Exit Function
    End If

    ' Читаем данные (столбцы: A=Code, B=Name, C=Unit, D=NormHours, E=Price, F=Note)
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            entry.Code = CStr(ws.Cells(i, 1).Value)
            entry.Name = CStr(ws.Cells(i, 2).Value)
            entry.Unit = CStr(ws.Cells(i, 3).Value)
            entry.NormHours = Val(ws.Cells(i, 4).Value)
            entry.Price = Val(ws.Cells(i, 5).Value)
            entry.Note = CStr(ws.Cells(i, 6).Value)
            result.Add entry
        End If
    Next i

    Set GetWorks = result
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorks: Ошибка — " & Err.Description)
    Set GetWorks = New Collection
End Function