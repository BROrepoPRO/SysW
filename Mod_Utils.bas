Attribute VB_Name = "Mod_Utils"
Option Explicit

' ============================================================
' Модуль: Mod_Utils
' Назначение: Вспомогательные функции для работы с Excel
' ============================================================

' Тип для хранения данных заказа-наряда
' Поля соответствуют столбцам листа "spisok":
' A=№ п/п, B=Модель, C=ГРЗ, D=VIN, E=гараж.№, F=год вып., G=пробег, H=дата
Public Type OrderHeader
    OrderNumber As String    ' № п/п (колонка A)
    ModelName As String      ' Модель (колонка B)
    grz As String            ' ГРЗ/госномер (колонка C)
    VIN As String            ' VIN (колонка D)
    GarageNumber As String   ' Гаражный № (колонка E)
    YearMade As Integer      ' Год выпуска (колонка F)
    MileageValue As Long     ' Пробег (колонка G)
    DateValue As Date        ' Дата (колонка H)
End Type

' Функция: получение листа по имени (без ошибки если нет)
Public Function GetSheetByName(ByVal wb As Workbook, ByVal SheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = wb.Sheets(SheetName)
    On Error GoTo 0
End Function

' Функция: получение пути к книге
Public Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.path
End Function

' Функция: проверка существования файла
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(filePath)) > 0)
    On Error GoTo 0
End Function

' Функция: получение имени пользователя Windows
Public Function GetCurrentUser() As String
    GetCurrentUser = Environ("USERNAME")
End Function

' Функция: форматирование даты в формат SQLite (ГГГГ-ММ-ДД)
Public Function FormatDateSQL(ByVal d As Date) As String
    FormatDateSQL = Format(d, "yyyy-mm-dd")
End Function

' Процедура: запись в лог-файл
Public Sub WriteLog(ByVal message As String)
    Dim LogPath As String
    Dim F As Long
    LogPath = ThisWorkbook.path & "\log.txt"
    F = FreeFile
    Open LogPath For Append As #F
    Print #F, Now & " - " & message
    Close #F
End Sub

' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' WriteLog_UI
' Запрашивает сообщение через InputBox и записывает в лог
' --------------------------------------------------------------------------
Public Sub WriteLog_UI()
    On Error GoTo ErrHandler

    Dim msg As String
    msg = InputBox("Введите сообщение для записи в лог:", "Запись в лог")

    If msg = "" Then
        Exit Sub
    End If

    Call WriteLog("WriteLog_UI: " & msg)

    MsgBox "Сообщение записано в лог.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в WriteLog_UI: " & Err.Description, vbCritical, "Ошибка"
    Call WriteLog("WriteLog_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ShowWorkbookPath_UI
' Показывает путь к текущей книге через MsgBox
' --------------------------------------------------------------------------
Public Sub ShowWorkbookPath_UI()
    On Error GoTo ErrHandler

    Dim path As String
    path = GetWorkbookPath()

    MsgBox "Путь к книге:" & vbCrLf & path, vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ShowWorkbookPath_UI: " & Err.Description, vbCritical, "Ошибка"
    Call WriteLog("ShowWorkbookPath_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ShowCurrentUser_UI
' Показывает имя текущего пользователя Windows через MsgBox
' --------------------------------------------------------------------------
Public Sub ShowCurrentUser_UI()
    On Error GoTo ErrHandler

    Dim user As String
    user = GetCurrentUser()

    MsgBox "Текущий пользователь: " & user, vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ShowCurrentUser_UI: " & Err.Description, vbCritical, "Ошибка"
    Call WriteLog("ShowCurrentUser_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' CheckFileExists_UI
' Запрашивает путь к файлу через InputBox и проверяет его существование
' --------------------------------------------------------------------------
Public Sub CheckFileExists_UI()
    On Error GoTo ErrHandler

    Dim filePath As String
    Dim exists As Boolean

    filePath = InputBox("Введите полный путь к файлу:", "Проверка существования файла")

    If filePath = "" Then
        Exit Sub
    End If

    exists = FileExists(filePath)

    If exists Then
        MsgBox "Файл существует:" & vbCrLf & filePath, vbInformation, "SysW"
    Else
        MsgBox "Файл не найден:" & vbCrLf & filePath, vbExclamation, "SysW"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в CheckFileExists_UI: " & Err.Description, vbCritical, "Ошибка"
    Call WriteLog("CheckFileExists_UI: " & Err.Description)
End Sub
