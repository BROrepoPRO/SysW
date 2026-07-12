Attribute VB_Name = "Mod_Utils"
Option Explicit

' ============================================================
' Модуль: Mod_Utils
' Назначение: вспомогательные функции для работы с Excel и SQLite
' ============================================================

' Константы для SQLite
Public Const SQLITE_DLL As String = "C:\SQLite\sqlite3.dll"
Public Const DB_PATH As String = "C:\SQLite\SysW.db"

' Тип для хранения данных заказ-наряда
Public Type OrderHeader
    OrderNumber As String
    ClientName As String
    CarModel As String
    CarPlate As String
    Mileage As Long
    DateIn As Date
    DateOut As Date
    Status As String
End Type

' Функция: проверка существования файла
Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(FilePath)) > 0)
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
Public Sub WriteLog(ByVal Message As String)
    Dim LogPath As String
    Dim F As Long
    LogPath = ThisWorkbook.Path + "\log.txt"
    F = FreeFile
    Open LogPath For Append As #F
    Print #F, Now + " - " + Message
    Close #F
End Sub