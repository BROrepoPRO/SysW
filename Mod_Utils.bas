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
    GRZ As String            ' ГРЗ/госномер (колонка C)
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
    GetWorkbookPath = ThisWorkbook.Path
End Function

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
    LogPath = ThisWorkbook.Path & "\log.txt"
    F = FreeFile
    Open LogPath For Append As #F
    Print #F, Now & " - " & Message
    Close #F
End Sub