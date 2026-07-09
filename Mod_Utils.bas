Attribute VB_Name = "Mod_Utils"
Option Explicit

' ============================================================================
' Модуль: Mod_Utils
' Назначение: Общие вспомогательные функции для всех модулей проекта
' Автор: SourceCraft
' Дата: 2026-07-09
' ============================================================================

' --------------------------------------------------------------------------
' Функция: GetSheetByName
' Возвращает лист по имени или Nothing, если лист не найден.
' --------------------------------------------------------------------------
Public Function GetSheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then Set GetSheetByName = Nothing
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' Функция: GetWorkbookPath
' Возвращает путь к книге, где хранится макрос, с завершающим "\".
' --------------------------------------------------------------------------
Public Function GetWorkbookPath() As String
    Dim result As String
    result = ThisWorkbook.Path
    If Right(result, 1) <> "\" Then result = result & "\"
    GetWorkbookPath = result
End Function

' --------------------------------------------------------------------------
' Функция: FileExists
' Проверяет существование файла по указанному пути.
' Обёртка над Scripting.FileSystemObject.
' --------------------------------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function