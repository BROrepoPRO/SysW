Attribute VB_Name = "Mod_ButtonDispatcher"
Option Explicit

' ============================================================
' Модуль: Mod_ButtonDispatcher
' Назначение: диспетчер кнопок на листах
' ============================================================

' Обработчик нажатия кнопки "Импорт"
Public Sub btnImport_Click()
    Mod_Import.RunImport
End Sub

' Обработчик нажатия кнопки "Тест"
Public Sub btnTest_Click()
    Mod_MinimalTestRunner.RunAllTests
End Sub

' Обработчик нажатия кнопки "Очистить"
Public Sub btnClear_Click()
    Dim ws As Worksheet
    Dim Response As VbMsgBoxResult
    
    Set ws = ActiveSheet
    Response = MsgBox("Очистить все данные на листе " & ws.Name & "?", vbYesNo + vbQuestion, "Очистка")
    
    If Response = vbYes Then
        ws.Cells.ClearContents
        MsgBox "Данные очищены.", vbInformation, "Очистка"
    End If
End Sub

' Обработчик нажатия кнопки "Обновить"
Public Sub btnRefresh_Click()
    MsgBox "Обновление данных из SQLite...", vbInformation, "Обновление"
    ' Здесь будет вызов загрузки из SQLite
End Sub