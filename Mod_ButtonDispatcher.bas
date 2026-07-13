Attribute VB_Name = "Mod_ButtonDispatcher"
Option Explicit

' ============================================================
' Модуль: Mod_ButtonDispatcher
' Назначение: Обработка нажатий на кнопки
' ============================================================

' Обработчик кнопки "Импорт"
Public Sub Btn_main_Import_Click()
    Mod_Import.RunImport
End Sub

' Обработчик кнопки "Тест"
Public Sub Btn_main_Test_Click()
    Mod_FullTestRunner.RunAllTests
End Sub

' Обработчик кнопки "Очистить"
Public Sub Btn_main_Clear_Click()
    Dim ws As Worksheet
    Dim Response As VbMsgBoxResult

    Set ws = ActiveSheet
    Response = MsgBox("Очистить все данные на листе " & ws.Name & "?", vbYesNo + vbQuestion, "Очистка")

    If Response = vbYes Then
        ws.Cells.ClearContents
        MsgBox "Данные очищены.", vbInformation, "Очистка"
    End If
End Sub

' Обработчик кнопки "Обновить"
Public Sub Btn_main_Refresh_Click()
    MsgBox "Обновление данных...", vbInformation, "Обновление"
End Sub

' Заглушка для z4
Public Sub Btn_z4_Action_Click()
    MsgBox "Функция z4 в разработке.", vbInformation, "z4"
End Sub

' Заглушка для work
Public Sub Btn_work_Action_Click()
    MsgBox "Функция work в разработке.", vbInformation, "Work"
End Sub
