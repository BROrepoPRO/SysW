' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' RunAllTests_UI
' Запускает все тесты (TC-01..TC-20) и показывает результат
' --------------------------------------------------------------------------
Public Sub RunAllTests_UI()
    On Error GoTo ErrHandler

    Call RunAllTests

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RunAllTests_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RunAllTests_UI: " & Err.Description)
End Sub