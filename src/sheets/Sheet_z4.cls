VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_z4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист z4 для работы с запчастями"
Attribute VB_HelpID = 0
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

' ============================================================
' Класс: Sheet_z4
' Назначение: Обработчики событий листа z4 (запчасти)
' ============================================================

' --------------------------------------------------------------------------
' Worksheet_Activate
' Закрепление первых двух строк при активации листа.
' Проверяет, не установлен ли уже FreezePanes, чтобы не сбрасывать.
' --------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    ' Закрепление первых двух строк при активации листа
    If ActiveWindow.FreezePanes = False Then
        Rows("3").Select
        ActiveWindow.FreezePanes = True
        Range("A1").Select
    End If
End Sub