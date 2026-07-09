VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист main в книге заказ-наряда"
Attribute VB_HelpID = 0
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ============================================================================
' Класс main
' Назначение: Обработка изменения ячейки B2 для автозаполнения шапки
' ============================================================================

Option Explicit

' --------------------------------------------------------------------------
' Событие Worksheet_Change
' Запускает FillHeaderFromOrder при изменении ячейки B2
' --------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Защита от рекурсивного вызова
    Static IsProcessing As Boolean
    If IsProcessing Then Exit Sub
    
    ' Отключаем события (чтобы не было цикла)
    If Not Application.EnableEvents Then Exit Sub
    ' Только одна ячейка
    If Target.CountLarge <> 1 Then Exit Sub
    ' Только ячейка B2
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    ' Исключаем B3:B15 (чтобы не было цикла при заполнении)
    If Not Intersect(Target, Me.Range("B3:B15")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    IsProcessing = True
    Application.EnableEvents = False
    Mod_OrderHeader.FillHeaderFromOrder Target.Value, Me, ThisWorkbook.Sheets("spisok"), ThisWorkbook.Sheets("model")

CleanUp:
    Application.EnableEvents = True
    IsProcessing = False
    Exit Sub

ErrorHandler:
    Resume CleanUp
End Sub