VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист main с формой заказа-наряда"
Attribute VB_HelpID = 0
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ============================================================================
' Лист main
' Назначение: Обработка изменения ячейки B2 для автозаполнения полей
' ============================================================================

Option Explicit

' --------------------------------------------------------------------------
' Событие Worksheet_Change
' Вызывает FillHeaderFromOrder при изменении ячейки B2
' --------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Защита от рекурсивного вызова
    Static IsProcessing As Boolean
    If IsProcessing Then Exit Sub
    
    ' Проверка флага (сброс после ошибки)
    If Not Application.EnableEvents Then Exit Sub
    ' Только одна ячейка
    If Target.CountLarge <> 1 Then Exit Sub
    ' Только ячейка B2
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    ' Исключаем B3:B15 (чтобы не было цикла при автозаполнении)
    If Not Intersect(Target, Me.Range("B3:B15")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    IsProcessing = True
    Application.EnableEvents = False
    
    Dim wsSpisok As Worksheet
    Dim wsModel As Worksheet
    
    Set wsSpisok = GetSheetByName(ThisWorkbook, "spisok")
    Set wsModel = GetSheetByName(ThisWorkbook, "model")
    
    Mod_OrderHeader.FillHeaderFromOrder Trim(Target.Value), Me, wsSpisok, wsModel

CleanUp:
    Application.EnableEvents = True
    IsProcessing = False
    Exit Sub

ErrorHandler:
    Resume CleanUp
End Sub
