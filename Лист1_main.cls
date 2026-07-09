VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист main — шапка заказ-наряда"
Attribute VB_HelpID = 0
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
' ============================================================================
' Код листа main
' Назначение: Обработчик изменения ячейки B2 для автозаполнения шапки
' ============================================================================

Option Explicit

' --------------------------------------------------------------------------
' Обработчик события Worksheet_Change
' Отслеживает изменение ячейки B2 и запускает заполнение шапки заказ-наряда
' --------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Если события уже отключены (защита от рекурсии)
    If Not Application.EnableEvents Then Exit Sub
    ' Только одна ячейка
    If Target.CountLarge <> 1 Then Exit Sub
    ' Только ячейка B2
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    ' Игнорируем изменения в B3:B15 (чтобы избежать рекурсии)
    If Not Intersect(Target, Me.Range("B3:B15")) Is Nothing Then Exit Sub

    On Error GoTo ВосстановитьСобытия
    Application.EnableEvents = False
    Mod_OrderHeader.FillHeaderFromOrder Target.Value, Me, ThisWorkbook.Sheets("spisok"), ThisWorkbook.Sheets("model")

ВосстановитьСобытия:
    Application.EnableEvents = True
End Sub