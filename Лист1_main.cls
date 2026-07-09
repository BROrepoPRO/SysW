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

    ' Проверяем, что изменение затронуло только одну ячейку
    If Target.CountLarge <> 1 Then
        Exit Sub
    End If

    ' Проверяем, что изменённая ячейка — это B2
    If Intersect(Target, Me.Range("B2")) Is Nothing Then
        Exit Sub
    End If

    ' Отключаем события для предотвращения рекурсивных вызовов
    Application.EnableEvents = False

    ' Защита от сбоя: если произойдёт ошибка, события всё равно включатся
    On Error GoTo ВосстановитьСобытия

    ' Вызываем функцию заполнения шапки
    Dim результат As Boolean
    результат = Mod_OrderHeader.FillHeaderFromOrder( _
        Target.Value, _
        Me, _
        ThisWorkbook.Sheets("spisok"), _
        ThisWorkbook.Sheets("model"))

ВосстановитьСобытия:
    ' Включаем события обратно в любом случае (при ошибке или успехе)
    Application.EnableEvents = True

End Sub