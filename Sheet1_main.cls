Attribute VB_Name = "Sheet1_main"
Attribute VB_Description = "Лист main для ввода заказ-наряда"

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' обработка изменений
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub

    ' проверка ячейки на изменение B2
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    If Target.Address = "$B$2" Then
        ' очистка B3:B15
        If Target.Row >= 3 And Target.Row <= 15 Then Exit Sub

        isProcessing = True
        On Error Resume Next
        Dim b2Value As Variant
        b2Value = Me.Range("B2").Value
        If Not IsEmpty(b2Value) And b2Value <> "" Then
            Dim result As Boolean
            result = Mod_OrderHeader.FillHeaderFromOrder(CStr(b2Value))
        End If
        On Error GoTo 0
        isProcessing = False
    End If
End Sub
