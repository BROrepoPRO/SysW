VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист main для ввода заказ-наряда"

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' обработка изменений
    On Error GoTo ErrHandler

    Static isProcessing As Boolean
    If isProcessing Then Exit Sub

    ' проверка ячейки на изменение B2
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range("B2")) Is Nothing Then Exit Sub
    If Target.Address = "$B$2" Then
        ' очистка B3:B15
        If Target.Row >= 3 And Target.Row <= 15 Then Exit Sub

        isProcessing = True

        Dim b2Value As Variant
        b2Value = Me.Range("B2").Value
        If Not IsEmpty(b2Value) And b2Value <> "" Then
            Dim result As Boolean
            result = Mod_OrderHeader.FillHeaderFromOrder(CStr(b2Value))
        End If

        isProcessing = False
    End If

    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения и сброс флага
    Application.EnableEvents = True
    isProcessing = False
    Call Mod_Logger.WriteLog("Sheet1_main", "Worksheet_Change: " & Err.Description)
End Sub
