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

Private Sub Worksheet_Activate()
    ' Закрепление строк 1-3 (FreezePanes A4) при активации листа main
    Call Mod_SheetOps.ApplyFreezePanes(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' обработка изменений
    On Error GoTo ErrHandler

    Static isProcessing As Boolean
    If isProcessing Then Exit Sub

    ' проверка ячейки на изменение B4
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range("B4")) Is Nothing Then Exit Sub
    If Target.Address = "$B$4" Then
        isProcessing = True
        Application.EnableEvents = False  ' отключаем события

        Dim b2Value As Variant
        b2Value = Me.Range("B4").Value
        If Not IsEmpty(b2Value) And b2Value <> "" Then
            Dim result As Boolean
            result = Mod_OrderHeader.FillHeaderFromOrder(b2Value)
        End If

        Application.EnableEvents = True
        isProcessing = False
    End If

    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения и сброс флага
    Application.EnableEvents = True
    isProcessing = False
    Call Mod_Logger.WriteLog("Sheet1_main", "Worksheet_Change: " & Err.Description)
End Sub
