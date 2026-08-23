VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист spisok (список автомобилей)"

Option Explicit

Private Sub Worksheet_Activate()
    ' Закрепление строк 1-3 (FreezePanes A4) при активации листа spisok
    Call Mod_SheetOps.ApplyFreezePanes(Me)
End Sub