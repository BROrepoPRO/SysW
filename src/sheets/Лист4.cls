VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Лист models (модели, группы, цены н/ч)"

Option Explicit

Private Sub Worksheet_Activate()
    ' Закрепление строк 1-3 (FreezePanes A4) при активации листа models
    Call Mod_SheetOps.ApplyFreezePanes(Me)
End Sub