VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Класс: WorkIdentity
' Назначение: Объектное представление тождества работ
' ============================================================

Public OutArticle As String   ' B — Артикул (модельный)
Public OutName As String      ' C — Наименование (модельное)
Public NormHours As Double    ' D — н/ч
Public QtyZN As Double        ' G — Кол-во ЗН
Public Aggregate As String    ' I — Агрегат (код)
Public InName As String       ' J — Наим-ние (входящее)