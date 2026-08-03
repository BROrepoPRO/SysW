VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PartIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Класс: PartIdentity
' Назначение: Объектное представление тождества запчастей
' Заменяет Public Type PartIdentity в Mod_ModelTypes
' ============================================================

Public OutArticle As String   ' B — Артикул (модельный)
Public OutName As String      ' C — Наименование (модельное)
Public QtyZN As Double        ' G — Кол-во ЗН
Public Price As Currency      ' F — Цена за ед. изм.
Public Aggregate As String    ' I — АГРЕГАТ (код)
Public InCatNum As String     ' J — № кат. (входящий)
Public InName As String       ' K — Наим-ние (входящее)