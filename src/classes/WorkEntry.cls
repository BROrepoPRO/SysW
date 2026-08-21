VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Класс: WorkEntry
' Назначение: Объектное представление записи работы из каталога
'             группы (лист {GroupName}). Заменяет UDT WorkEntry
'             в Mod_ModelTypes для единообразного возврата из
'             методов IModelDataProvider.
' ============================================================

Public Code As String        ' B — Артикул
Public Name As String         ' C — Наименование
Public Unit As String         ' Ед. изм. (резерв)
Public NormHours As Double    ' D — н/ч
Public Price As Currency      ' F — цена н/ч
Public Note As String         ' примечание (резерв)