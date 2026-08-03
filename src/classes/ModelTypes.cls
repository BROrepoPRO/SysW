Attribute VB_Name = "ModelTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =============================================
' Модельные типы данных
' Перенесены из Mod_ModelDB.bas для исправления
' ошибки компиляции "Only public user defined types..."
' =============================================

' --------------------------------------------------------------------------
' WorkEntry
' Структура записи работы из листа {GroupName} файла группы
' --------------------------------------------------------------------------
Public Type WorkEntry
    Code As String
    Name As String
    Unit As String
    NormHours As Double
    Price As Currency
    Note As String
End Type

' --------------------------------------------------------------------------
' WorkIdentity
' Структура записи тождества работ из листа {GroupName}w
' Соответствует колонкам UAZw (A-O):
'   OutArticle(B), OutName(C), NormHours(D), QtyZN(G),
'   Aggregate(I), InName(J)
' --------------------------------------------------------------------------
Public Type WorkIdentity
    OutArticle As String   ' B — Артикул (модельный)
    OutName As String      ' C — Наименование (модельное)
    NormHours As Double    ' D — н/ч
    QtyZN As Double        ' G — Кол-во ЗН
    Aggregate As String    ' I — Агрегат (код)
    InName As String       ' J — Наим-ние (входящее)
End Type

' --------------------------------------------------------------------------
' PartIdentity
' Структура записи тождества запчастей из листа {GroupName}z4
' Соответствует колонкам UAZz4 (A-P):
'   OutArticle(B), OutName(C), QtyZN(G), Price(F),
'   Aggregate(I), InCatNum(J), InName(K)
' --------------------------------------------------------------------------
Public Type PartIdentity
    OutArticle As String   ' B — Артикул (модельный)
    OutName As String      ' C — Наименование (модельное)
    QtyZN As Double        ' G — Кол-во ЗН
    Price As Currency      ' F — Цена за ед. изм.
    Aggregate As String    ' I — АГРЕГАТ (код)
    InCatNum As String     ' J — № кат. (входящий)
    InName As String       ' K — Наим-ние (входящее)
End Type