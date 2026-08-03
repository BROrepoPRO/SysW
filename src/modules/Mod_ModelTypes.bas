Attribute VB_Name = "Mod_ModelTypes"
Option Explicit

' =============================================
' Модуль: Mod_ModelTypes
' Назначение: Определения пользовательских типов (UDT)
' Перенесены из ModelTypes.cls для исправления ошибки компиляции
' "Only public user defined types defined in public object modules
'  can be used as parameters or return types for public procedures
'  of class modules or as fields of public user defined types"
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
' WorkIdentity и PartIdentity перенесены в классы:
'   src/classes/WorkIdentity.cls
'   src/classes/PartIdentity.cls
' --------------------------------------------------------------------------