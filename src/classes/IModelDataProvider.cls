VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IModelDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Класс: IModelDataProvider (интерфейс-контракт)
' Назначение: контракт провайдера данных моделей.
' Реализации: Mod_ModelDBProvider (Excel), Mod_SQLiteDB (SQLite).
' Все методы возвращают коллекции ОБЪЕКТОВ (или массивов для частей).
' Заглушки Err.Raise не должны вызываться напрямую — реализация
' обязана переопределить каждый метод через Implements.
' ============================================================

Public Function GetWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetWorks: Not implemented"
End Function

Public Function GetParts(ByVal groupName As String, ByRef filters As Variant) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetParts: Not implemented"
End Function

Public Function GetModelWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetModelWorks: Not implemented"
End Function

Public Function GetModelParts(ByVal groupName As String, ByRef filters As Variant) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetModelParts: Not implemented"
End Function

Public Function GetMatLibEntries(ByVal groupName As String, ByVal entryCode As String) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetMatLibEntries: Not implemented"
End Function

Public Function GetWorkIdentities(ByVal groupName As String) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetWorkIdentities: Not implemented"
End Function

Public Function GetPartIdentities(ByVal groupName As String) As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetPartIdentities: Not implemented"
End Function

Public Function GetAllModelGroups() As Collection
    Err.Raise vbObjectError + 1, "IModelDataProvider", "GetAllModelGroups: Not implemented"
End Function

Public Function CreateModelGroupFile(ByVal groupName As String) As Boolean
    Err.Raise vbObjectError + 1, "IModelDataProvider", "CreateModelGroupFile: Not implemented"
End Function

Public Function ModelGroupFileExists(ByVal groupName As String) As Boolean
    Err.Raise vbObjectError + 1, "IModelDataProvider", "ModelGroupFileExists: Not implemented"
End Function