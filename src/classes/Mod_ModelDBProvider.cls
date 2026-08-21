VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mod_ModelDBProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IModelDataProvider

' ============================================================
' Класс: Mod_ModelDBProvider
' Назначение: РЕЗЕРВНАЯ (Excel) реализация интерфейса IModelDataProvider.
'             Читает модельные данные из легаси-файлов base/models/*.xlsm
'             через живой Workbook (OpenModelGroupFile).
'             Используется фабрикой GetModelDataProvider() как fallback,
'             когда SysW.db или ODBC-драйвер недоступны.
'
' Основной провайдер — Mod_SQLiteDB (SQLite/ADO).
'
' ВАЖНО: класс самодостаточен и не вызывает Mod_ModelDB.* методы чтения
' тождеств (они являются делегатами к провайдеру — была бы рекурсия).
' Реализация тождеств выполняется прямым чтением листов {G}w / {G}z4.
' ============================================================

' --------------------------------------------------------------------------
' Вспомогательное: чтение тождеств работ из листа {groupName}w
' --------------------------------------------------------------------------
Private Function ReadWorkIdentitiesFromSheet(ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection

    If ws Is Nothing Then
        Set ReadWorkIdentitiesFromSheet = result
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row   ' B — Артикул
    If lastRow < 4 Then
        Set ReadWorkIdentitiesFromSheet = result
        Exit Function
    End If

    Dim i As Long
    For i = 4 To lastRow
        ' Критерий тождества: не пустой B (OutArticle) И не пустой I (Агрегат)
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then
                Dim wi As WorkIdentity
                Set wi = New WorkIdentity
                wi.OutArticle = CStr(ws.Cells(i, 2).Value)   ' B
                wi.OutName = CStr(ws.Cells(i, 3).Value)      ' C
                wi.NormHours = Val(ws.Cells(i, 4).Value)     ' D
                wi.QtyZN = Val(ws.Cells(i, 7).Value)         ' G
                wi.Aggregate = CStr(ws.Cells(i, 9).Value)    ' I
                wi.InName = CStr(ws.Cells(i, 10).Value)      ' J
                result.Add wi
            End If
        End If
    Next i

    Set ReadWorkIdentitiesFromSheet = result
End Function

' --------------------------------------------------------------------------
' Вспомогательное: чтение тождеств запчастей из листа {groupName}z4
' --------------------------------------------------------------------------
Private Function ReadPartIdentitiesFromSheet(ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection

    If ws Is Nothing Then
        Set ReadPartIdentitiesFromSheet = result
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row   ' B — Артикул
    If lastRow < 4 Then
        Set ReadPartIdentitiesFromSheet = result
        Exit Function
    End If

    Dim i As Long
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then
                Dim pi As PartIdentity
                Set pi = New PartIdentity
                pi.OutArticle = CStr(ws.Cells(i, 2).Value)   ' B
                pi.OutName = CStr(ws.Cells(i, 3).Value)      ' C
                pi.QtyZN = Val(ws.Cells(i, 7).Value)         ' G
                pi.Price = Val(ws.Cells(i, 6).Value)         ' F
                pi.Aggregate = CStr(ws.Cells(i, 9).Value)    ' I
                pi.InCatNum = CStr(ws.Cells(i, 10).Value)    ' J
                pi.InName = CStr(ws.Cells(i, 11).Value)      ' K
                result.Add pi
            End If
        End If
    Next i

    Set ReadPartIdentitiesFromSheet = result
End Function

' --------------------------------------------------------------------------
' GetWorks — читает лист {groupName} и возвращает Collection объектов WorkEntry
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetWorks(ByVal groupName As String, _
                                             ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetWorks = result
        Exit Function
    End If

    On Error Resume Next
    Set ws = wb.Sheets(groupName)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set IModelDataProvider_GetWorks = result
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row   ' B — Артикул
    If lastRow < 4 Then
        Set IModelDataProvider_GetWorks = result
        Exit Function
    End If

    Dim i As Long
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            Dim we As WorkEntry
            Set we = New WorkEntry
            we.Code = CStr(ws.Cells(i, 2).Value)     ' B — Артикул
            we.Name = CStr(ws.Cells(i, 3).Value)     ' C — Наименование
            we.Unit = ""                             ' резерв
            we.NormHours = Val(ws.Cells(i, 4).Value) ' D — н/ч
            we.Price = Val(ws.Cells(i, 6).Value)     ' F — цена н/ч
            we.Note = ""
            result.Add we
        End If
    Next i

    Set IModelDataProvider_GetWorks = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetWorks: " & Err.Description)
    Set IModelDataProvider_GetWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetParts — Collection массивов [code,name,unit,price,note] из листа z4
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetParts(ByVal groupName As String, _
                                             ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetParts = result
        Exit Function
    End If

    On Error Resume Next
    Set ws = wb.Sheets("z4")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set IModelDataProvider_GetParts = result
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 4 Then
        Set IModelDataProvider_GetParts = result
        Exit Function
    End If

    Dim i As Long
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            Dim arr As Variant
            arr = Array(CStr(ws.Cells(i, 2).Value), CStr(ws.Cells(i, 3).Value), _
                        CStr(ws.Cells(i, 4).Value), Val(ws.Cells(i, 6).Value), "")
            result.Add arr
        End If
    Next i

    Set IModelDataProvider_GetParts = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetParts: " & Err.Description)
    Set IModelDataProvider_GetParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelWorks — Collection массивов из листа {groupName}w (тождества работ)
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetModelWorks(ByVal groupName As String, _
                                                  ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim wb As Workbook
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetModelWorks = result
        Exit Function
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(groupName & "w")
    On Error GoTo ErrHandler

    Dim col As Collection
    Set col = ReadWorkIdentitiesFromSheet(ws)

    Dim item As Variant
    For Each item In col
        Dim wi As WorkIdentity
        Set wi = item
        Dim arr As Variant
        arr = Array(wi.OutArticle, wi.OutName, wi.NormHours, wi.QtyZN, _
                    wi.Aggregate, wi.InName)
        result.Add arr
    Next item

    Set IModelDataProvider_GetModelWorks = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetModelWorks: " & Err.Description)
    Set IModelDataProvider_GetModelWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelParts — Collection массивов из листа {groupName}z4 (тождества запчастей)
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetModelParts(ByVal groupName As String, _
                                                  ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim wb As Workbook
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetModelParts = result
        Exit Function
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(groupName & "z4")
    On Error GoTo ErrHandler

    Dim col As Collection
    Set col = ReadPartIdentitiesFromSheet(ws)

    Dim item As Variant
    For Each item In col
        Dim pi As PartIdentity
        Set pi = item
        Dim arr As Variant
        arr = Array(pi.OutArticle, pi.OutName, pi.QtyZN, pi.Price, _
                    pi.Aggregate, pi.InCatNum, pi.InName)
        result.Add arr
    Next item

    Set IModelDataProvider_GetModelParts = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetModelParts: " & Err.Description)
    Set IModelDataProvider_GetModelParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetMatLibEntries — Collection массивов [target_type,target_code,target_name,
'                                         coefficient] по входящему коду.
' Строится из тождеств работ/запчастей листов {G}w/{G}z4.
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetMatLibEntries(ByVal groupName As String, _
                                                     ByVal entryCode As String) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim wb As Workbook
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetMatLibEntries = result
        Exit Function
    End If

    On Error Resume Next
    Dim wsWork As Worksheet
    Set wsWork = wb.Sheets(groupName & "w")
    Dim wsPart As Worksheet
    Set wsPart = wb.Sheets(groupName & "z4")
    On Error GoTo ErrHandler

    ' Тождества работ (по InName)
    Dim wcol As Collection
    Set wcol = ReadWorkIdentitiesFromSheet(wsWork)
    Dim item As Variant
    For Each item In wcol
        Dim wi As WorkIdentity
        Set wi = item
        If UCase$(Trim$(wi.InName)) = UCase$(Trim$(entryCode)) Then
            result.Add Array("mod_work", wi.OutArticle, wi.OutName, wi.QtyZN)
        End If
    Next item

    ' Тождества запчастей (по InCatNum)
    Dim pcol As Collection
    Set pcol = ReadPartIdentitiesFromSheet(wsPart)
    For Each item In pcol
        Dim pi As PartIdentity
        Set pi = item
        If UCase$(Trim$(pi.InCatNum)) = UCase$(Trim$(entryCode)) Then
            result.Add Array("mod_part", pi.OutArticle, pi.OutName, pi.QtyZN)
        End If
    Next item

    Set IModelDataProvider_GetMatLibEntries = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetMatLibEntries: " & Err.Description)
    Set IModelDataProvider_GetMatLibEntries = New Collection
End Function

' --------------------------------------------------------------------------
' GetWorkIdentities — Collection объектов WorkIdentity из листа {groupName}w
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetWorkIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetWorkIdentities = New Collection
        Exit Function
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(groupName & "w")
    On Error GoTo ErrHandler

    Set IModelDataProvider_GetWorkIdentities = ReadWorkIdentitiesFromSheet(ws)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetWorkIdentities: " & Err.Description)
    Set IModelDataProvider_GetWorkIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetPartIdentities — Collection объектов PartIdentity из листа {groupName}z4
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetPartIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Set wb = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set IModelDataProvider_GetPartIdentities = New Collection
        Exit Function
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(groupName & "z4")
    On Error GoTo ErrHandler

    Set IModelDataProvider_GetPartIdentities = ReadPartIdentitiesFromSheet(ws)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetPartIdentities: " & Err.Description)
    Set IModelDataProvider_GetPartIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetAllModelGroups — сканирует каталог base/models/*.xlsm
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetAllModelGroups() As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim basePath As String
    basePath = Mod_ModelDB.GetModelDBBasePath()

    Dim f As String
    f = Dir(basePath & "*.xlsm")
    Do While Len(f) > 0
        Dim gName As String
        gName = Mid$(f, 1, Len(f) - 5)
        result.Add gName
        f = Dir
    Loop

    Set IModelDataProvider_GetAllModelGroups = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "GetAllModelGroups: " & Err.Description)
    Set IModelDataProvider_GetAllModelGroups = New Collection
End Function

' --------------------------------------------------------------------------
' CreateModelGroupFile — создание модели (резерв). В Excel-режиме это заглушка.
' --------------------------------------------------------------------------
Private Function IModelDataProvider_CreateModelGroupFile(ByVal groupName As String) As Boolean
    On Error GoTo ErrHandler
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", _
        "CreateModelGroupFile: заглушка для группы '" & groupName & "'")
    IModelDataProvider_CreateModelGroupFile = True
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDBProvider", "CreateModelGroupFile: " & Err.Description)
    IModelDataProvider_CreateModelGroupFile = False
End Function

' --------------------------------------------------------------------------
' ModelGroupFileExists — проверяет наличие легаси-файла группы
' --------------------------------------------------------------------------
Private Function IModelDataProvider_ModelGroupFileExists(ByVal groupName As String) As Boolean
    IModelDataProvider_ModelGroupFileExists = Mod_ModelDB.ModelGroupFileExists(groupName)
End Function