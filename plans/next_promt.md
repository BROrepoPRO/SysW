соблюдаем: .ycarules


Architect - планирование, декомпозиция задач
Code - исполнение
Ask - анализ, объяснение
Debug - отладка, ошибки
Orchestrator - межролевая координация


исправить ошибку:
при запуске импорта ошибка здесь:

compille error:Only public user defined types defined in public object modules can be used as parameters or return types for public procedures of class modules or as fields of public user defined types

Эта ошибка имеет следующие причину и решение:

Предпринята попытка использовать общедоступный пользовательский тип в качестве параметра или возвращаемого типа для общедоступной процедуры модуля класса либо в качестве поля общедоступного пользовательского типа. Таким способом могут использоваться только общедоступные пользовательские типы, которые определены в общедоступном объектном модуле.

Public Function GetWorkIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim identity As Variant
    Dim tmpIdentity As WorkIdentity
    Dim SheetName As String

    SheetName = groupName & "w"
    Set result = New Collection

    ' Открываем файл группы
    Set wb = OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set GetWorkIdentities = result
        Exit Function
    End If

    ' Получаем лист {GroupName}w
    On Error Resume Next
    Set ws = wb.Sheets(SheetName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorkIdentities: Лист " & SheetName & " не найден")
        Set GetWorkIdentities = result
        Exit Function
    End If

    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow < 4 Then
        Set GetWorkIdentities = result
        Exit Function
    End If

    ' Читаем данные
    ' Колонки UAZw: A=№п/п, B=Артикул, C=Наименование, D=н/ч,
    '   E=кол-во оп, F=цена н/ч, G=Кол-во ЗН, H=Сумма ЗН,
    '   I=Агрегат, J=Наим-ние, K=Кол. оп., L=Цена, M=Всего
    For i = 4 To lastRow
        ' Пропускаем пустые строки и строки без агрегата (I)
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then ' I — Агрегат
                tmpIdentity.OutArticle = CStr(ws.Cells(i, 2).Value)  ' B
                tmpIdentity.OutName = CStr(ws.Cells(i, 3).Value)     ' C
                tmpIdentity.NormHours = Val(ws.Cells(i, 4).Value)    ' D
                tmpIdentity.QtyZN = Val(ws.Cells(i, 7).Value)        ' G
                tmpIdentity.Aggregate = CStr(ws.Cells(i, 9).Value)   ' I
                tmpIdentity.InName = CStr(ws.Cells(i, 10).Value)     ' J
                identity