Attribute VB_Name = "Mod_ModelDB"
Option Explicit

' ============================================================
' Модуль: Mod_ModelDB
' Назначение: Доступ к файлам модельных групп
'             (базовый слой абстракции для работы с base/models/)
' ============================================================

' ============================================================
' Функция получения базового пути к моделям
' ============================================================

' --------------------------------------------------------------------------
' GetModelDBBasePath
' Возвращает путь к каталогу base\models\ относительно расположения work.xlsm.
' Это заменяет жёстко заданный абсолютный путь MODELDB_BASE_PATH.
' Пример: если work.xlsm в L:\PROject\SysW\, то вернёт L:\PROject\SysW\base\models\
' --------------------------------------------------------------------------
Public Function GetModelDBBasePath() As String
    GetModelDBBasePath = ThisWorkbook.Path & "\base\models\"
End Function

' ============================================================
' Вспомогательные функции
' ============================================================

' --------------------------------------------------------------------------
' GetModelGroupFilePath
' Возвращает полный путь к файлу группы.
' Сначала проверяет .xlsm, затем .xlsx.
' --------------------------------------------------------------------------
Public Function GetModelGroupFilePath(ByVal groupName As String) As String
    Dim xlsmPath As String
    Dim xlsxPath As String

    Dim basePath As String
    basePath = GetModelDBBasePath()

    xlsmPath = basePath & groupName & ".xlsm"
    xlsxPath = basePath & groupName & ".xlsx"

    If Len(Dir(xlsmPath)) > 0 Then
        GetModelGroupFilePath = xlsmPath
    Else
        GetModelGroupFilePath = xlsxPath
    End If
End Function

' --------------------------------------------------------------------------
' ModelGroupFileExists
' Проверяет существование файла группы
' --------------------------------------------------------------------------
Public Function ModelGroupFileExists(ByVal groupName As String) As Boolean
    Dim filePath As String
    filePath = GetModelGroupFilePath(groupName)
    ModelGroupFileExists = (Len(Dir(filePath)) > 0)
End Function

' ============================================================
' Основные функции
' ============================================================

' --------------------------------------------------------------------------
' OpenModelGroupFile
' Открывает файл группы {groupName}.xlsm/.xlsx (если ещё не открыт)
' и возвращает ссылку на Workbook.
' Если файл не найден — возвращает Nothing.
' --------------------------------------------------------------------------
Public Function OpenModelGroupFile(ByVal groupName As String) As Workbook
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim filePath As String
    Dim wbName As String

    ' 1. Определяем путь к файлу (сначала .xlsm, потом .xlsx)
    filePath = GetModelGroupFilePath(groupName)
    If Len(filePath) = 0 Then
        Set OpenModelGroupFile = Nothing
        Exit Function
    End If

    ' Извлекаем имя файла из пути
    wbName = Mid$(filePath, InStrRev(filePath, "\") + 1)

    ' 2. Проверяем, не открыт ли уже файл
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo ErrHandler

    If Not wb Is Nothing Then
        ' Файл уже открыт — возвращаем ссылку
        Set OpenModelGroupFile = wb
        Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Файл " & wbName & " уже открыт")
        Exit Function
    End If

    ' 3. Проверяем существование файла
    If Not Mod_Utils.FileExists(filePath) Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Файл не найден — " & filePath)
        Set OpenModelGroupFile = Nothing
        Exit Function
    End If

    ' 4. Открываем файл
    Set wb = Workbooks.Open(filePath, ReadOnly:=False)
    Set OpenModelGroupFile = wb

    Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Открыт файл " & filePath)
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Ошибка — " & Err.Description)
    Set OpenModelGroupFile = Nothing
End Function

' --------------------------------------------------------------------------
' GetWorks
' Возвращает коллекцию работ из листа {groupName} файла группы
' с применением фильтров.
' На данном этапе — заглушка для будущего использования.
' --------------------------------------------------------------------------
Public Function GetWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim entry As Variant
    Dim tmpEntry As ModelTypes.WorkEntry

    Set result = New Collection

    ' Открываем файл группы
    Set wb = OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set GetWorks = result
        Exit Function
    End If

    ' Получаем лист {groupName}
    On Error Resume Next
    Set ws = wb.Sheets(groupName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorks: Лист " & groupName & " не найден в файле " & groupName & ".xlsx")
        Set GetWorks = result
        Exit Function
    End If

    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 4 Then
        Set GetWorks = result
        Exit Function
    End If

    ' Читаем данные (столбцы: A=Code, B=Name, C=Unit, D=NormHours, E=Price, F=Note)
    For i = 4 To lastRow
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            tmpEntry.Code = CStr(ws.Cells(i, 1).Value)
            tmpEntry.Name = CStr(ws.Cells(i, 2).Value)
            tmpEntry.Unit = CStr(ws.Cells(i, 3).Value)
            tmpEntry.NormHours = Val(ws.Cells(i, 4).Value)
            tmpEntry.Price = Val(ws.Cells(i, 5).Value)
            tmpEntry.Note = CStr(ws.Cells(i, 6).Value)
            entry = tmpEntry
            result.Add entry
        End If
    Next i

    Set GetWorks = result
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorks: Ошибка — " & Err.Description)
    Set GetWorks = New Collection
End Function

' ============================================================
' Функции чтения тождеств
' ============================================================

' --------------------------------------------------------------------------
' GetWorkIdentities
' Читает тождества работ из листа {GroupName}w файла группы.
' Возвращает коллекцию WorkIdentity.
' Данные начинаются с 4-й строки.
' Пустая строка или отсутствие агрегата (столбец I) = разделитель/пропуск.
' --------------------------------------------------------------------------
Public Function GetWorkIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim identity As Variant
    Dim tmpIdentity As ModelTypes.WorkIdentity
    Dim sheetName As String

    sheetName = groupName & "w"
    Set result = New Collection

    ' Открываем файл группы
    Set wb = OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set GetWorkIdentities = result
        Exit Function
    End If

    ' Получаем лист {GroupName}w
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorkIdentities: Лист " & sheetName & " не найден")
        Set GetWorkIdentities = result
        Exit Function
    End If

    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
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
                identity = tmpIdentity
                result.Add identity
            End If
        End If
    Next i

    Set GetWorkIdentities = result
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorkIdentities: Ошибка — " & Err.Description)
    Set GetWorkIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetPartIdentities
' Читает тождества запчастей из листа {GroupName}z4 файла группы.
' Возвращает коллекцию PartIdentity.
' Данные начинаются с 4-й строки.
' Пустая строка или отсутствие агрегата (столбец I) = разделитель/пропуск.
' --------------------------------------------------------------------------
Public Function GetPartIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim identity As Variant
    Dim tmpIdentity As ModelTypes.PartIdentity
    Dim sheetName As String

    sheetName = groupName & "z4"
    Set result = New Collection

    ' Открываем файл группы
    Set wb = OpenModelGroupFile(groupName)
    If wb Is Nothing Then
        Set GetPartIdentities = result
        Exit Function
    End If

    ' Получаем лист {GroupName}z4
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetPartIdentities: Лист " & sheetName & " не найден")
        Set GetPartIdentities = result
        Exit Function
    End If

    ' Определяем последнюю строку (данные с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 4 Then
        Set GetPartIdentities = result
        Exit Function
    End If

    ' Читаем данные
    ' Колонки UAZz4: A=№ п/п, B=Артикул, C=Наименование, D=Ед. изм.,
    '   E=кол-во, F=Цена за ед. изм., G=Кол-во ЗН, H=Сумма ЗН,
    '   I=АГРЕГАТ, J=№ кат., K=Наим-ние, L=Колво, M=Цена, N=Всего
    For i = 4 To lastRow
        ' Пропускаем пустые строки и строки без агрегата (I)
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            If Not IsEmpty(ws.Cells(i, 9).Value) Then ' I — АГРЕГАТ
                tmpIdentity.OutArticle = CStr(ws.Cells(i, 2).Value)  ' B
                tmpIdentity.OutName = CStr(ws.Cells(i, 3).Value)     ' C
                tmpIdentity.QtyZN = Val(ws.Cells(i, 7).Value)        ' G
                tmpIdentity.Price = Val(ws.Cells(i, 6).Value)        ' F
                tmpIdentity.Aggregate = CStr(ws.Cells(i, 9).Value)   ' I
                tmpIdentity.InCatNum = CStr(ws.Cells(i, 10).Value)   ' J
                tmpIdentity.InName = CStr(ws.Cells(i, 11).Value)     ' K
                identity = tmpIdentity
                result.Add identity
            End If
        End If
    Next i

    Set GetPartIdentities = result
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetPartIdentities: Ошибка — " & Err.Description)
    Set GetPartIdentities = New Collection
End Function