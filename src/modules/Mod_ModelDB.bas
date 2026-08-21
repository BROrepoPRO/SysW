Attribute VB_Name = "Mod_ModelDB"
Option Explicit

' ============================================================
' Модуль: Mod_ModelDB
' Назначение: Доступ к данным модельных групп (базовый слой абстракции).
'
' С v1.0.5 модуль является ФАБРИКОЙ и обёрткой над провайдерами,
' реализующими IModelDataProvider:
'   - Mod_SQLiteDB       — основной провайдер (SQLite через ADO/ODBC),
'                           активен при наличии SysW.db и ODBC-драйвера.
'   - Mod_ModelDBProvider — резервный (Excel) провайдер (чтение легаси-файлов).
'
' Фабрика GetModelDataProvider() выбирает активный источник данных.
' Старые GetWorks/GetWorkIdentities/GetPartIdentities сохранены как
' ДЕЛЕГАТЫ к активному провайдеру (обратная совместимость вызовов
' из Mod_AutoMatch и тестов).
' ============================================================

' ============================================================
' Фабрика провайдера данных моделей
' ============================================================

' --------------------------------------------------------------------------
' GetModelDataProvider
' Заполняет провайдер через ByRef-параметр (а не возвратом интерфейса из функции),
' чтобы исключить COM-маршалинг интерфейсного объекта в невидимом Excel при
' внешнем запуске (COM-ошибка 0x80020009 / 0x800AC472 на RunModelDBTests).
' Приоритет: SQLite (если константа MODELDB_PROVIDER_SQLITE = True,
' SysW.db существует рядом с work.xlsm и ODBC-драйвер доступен).
' Иначе — резервный Excel-провайдер.
' --------------------------------------------------------------------------
Public Sub GetModelDataProvider(ByRef provider As IModelDataProvider)
    On Error GoTo ErrDiag

    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: START")

    ' Если SQLite отключён константой — сразу Excel.
    ' Чтение флага через функцию Mod_Constants.SqliteProviderEnabled():
    ' прямое обращение к Public Const из другого модуля в этой книге
    ' даёт ошибку компиляции (461 / Variable not defined), поэтому
    ' значение константы возвращается через функцию модуля.
    If Not Mod_Constants.SqliteProviderEnabled() Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: SQLite выкл константой -> Excel")
        GoTo ExcelFallback
    End If

    ' Проверяем наличие SysW.db рядом с work.xlsm
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\SysW.db"
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: dbPath=" & dbPath)
    If Len(Dir(dbPath)) = 0 Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", _
            "GetModelDataProvider: SysW.db не найден (" & dbPath & ") — Excel fallback")
        GoTo ExcelFallback
    End If

    ' Пробуем создать SQLite-провайдер и открыть соединение
    Dim sqlite As Mod_SQLiteDB
    Set sqlite = New Mod_SQLiteDB
    sqlite.DbPath = dbPath
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: создан Mod_SQLiteDB, вызов OpenConnection")

    On Error Resume Next
    Call sqlite.OpenConnection
    If Err.Number <> 0 Then
        Call Mod_Logger.WriteLog("Mod_ModelDB", _
            "GetModelDataProvider: OpenConnection error — Excel fallback: " & _
            Err.Description & " | num=" & CStr(Err.Number))
        Err.Clear
        Call sqlite.CloseConnection
        Set sqlite = Nothing
        On Error GoTo ExcelFallback
        GoTo ExcelFallback
    End If
    On Error GoTo ExcelFallback

    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: выбран SQLite-провайдер")
    Set provider = sqlite
    Exit Sub

ErrDiag:
    Call Mod_Logger.WriteLog("Mod_ModelDB", _
        "GetModelDataProvider: ERRDIAG num=" & CStr(Err.Number) & " desc=" & Err.Description)
    Err.Clear
    GoTo ExcelFallback

ExcelFallback:
    On Error Resume Next
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelDataProvider: выбран Excel-провайдер")
    Set provider = New Mod_ModelDBProvider
End Sub

' ============================================================
' Функции путей к модельным файлам
' ============================================================

' --------------------------------------------------------------------------
' GetModelDBBasePath
' Возвращает путь к каталогу base\models\ относительно расположения work.xlsm.
' --------------------------------------------------------------------------
Public Function GetModelDBBasePath() As String
    GetModelDBBasePath = ThisWorkbook.Path & "\base\models\"
End Function

' --------------------------------------------------------------------------
' GetModelGroupFilePath
' Возвращает полный путь к файлу группы (сначала .xlsm, затем .xlsx).
' --------------------------------------------------------------------------
Public Function GetModelGroupFilePath(ByVal groupName As String) As String
    Dim basePath As String
    basePath = GetModelDBBasePath()

    Dim xlsmPath As String
    Dim xlsxPath As String
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
' Проверяет существование легаси-файла группы.
' --------------------------------------------------------------------------
Public Function ModelGroupFileExists(ByVal groupName As String) As Boolean
    Dim filePath As String
    filePath = GetModelGroupFilePath(groupName)
    ModelGroupFileExists = (Len(Dir(filePath)) > 0)
End Function

' ============================================================
' Открытие файла группы (используется ручным подбором и Excel-провайдером)
' ============================================================

' --------------------------------------------------------------------------
' OpenModelGroupFile
' Открывает файл группы {groupName}.xlsm/.xlsx (если ещё не открыт)
' и возвращает ссылку на Workbook. Если файл не найден — возвращает Nothing.
' --------------------------------------------------------------------------
Public Function OpenModelGroupFile(ByVal groupName As String) As Workbook
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim filePath As String
    Dim wbName As String

    filePath = GetModelGroupFilePath(groupName)
    If Len(filePath) = 0 Then
        Set OpenModelGroupFile = Nothing
        Exit Function
    End If

    wbName = Mid$(filePath, InStrRev(filePath, "\") + 1)

    ' Проверяем, не открыт ли уже файл
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo ErrHandler

    If Not wb Is Nothing Then
        Set OpenModelGroupFile = wb
        Exit Function
    End If

    If Not Mod_Utils.FileExists(filePath) Then
        Set OpenModelGroupFile = Nothing
        Exit Function
    End If

    Set wb = Workbooks.Open(filePath, ReadOnly:=False)
    Set OpenModelGroupFile = wb
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "OpenModelGroupFile: Ошибка — " & Err.Description)
    Set OpenModelGroupFile = Nothing
End Function

' ============================================================
' Делегаты к активному провайдеру (обратная совместимость API)
' ============================================================

' --------------------------------------------------------------------------
' GetWorks
' Возвращает Collection записей работ группы (объекты WorkEntry).
' --------------------------------------------------------------------------
Public Function GetWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetWorks = provider.GetWorks(groupName, filters)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorks: Ошибка — " & Err.Description)
    Set GetWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetParts
' Возвращает Collection запчастей группы.
' --------------------------------------------------------------------------
Public Function GetParts(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetParts = provider.GetParts(groupName, filters)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetParts: Ошибка — " & Err.Description)
    Set GetParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelWorks
' Возвращает Collection модельных работ группы (тождества работ).
' --------------------------------------------------------------------------
Public Function GetModelWorks(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetModelWorks = provider.GetModelWorks(groupName, filters)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelWorks: Ошибка — " & Err.Description)
    Set GetModelWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelParts
' Возвращает Collection модельных запчастей группы (тождества запчастей).
' --------------------------------------------------------------------------
Public Function GetModelParts(ByVal groupName As String, ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetModelParts = provider.GetModelParts(groupName, filters)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetModelParts: Ошибка — " & Err.Description)
    Set GetModelParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetMatLibEntries
' Возвращает Collection библиотеки соответствий по входящему коду.
' --------------------------------------------------------------------------
Public Function GetMatLibEntries(ByVal groupName As String, _
                                 ByVal entryCode As String) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetMatLibEntries = provider.GetMatLibEntries(groupName, entryCode)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetMatLibEntries: Ошибка — " & Err.Description)
    Set GetMatLibEntries = New Collection
End Function

' --------------------------------------------------------------------------
' GetWorkIdentities
' Возвращает Collection тождеств работ группы (объекты WorkIdentity).
' --------------------------------------------------------------------------
Public Function GetWorkIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetWorkIdentities = provider.GetWorkIdentities(groupName)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetWorkIdentities: Ошибка — " & Err.Description)
    Set GetWorkIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetPartIdentities
' Возвращает Collection тождеств запчастей группы (объекты PartIdentity).
' --------------------------------------------------------------------------
Public Function GetPartIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetPartIdentities = provider.GetPartIdentities(groupName)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetPartIdentities: Ошибка — " & Err.Description)
    Set GetPartIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetAllModelGroups
' Возвращает Collection имён всех групп моделей.
' --------------------------------------------------------------------------
Public Function GetAllModelGroups() As Collection
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    Set GetAllModelGroups = provider.GetAllModelGroups()
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "GetAllModelGroups: Ошибка — " & Err.Description)
    Set GetAllModelGroups = New Collection
End Function

' --------------------------------------------------------------------------
' CreateModelGroupFile
' Создаёт группу моделей (в SQLite — запись в model_groups).
' Возвращает True при успехе.
' --------------------------------------------------------------------------
Public Function CreateModelGroupFile(ByVal groupName As String) As Boolean
    On Error GoTo ErrHandler
    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)
    CreateModelGroupFile = provider.CreateModelGroupFile(groupName)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "CreateModelGroupFile: Ошибка — " & Err.Description)
    CreateModelGroupFile = False
End Function

' --------------------------------------------------------------------------
' FindModelGroupByModel
' Ищет имя группы моделей по имени модели (названию ТС).
' Перебирает все группы из провайдера и возвращает ту, чьё имя совпадает
' с modelName (с учётом регистра) либо содержится в нём.
' Если группа не найдена — возвращает пустую строку.
' Интеграция Фазы B: позволяет читать группу из SysW.db (через провайдер).
' --------------------------------------------------------------------------
Public Function FindModelGroupByModel(ByVal modelName As String) As String
    On Error GoTo ErrHandler

    Dim provider As IModelDataProvider
    Call GetModelDataProvider(provider)

    Dim groups As Collection
    Set groups = provider.GetAllModelGroups()

    Dim g As Variant
    Dim key As String
    key = UCase$(Trim$(modelName))

    If key <> "" Then
        ' Точное совпадение имени группы с именем модели
        For Each g In groups
            Dim gName As String
            gName = UCase$(Trim$(CStr(g)))
            If gName <> "" And gName = key Then
                FindModelGroupByModel = CStr(g)
                Exit Function
            End If
        Next g
        ' Совпадение по вхождению имени группы в название модели
        For Each g In groups
            gName = UCase$(Trim$(CStr(g)))
            If gName <> "" And InStr(1, key, gName, vbTextCompare) > 0 Then
                FindModelGroupByModel = CStr(g)
                Exit Function
            End If
        Next g
    End If

    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_ModelDB", "FindModelGroupByModel: Ошибка — " & Err.Description)
    FindModelGroupByModel = ""
End Function