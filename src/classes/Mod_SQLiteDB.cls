VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mod_SQLiteDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IModelDataProvider

' ============================================================
' Класс: Mod_SQLiteDB
' Назначение: SQLite-провайдер данных моделей через ADO/ODBC.
' Реализует IModelDataProvider. Читает и пишет SysW.db.
'
' Требование: установленный ODBC-драйвер "SQLite3 ODBC Driver"
' (sqliteodbc.dll). Используется ПОЗДНЕЕ связывание (CreateObject),
' поэтому ссылка Microsoft ActiveX Data Objects не обязательна —
' код компилируется без дополнительных References.
'
' При открытии задаются PRAGMA journal_mode=WAL и busy_timeout=5000
' (конкурентный доступ Python-писатель + VBA-читатели).
' ============================================================

' ---- Имя ODBC-драйвера SQLite (можно переопределить) ----
Private Const SQLITE_DRIVER_NAME As String = "SQLite3 ODBC Driver"

' ---- Константы ADO (позднее связывание) ----
Private Const adOpenForwardOnly As Long = 0
Private Const adLockReadOnly As Long = 1
Private Const adCmdText As Long = 1
Private Const adParamInput As Long = 1
Private Const adVarChar As Long = 200
Private Const adDouble As Long = 5
Private Const adInteger As Long = 3
Private Const adBoolean As Long = 11

' Таймаут ожидания при блокировке БД (мс)
Private Const BUSY_TIMEOUT_MS As Long = 5000

' Путь к БД (по умолчанию ThisWorkbook.Path & "\SysW.db")
Private m_dbPath As String

' ADO-подключение (позднее связывание)
Private m_conn As Object

' ============================================================
' Свойства
' ============================================================
Public Property Get DbPath() As String
    DbPath = m_dbPath
End Property

Public Property Let DbPath(ByVal v As String)
    m_dbPath = v
End Property

' --------------------------------------------------------------------------
' Class_Terminate — гарантированно закрывает соединение при уничтожении объекта
' --------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call CloseConnection
End Sub

' ============================================================
' Управление подключением
' ============================================================

' --------------------------------------------------------------------------
' IsConnected — возвращает True, если соединение открыто
' --------------------------------------------------------------------------
Public Function IsConnected() As Boolean
    On Error GoTo ErrHandler
    If Not m_conn Is Nothing Then
        If m_conn.State = 1 Then  ' adStateOpen = 1
            IsConnected = True
        End If
    End If
    Exit Function
ErrHandler:
    IsConnected = False
End Function

' --------------------------------------------------------------------------
' OpenConnection — открывает ADO-подключение к SQLite через ODBC.
' Задаёт WAL и busy_timeout. При недоступности драйвера — Err.Raise.
' --------------------------------------------------------------------------
Public Sub OpenConnection()
    On Error GoTo ErrHandler

    Dim db As String
    If Len(m_dbPath) = 0 Then
        m_dbPath = ThisWorkbook.Path & "\SysW.db"
    End If
    db = m_dbPath

    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: db=" & db)
    If IsConnected() Then
        Exit Sub
    End If

    ' Позднее связывание ADODB.Connection
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: создаю ADODB.Connection")
    Set m_conn = CreateObject("ADODB.Connection")
    m_conn.CommandTimeout = 30
    m_conn.ConnectionString = _
        "Driver={" & SQLITE_DRIVER_NAME & "};Database=" & db & ";ReadOnly=0;"
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: вызываю Open")
    m_conn.Open
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: Open OK, state=" & CStr(m_conn.State))

    ' WAL и busy_timeout для конкурентного доступа
    Call ExecuteNonQuery("PRAGMA journal_mode=WAL", Empty)
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: PRAGMA WAL OK")
    Call ExecuteNonQuery("PRAGMA busy_timeout=" & CStr(BUSY_TIMEOUT_MS), Empty)
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: PRAGMA busy_timeout OK")

    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: OK " & db)
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "OpenConnection: Ошибка num=" & CStr(Err.Number) & _
        " desc=" & Err.Description)
    Set m_conn = Nothing
    Err.Raise Err.Number, "Mod_SQLiteDB", "OpenConnection: " & Err.Description
End Sub

' --------------------------------------------------------------------------
' CloseConnection — закрывает соединение
' --------------------------------------------------------------------------
Public Sub CloseConnection()
    On Error Resume Next
    If Not m_conn Is Nothing Then
        If m_conn.State = 1 Then m_conn.Close  ' adStateOpen
        Set m_conn = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Низкоуровневые операции с БД
' ============================================================

' --------------------------------------------------------------------------
' AppendParam — добавляет параметр в команду Command по типу значения.
' --------------------------------------------------------------------------
Private Sub AppendParam(ByRef cmd As Object, ByVal value As Variant)
    Dim p As Object
    Select Case VarType(value)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            Set p = cmd.CreateParameter("p", adDouble, adParamInput, 8, CDbl(value))
        Case vbLong, vbInteger
            Set p = cmd.CreateParameter("p", adInteger, adParamInput, 4, CLng(value))
        Case vbBoolean
            Set p = cmd.CreateParameter("p", adBoolean, adParamInput, 2, value)
        Case vbNull, vbEmpty
            Set p = cmd.CreateParameter("p", adVarChar, adParamInput, 1, Null)
        Case Else
            Dim s As String
            s = CStr(value)
            Set p = cmd.CreateParameter("p", adVarChar, adParamInput, Len(s) + 1, s)
    End Select
    cmd.Parameters.Append p
End Sub

' --------------------------------------------------------------------------
' ExecuteScalar — выполняет SQL и возвращает скалярное значение.
' params — либо пустой Variant, либо массив значений для параметризации.
' --------------------------------------------------------------------------
Public Function ExecuteScalar(ByVal sql As String, ByRef params As Variant) As Variant
    On Error GoTo ErrHandler
    Dim rs As Object
    Set rs = ExecuteQuery(sql, params)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            ExecuteScalar = rs.Fields(0).Value
        End If
        rs.Close
    End If
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "ExecuteScalar: " & Err.Description)
    Err.Raise Err.Number, "Mod_SQLiteDB", "ExecuteScalar: " & Err.Description
End Function

' --------------------------------------------------------------------------
' ExecuteQuery — выполняет SELECT, возвращает Recordset (позднее связывание).
' --------------------------------------------------------------------------
Public Function ExecuteQuery(ByVal sql As String, ByRef params As Variant) As Object
    On Error GoTo ErrHandler
    If Not IsConnected() Then Call OpenConnection

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = m_conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    If IsArray(params) Then
        Dim i As Long
        For i = LBound(params) To UBound(params)
            Call AppendParam(cmd, params(i))
        Next i
    End If

    Set ExecuteQuery = cmd.Execute
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "ExecuteQuery: " & Err.Description)
    Err.Raise Err.Number, "Mod_SQLiteDB", "ExecuteQuery: " & Err.Description
End Function

' --------------------------------------------------------------------------
' ExecuteNonQuery — INSERT/UPDATE/DELETE/DDL/PRAGMA
' --------------------------------------------------------------------------
Public Sub ExecuteNonQuery(ByVal sql As String, ByRef params As Variant)
    On Error GoTo ErrHandler
    If Not IsConnected() Then Call OpenConnection

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = m_conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    If IsArray(params) Then
        Dim i As Long
        For i = LBound(params) To UBound(params)
            Call AppendParam(cmd, params(i))
        Next i
    End If

    Call cmd.Execute
    Exit Sub
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "ExecuteNonQuery: " & Err.Description)
    Err.Raise Err.Number, "Mod_SQLiteDB", "ExecuteNonQuery: " & Err.Description
End Sub

' ============================================================
' Реализация IModelDataProvider
' ============================================================

' --------------------------------------------------------------------------
' GetWorks — возвращает Collection объектов WorkEntry из таблицы works.
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetWorks(ByVal groupName As String, _
                                             ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT code, name, unit, norm_hours, price, note FROM works " & _
        "WHERE group_name = ?", Array(groupName))

    Do While Not rs.EOF
        Dim we As WorkEntry
        Set we = New WorkEntry
        we.Code = CStr(rs.Fields(0).Value)
        we.Name = CStr(rs.Fields(1).Value)
        we.Unit = CStr(rs.Fields(2).Value)
        we.NormHours = Val(rs.Fields(3).Value)
        we.Price = Val(rs.Fields(4).Value)
        we.Note = CStr(rs.Fields(5).Value)
        result.Add we
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetWorks = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetWorks: " & Err.Description)
    Set IModelDataProvider_GetWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetParts — возвращает Collection массивов [code,name,unit,price,note]
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetParts(ByVal groupName As String, _
                                             ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT code, name, unit, price, note FROM parts WHERE group_name = ?", _
        Array(groupName))

    Do While Not rs.EOF
        Dim arr As Variant
        arr = Array(CStr(rs.Fields(0).Value), CStr(rs.Fields(1).Value), _
                    CStr(rs.Fields(2).Value), Val(rs.Fields(3).Value), _
                    CStr(rs.Fields(4).Value))
        result.Add arr
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetParts = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetParts: " & Err.Description)
    Set IModelDataProvider_GetParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelWorks — Collection массивов [out_article,out_name,norm_hours,qty_zn,
'                                      aggregate,in_name]
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetModelWorks(ByVal groupName As String, _
                                                  ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT out_article, out_name, norm_hours, qty_zn, aggregate, in_name " & _
        "FROM model_works WHERE group_name = ?", Array(groupName))

    Do While Not rs.EOF
        Dim arr As Variant
        arr = Array(CStr(rs.Fields(0).Value), CStr(rs.Fields(1).Value), _
                    Val(rs.Fields(2).Value), Val(rs.Fields(3).Value), _
                    CStr(rs.Fields(4).Value), CStr(rs.Fields(5).Value))
        result.Add arr
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetModelWorks = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetModelWorks: " & Err.Description)
    Set IModelDataProvider_GetModelWorks = New Collection
End Function

' --------------------------------------------------------------------------
' GetModelParts — Collection массивов [out_article,out_name,qty_zn,price,
'                                      aggregate,in_catnum,in_name]
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetModelParts(ByVal groupName As String, _
                                                  ByRef filters As Variant) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT out_article, out_name, qty_zn, price, aggregate, in_catnum, in_name " & _
        "FROM model_parts WHERE group_name = ?", Array(groupName))

    Do While Not rs.EOF
        Dim arr As Variant
        arr = Array(CStr(rs.Fields(0).Value), CStr(rs.Fields(1).Value), _
                    Val(rs.Fields(2).Value), Val(rs.Fields(3).Value), _
                    CStr(rs.Fields(4).Value), CStr(rs.Fields(5).Value), _
                    CStr(rs.Fields(6).Value))
        result.Add arr
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetModelParts = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetModelParts: " & Err.Description)
    Set IModelDataProvider_GetModelParts = New Collection
End Function

' --------------------------------------------------------------------------
' GetMatLibEntries — Collection массивов [target_type,target_code,target_name,
'                                         coefficient] по входящему коду.
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetMatLibEntries(ByVal groupName As String, _
                                                     ByVal entryCode As String) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT target_type, target_code, target_name, coefficient " & _
        "FROM matlib_entries WHERE group_name = ? AND entry_code = ?", _
        Array(groupName, entryCode))

    Do While Not rs.EOF
        Dim arr As Variant
        arr = Array(CStr(rs.Fields(0).Value), CStr(rs.Fields(1).Value), _
                    CStr(rs.Fields(2).Value), Val(rs.Fields(3).Value))
        result.Add arr
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetMatLibEntries = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetMatLibEntries: " & Err.Description)
    Set IModelDataProvider_GetMatLibEntries = New Collection
End Function

' --------------------------------------------------------------------------
' GetWorkIdentities — Collection объектов WorkIdentity из model_works
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetWorkIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT out_article, out_name, norm_hours, qty_zn, aggregate, in_name " & _
        "FROM model_works WHERE group_name = ?", Array(groupName))

    Do While Not rs.EOF
        Dim wi As WorkIdentity
        Set wi = New WorkIdentity
        wi.OutArticle = CStr(rs.Fields(0).Value)
        wi.OutName = CStr(rs.Fields(1).Value)
        wi.NormHours = Val(rs.Fields(2).Value)
        wi.QtyZN = Val(rs.Fields(3).Value)
        wi.Aggregate = CStr(rs.Fields(4).Value)
        wi.InName = CStr(rs.Fields(5).Value)
        result.Add wi
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetWorkIdentities = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetWorkIdentities: " & Err.Description)
    Set IModelDataProvider_GetWorkIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetPartIdentities — Collection объектов PartIdentity из model_parts
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetPartIdentities(ByVal groupName As String) As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery( _
        "SELECT out_article, out_name, qty_zn, price, aggregate, in_catnum, in_name " & _
        "FROM model_parts WHERE group_name = ?", Array(groupName))

    Do While Not rs.EOF
        Dim pi As PartIdentity
        Set pi = New PartIdentity
        pi.OutArticle = CStr(rs.Fields(0).Value)
        pi.OutName = CStr(rs.Fields(1).Value)
        pi.QtyZN = Val(rs.Fields(2).Value)
        pi.Price = Val(rs.Fields(3).Value)
        pi.Aggregate = CStr(rs.Fields(4).Value)
        pi.InCatNum = CStr(rs.Fields(5).Value)
        pi.InName = CStr(rs.Fields(6).Value)
        result.Add pi
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetPartIdentities = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetPartIdentities: " & Err.Description)
    Set IModelDataProvider_GetPartIdentities = New Collection
End Function

' --------------------------------------------------------------------------
' GetAllModelGroups — Collection строк с именами групп из model_groups
' --------------------------------------------------------------------------
Private Function IModelDataProvider_GetAllModelGroups() As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection

    Dim rs As Object
    Set rs = ExecuteQuery("SELECT group_name FROM model_groups ORDER BY group_name", Empty)

    Do While Not rs.EOF
        result.Add CStr(rs.Fields(0).Value)
        rs.MoveNext
    Loop
    rs.Close

    Set IModelDataProvider_GetAllModelGroups = result
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "GetAllModelGroups: " & Err.Description)
    Set IModelDataProvider_GetAllModelGroups = New Collection
End Function

' --------------------------------------------------------------------------
' CreateModelGroupFile — создаёт запись группы в model_groups (INSERT OR REPLACE)
' --------------------------------------------------------------------------
Private Function IModelDataProvider_CreateModelGroupFile(ByVal groupName As String) As Boolean
    On Error GoTo ErrHandler
    Call ExecuteNonQuery( _
        "INSERT OR REPLACE INTO model_groups(group_name) VALUES (?)", _
        Array(groupName))
    IModelDataProvider_CreateModelGroupFile = True
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "CreateModelGroupFile: " & Err.Description)
    IModelDataProvider_CreateModelGroupFile = False
End Function

' --------------------------------------------------------------------------
' ModelGroupFileExists — проверяет наличие группы в model_groups
' --------------------------------------------------------------------------
Private Function IModelDataProvider_ModelGroupFileExists(ByVal groupName As String) As Boolean
    On Error GoTo ErrHandler
    Dim v As Variant
    v = ExecuteScalar( _
        "SELECT COUNT(*) FROM model_groups WHERE group_name = ?", Array(groupName))
    IModelDataProvider_ModelGroupFileExists = (CLng(v) > 0)
    Exit Function
ErrHandler:
    Call Mod_Logger.WriteLog("Mod_SQLiteDB", "ModelGroupFileExists: " & Err.Description)
    IModelDataProvider_ModelGroupFileExists = False
End Function