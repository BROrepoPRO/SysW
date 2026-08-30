Attribute VB_Name = "Mod_Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Private Module

' ============================================================
' Модуль: Mod_Logger
' Назначение: Централизованное логирование «2 лога» с ротацией:
'   - Системный лог: logs/log.txt           (общие системные события)
'   - Тестовый лог:  logs/test_results.log  (детальный лог тестов,
'                     уровни INFO/WARN/ERROR, включая ошибки VBA)
' ============================================================

' Размер лога по умолчанию для ротации (в KB)
Private Const DEFAULT_MAX_LOG_SIZE_KB As Long = 100

' Допустимые уровни тестового лога
Private Const LVL_INFO As String = "INFO"
Private Const LVL_WARN As String = "WARN"
Private Const LVL_ERROR As String = "ERROR"

' --------------------------------------------------------------------------
' GetLogPath
' Возвращает путь к системному логу (logs/log.txt рядом с книгой)
' --------------------------------------------------------------------------
Public Function GetLogPath() As String
    GetLogPath = ThisWorkbook.path & "\" & Mod_Constants.LOGS_DIR & "\" & Mod_Constants.LOG_FILE
End Function

' --------------------------------------------------------------------------
' GetTestLogPath
' Возвращает путь к расширенному тестовому логу (logs/test_results.log)
' --------------------------------------------------------------------------
Public Function GetTestLogPath() As String
    GetTestLogPath = ThisWorkbook.path & "\" & Mod_Constants.LOGS_DIR & "\" & Mod_Constants.TEST_LOG_FILE
End Function

' --------------------------------------------------------------------------
' FormatTimestamp
' Форматирует текущее время в формат [YYYY-MM-DD HH:MM:SS]
' --------------------------------------------------------------------------
Private Function FormatTimestamp() As String
    FormatTimestamp = "[" & Format(VBA.DateTime.Now, "yyyy-mm-dd hh:nn:ss") & "]"
End Function

' --------------------------------------------------------------------------
' WriteLog
' Основной метод логирования СИСТЕМНОГО лога (logs/log.txt)
' Формат: [2026-07-15 14:30:00] [ModuleName] message
' --------------------------------------------------------------------------
Public Sub WriteLog(ByVal moduleName As String, ByVal message As String)
    On Error GoTo ErrHandler

    Dim LogPath As String
    Dim F As Long

    LogPath = GetLogPath()

    ' Проверка ротации перед записью
    Call RotateLogIfNeeded(DEFAULT_MAX_LOG_SIZE_KB)

    F = FreeFile
    Open LogPath For Append As #F
    Print #F, FormatTimestamp & " [" & moduleName & "] " & message
    Close #F

    Exit Sub

ErrHandler:
    ' При ошибке логирования — игнорируем, чтобы не вызывать каскад ошибок
End Sub

' --------------------------------------------------------------------------
' WriteLogE
' Логирование ошибок в СИСТЕМНЫЙ лог с префиксом [ERROR]
' --------------------------------------------------------------------------
Public Sub WriteLogE(ByVal moduleName As String, ByVal message As String)
    Call WriteLog(moduleName, "[ERROR] " & message)
End Sub

' --------------------------------------------------------------------------
' WriteTestLog
' Запись в РАСШИРЕННЫЙ ТЕСТОВЫЙ лог (logs/test_results.log) с уровнем.
' level — один из: INFO / WARN / ERROR (нечувствителен к регистру;
' некорректное значение приводится к INFO).
' Формат: [2026-07-15 14:30:00] [LEVEL] [ModuleName] message
' Используется Mod_FullTestRunner для детального лога тестов (включая
' ошибки VBA) и run_tests.py для согласованного чтения/дополнения.
' --------------------------------------------------------------------------
Public Sub WriteTestLog(ByVal moduleName As String, ByVal level As String, ByVal message As String)
    On Error GoTo ErrHandler

    Dim TestLogPath As String
    Dim F As Long
    Dim Lvl As String

    Lvl = UCase$(Trim$(level))
    If Lvl <> LVL_INFO And Lvl <> LVL_WARN And Lvl <> LVL_ERROR Then
        Lvl = LVL_INFO
    End If

    TestLogPath = GetTestLogPath()

    F = FreeFile
    Open TestLogPath For Append As #F
    Print #F, FormatTimestamp & " [" & Lvl & "] [" & moduleName & "] " & message
    Close #F

    Exit Sub

ErrHandler:
    ' При ошибке логирования — игнорируем, чтобы не вызывать каскад ошибок
End Sub

' --------------------------------------------------------------------------
' LogTestInfo / LogTestWarn / LogTestError
' Удобные обёртки записи в тестовый лог с фиксированным уровнем.
' --------------------------------------------------------------------------
Public Sub LogTestInfo(ByVal moduleName As String, ByVal message As String)
    Call WriteTestLog(moduleName, LVL_INFO, message)
End Sub

Public Sub LogTestWarn(ByVal moduleName As String, ByVal message As String)
    Call WriteTestLog(moduleName, LVL_WARN, message)
End Sub

Public Sub LogTestError(ByVal moduleName As String, ByVal message As String)
    Call WriteTestLog(moduleName, LVL_ERROR, message)
End Sub

' --------------------------------------------------------------------------
' RotateLogIfNeeded
' Ротация СИСТЕМНОГО лога если превышен указанный размер (в KB)
' Если log.txt > maxSizeKB, переименовывает в log_old.txt и создаёт новый
' --------------------------------------------------------------------------
Public Sub RotateLogIfNeeded(ByVal maxSizeKB As Long)
    On Error Resume Next

    Dim LogPath As String
    Dim OldLogPath As String
    Dim F As Long
    Dim fso As Object

    LogPath = GetLogPath()
    OldLogPath = ThisWorkbook.path & "\" & Mod_Constants.LOGS_DIR & "\log_old.txt"

    ' Проверяем существование файла
    If Len(Dir(LogPath)) = 0 Then
        Exit Sub
    End If

    ' Проверяем размер файла
    F = FreeFile
    Open LogPath For Input As #F
    If LOF(F) > maxSizeKB * 1024 Then
        Close #F

        ' Используем FileSystemObject для переименования
        Set fso = CreateObject("Scripting.FileSystemObject")

        ' Удаляем старый log_old.txt если существует
        If fso.FileExists(OldLogPath) Then
            fso.DeleteFile OldLogPath, True
        End If

        ' Переименовываем текущий лог в log_old.txt
        fso.MoveFile LogPath, OldLogPath

        Set fso = Nothing
    Else
        Close #F
    End If

    ' On Error GoTo 0 не вызываем — ошибки ротации игнорируем
End Sub

' --------------------------------------------------------------------------
' ClearLog
' Очистка СИСТЕМНОГО лога (удаление и создание пустого)
' --------------------------------------------------------------------------
Public Sub ClearLog()
    On Error Resume Next

    Dim LogPath As String
    Dim F As Long

    LogPath = GetLogPath()

    ' Удаляем файл если существует
    If Len(Dir(LogPath)) > 0 Then
        Kill LogPath
    End If

    ' Создаём новый пустой файл
    F = FreeFile
    Open LogPath For Output As #F
    Close #F

    ' On Error GoTo 0 не вызываем — ошибки очистки игнорируем
End Sub

' --------------------------------------------------------------------------
' ClearTestLog
' Очистка РАСШИРЕННОГО ТЕСТОВОГО лога (удаление и создание пустого)
' --------------------------------------------------------------------------
Public Sub ClearTestLog()
    On Error Resume Next

    Dim TestLogPath As String
    Dim F As Long

    TestLogPath = GetTestLogPath()

    ' Удаляем файл если существует
    If Len(Dir(TestLogPath)) > 0 Then
        Kill TestLogPath
    End If

    ' Создаём новый пустой файл
    F = FreeFile
    Open TestLogPath For Output As #F
    Close #F

    ' On Error GoTo 0 не вызываем — ошибки очистки игнорируем
End Sub