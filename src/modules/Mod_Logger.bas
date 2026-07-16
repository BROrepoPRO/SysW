Attribute VB_Name = "Mod_Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Модуль: Mod_Logger
' Назначение: Централизованное логирование с поддержкой ротации
' ============================================================

' Размер лога по умолчанию для ротации (в KB)
Private Const DEFAULT_MAX_LOG_SIZE_KB As Long = 100

' --------------------------------------------------------------------------
' GetLogPath
' Возвращает путь к файлу лога (log.txt рядом с книгой)
' --------------------------------------------------------------------------
Public Function GetLogPath() As String
    GetLogPath = ThisWorkbook.path & "\log.txt"
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
' Основной метод логирования
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
' Логирование ошибок с префиксом [ERROR]
' --------------------------------------------------------------------------
Public Sub WriteLogE(ByVal moduleName As String, ByVal message As String)
    Call WriteLog(moduleName, "[ERROR] " & message)
End Sub

' --------------------------------------------------------------------------
' RotateLogIfNeeded
' Ротация лога если превышен указанный размер (в KB)
' Если log.txt > maxSizeKB, переименовывает в log_old.txt и создаёт новый
' --------------------------------------------------------------------------
Public Sub RotateLogIfNeeded(ByVal maxSizeKB As Long)
    On Error Resume Next

    Dim LogPath As String
    Dim OldLogPath As String
    Dim F As Long
    Dim fso As Object

    LogPath = GetLogPath()
    OldLogPath = ThisWorkbook.path & "\log_old.txt"

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
' Очистка файла лога (удаление и создание пустого)
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