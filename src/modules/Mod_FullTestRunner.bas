Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_FullTestRunner
' Назначение: Набор технических тестов для проекта SysW
' Покрытие: TC-01 .. TC-13 (автоматические тесты)
' ============================================================

' ---- Счётчики результатов ----
Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

' ---- Накопитель результатов для GetTestResults ----
Private m_ResultsLog As String

' ============================================================
' Главная процедура: запуск всех тестов
' ============================================================
Public Sub RunAllTests()
    ' Инициализация счётчиков
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
    m_ResultsLog = ""

    Debug.Print "=============================================="
    Debug.Print "  Запуск набора тестов (TC-01..TC-13)"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Запуск групп тестов
    RunUtilsTests
    RunLoggerTests
    RunUtilsEdgeTests
    RunLibNameTests
    RunImportVHTests

    ' Финальный отчёт
    PrintFinalReport
End Sub

' ============================================================
' GetTestResults — возвращает строку с результатами для Python
' ============================================================
Public Function GetTestResults() As String
    Dim ReportMsg As String
    ReportMsg = "Total=" & m_Total & ";Passed=" & m_Passed & _
                ";Failed=" & m_Failed & ";Skipped=" & m_Skipped & vbCrLf
    ReportMsg = ReportMsg & m_ResultsLog
    GetTestResults = ReportMsg
End Function

' ============================================================
' Группа: тесты Utils (TC-01..TC-08)
' ============================================================
Private Sub RunUtilsTests()
    Dim LogPath As String
    Dim Result As Boolean
    Dim PathResult As String
    Dim UserResult As String

    Debug.Print "--- Mod_Utils Tests ---"

    ' -------------------------------------------------------
    ' TC-01: FileExists с существующим файлом
    ' -------------------------------------------------------
    On Error Resume Next
    Result = FileExists("C:\Windows\notepad.exe")
    If Err.number <> 0 Then
        AddResult "TC-01", "FileExists с существующим файлом", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-01", "FileExists с существующим файлом", (Result = True), _
                  "Ожидалось True, получено " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-02: FileExists с несуществующим файлом
    ' -------------------------------------------------------
    On Error Resume Next
    Result = FileExists("C:\nonexistent_file_12345.txt")
    If Err.number <> 0 Then
        AddResult "TC-02", "FileExists с несуществующим файлом", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-02", "FileExists с несуществующим файлом", (Result = False), _
                  "Ожидалось False, получено " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-03: FormatDateSQL с корректной датой
    ' -------------------------------------------------------
    On Error Resume Next
    Dim FmtResult As String
    FmtResult = FormatDateSQL(DateSerial(2026, 7, 12))
    If Err.number <> 0 Then
        AddResult "TC-03", "FormatDateSQL с корректной датой", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-03", "FormatDateSQL с корректной датой", (FmtResult = "2026-07-12"), _
                  "Ожидалось '2026-07-12', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-04: FormatDateSQL с нулевой датой
    ' -------------------------------------------------------
    On Error Resume Next
    FmtResult = FormatDateSQL(0)
    If Err.number <> 0 Then
        AddResult "TC-04", "FormatDateSQL с нулевой датой", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-04", "FormatDateSQL с нулевой датой", (FmtResult = "1899-12-30"), _
                  "Ожидалось '1899-12-30', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-05: GetSheetByName существующий
    ' -------------------------------------------------------
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "main")
    If Err.number <> 0 Then
        AddResult "TC-05", "GetSheetByName существующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-05", "GetSheetByName существующий лист", (Not ws Is Nothing), _
                  "Ожидалось Not Nothing, получено Nothing"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-06: GetSheetByName несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = GetSheetByName(ThisWorkbook, "NONEXISTENT")
    If Err.number <> 0 Then
        AddResult "TC-06", "GetSheetByName несуществующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-06", "GetSheetByName несуществующий лист", (ws Is Nothing), _
                  "Ожидалось Nothing, лист найден"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-07: WriteLog
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_Utils.WriteLog("Mod_FullTestRunner: выполнение проверки TC-07")
    LogPath = ThisWorkbook.path & "\log.txt"
    If Err.number <> 0 Then
        AddResult "TC-07", "WriteLog запись в лог", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-07", "WriteLog запись в лог", FileExists(LogPath), _
                  "файл лога не найден: " & LogPath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-08: GetWorkbookPath / GetCurrentUser
    ' -------------------------------------------------------
    On Error Resume Next
    PathResult = GetWorkbookPath()
    UserResult = GetCurrentUser()
    If Err.number <> 0 Then
        AddResult "TC-08", "GetWorkbookPath / GetCurrentUser", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim PathOk As Boolean
        Dim UserOk As Boolean
        PathOk = (Len(PathResult) > 0)
        UserOk = (Len(UserResult) > 0)
        AddResult "TC-08", "GetWorkbookPath / GetCurrentUser", (PathOk And UserOk), _
                  "Path пустой=" & CStr(Not PathOk) & ", User пустой=" & CStr(Not UserOk)
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты Logger (TC-09, TC-10, TC-11)
' ============================================================
Private Sub RunLoggerTests()
    Dim LogPath As String
    Dim OldLogPath As String
    Dim F As Long
    Dim i As Long

    Debug.Print "--- Mod_Logger Tests ---"

    ' -------------------------------------------------------
    ' TC-09: WriteLog — запись в лог-файл
    ' -------------------------------------------------------
    On Error Resume Next
    LogPath = Mod_Logger.GetLogPath()

    ' Очищаем лог перед тестом
    Call Mod_Logger.ClearLog

    ' Записываем сообщение
    Call Mod_Logger.WriteLog("TestModule", "Test message")

    If Err.number <> 0 Then
        AddResult "TC-09", "WriteLog запись в лог-файл", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim FileExistsAfterWrite As Boolean
        Dim FileContainsMessage As Boolean
        Dim FileContent As String
        Dim lineStr As String

        FileExistsAfterWrite = (Len(Dir(LogPath)) > 0)

        ' Читаем файл и ищем "Test message"
        FileContainsMessage = False
        If FileExistsAfterWrite Then
            F = FreeFile
            Open LogPath For Input As #F
            Do While Not EOF(F)
                Line Input #F, lineStr
                If InStr(1, lineStr, "Test message", vbTextCompare) > 0 Then
                    FileContainsMessage = True
                    Exit Do
                End If
            Loop
            Close #F
        End If

        Dim Tc09Passed As Boolean
        Tc09Passed = FileExistsAfterWrite And FileContainsMessage
        Dim Tc09Reason As String
        If Not FileExistsAfterWrite Then
            Tc09Reason = "Файл лога не создан: " & LogPath
        ElseIf Not FileContainsMessage Then
            Tc09Reason = "Файл лога не содержит 'Test message'"
        End If
        AddResult "TC-09", "WriteLog запись в лог-файл", Tc09Passed, Tc09Reason
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-10: RotateLogIfNeeded — ротация лога
    ' -------------------------------------------------------
    On Error Resume Next
    LogPath = Mod_Logger.GetLogPath()
    OldLogPath = ThisWorkbook.path & "\log_old.txt"

    ' Очищаем лог и старый лог
    Call Mod_Logger.ClearLog
    If Len(Dir(OldLogPath)) > 0 Then
        Kill OldLogPath
    End If

    ' Записываем много данных (больше 1 KB)
    For i = 1 To 50
        Call Mod_Logger.WriteLog("TestModule", "Test message for rotation #" & i & " - padding data to exceed 1KB limit")
    Next i

    ' Вызываем ротацию с лимитом 1 KB
    Call Mod_Logger.RotateLogIfNeeded(1)

    If Err.number <> 0 Then
        AddResult "TC-10", "RotateLogIfNeeded ротация лога", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim OldLogExists As Boolean
        OldLogExists = (Len(Dir(OldLogPath)) > 0)
        AddResult "TC-10", "RotateLogIfNeeded ротация лога", OldLogExists, _
                  "Старый лог не найден: " & OldLogPath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-11: ClearLog — очистка лога
    ' -------------------------------------------------------
    On Error Resume Next
    LogPath = Mod_Logger.GetLogPath()

    ' Записываем что-то в лог
    Call Mod_Logger.WriteLog("TestModule", "Message before clear")

    ' Очищаем лог
    Call Mod_Logger.ClearLog

    If Err.number <> 0 Then
        AddResult "TC-11", "ClearLog очистка лога", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim LogExistsAfterClear As Boolean
        Dim LogIsEmpty As Boolean

        LogExistsAfterClear = (Len(Dir(LogPath)) > 0)

        ' Проверяем, что файл пуст
        LogIsEmpty = True
        If LogExistsAfterClear Then
            F = FreeFile
            Open LogPath For Input As #F
            If Not EOF(F) Then
                ' Если есть хоть одна строка — не пуст
                LogIsEmpty = False
            End If
            Close #F
        End If

        ' Либо файл удалён, либо пуст — оба варианта acceptable
        Dim Tc11Passed As Boolean
        Tc11Passed = (Not LogExistsAfterClear) Or LogIsEmpty
        AddResult "TC-11", "ClearLog очистка лога", Tc11Passed, _
                  "Файл существует=" & CStr(LogExistsAfterClear) & ", пуст=" & CStr(LogIsEmpty)
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: граничные случаи Utils (TC-12)
' ============================================================
Private Sub RunUtilsEdgeTests()
    Dim FmtResult As String

    Debug.Print "--- Mod_Utils Edge Cases ---"

    ' -------------------------------------------------------
    ' TC-12: FormatDateSQL граничные случаи
    ' -------------------------------------------------------

    ' Кейс 1: Пустая строка (невалидный вызов — FormatDateSQL ожидает Date, а не String)
    ' Пропускаем, т.к. это ошибка типа на уровне выполнения VBA
    AddResult "TC-12", "FormatDateSQL граничные случаи (пустая строка)", True, "", True, _
              "FormatDateSQL принимает Date, пустая строка — невалидный аргумент"

    ' Кейс 2: Только год
    On Error Resume Next
    FmtResult = FormatDateSQL(DateSerial(2026, 1, 1))
    If Err.number <> 0 Then
        AddResult "TC-12", "FormatDateSQL граничные случаи (только год)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-12", "FormatDateSQL граничные случаи (только год)", (FmtResult = "2026-01-01"), _
                  "Ожидалось '2026-01-01', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' Кейс 3: Дата с временем
    On Error Resume Next
    FmtResult = FormatDateSQL(DateSerial(2026, 7, 12) + TimeSerial(14, 30, 0))
    If Err.number <> 0 Then
        AddResult "TC-12", "FormatDateSQL граничные случаи (дата+время)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-12", "FormatDateSQL граничные случаи (дата+время)", (FmtResult = "2026-07-12"), _
                  "Ожидалось '2026-07-12', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты LibName (TC-13)
' ============================================================
Private Sub RunLibNameTests()
    Dim wsLib As Worksheet
    Dim entryCount As Long

    Debug.Print "--- Mod_Constants (libname) Tests ---"

    ' -------------------------------------------------------
    ' TC-13: InitLibName — заполнение листа libname
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, "libname")

    If wsLib Is Nothing Then
        AddResult "TC-13", "InitLibName заполнение libname", False, "Лист 'libname' не найден"
    Else
        ' Очищаем лист перед тестом, чтобы InitLibName гарантированно заполнил его
        wsLib.Rows("2:" & wsLib.Rows.Count).ClearContents

        ' Проверяем, что лист libname содержит данные (непустая строка 2)
        If IsEmpty(wsLib.Cells(2, 1).Value) Then
            ' Лист пуст — вызываем InitLibName для заполнения
            Call Mod_Constants.InitLibName
        End If

        If Err.number <> 0 Then
            AddResult "TC-13", "InitLibName заполнение libname", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            ' Проверяем, что данные появились
            Dim hasSpisokColModel As Boolean
            Dim hasModelsColModelName As Boolean
            Dim hasModelsColHrpr As Boolean
            Dim hasModelsColGroup As Boolean
            Dim hasZ4 As Boolean
            Dim i As Long

            hasSpisokColModel = False
            hasModelsColModelName = False
            hasModelsColHrpr = False
            hasModelsColGroup = False
            hasZ4 = False

            ' Ищем последнюю заполненную строку
            entryCount = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row

            ' Проверяем наличие ключевых записей
            For i = 2 To entryCount
                Dim key As String
                key = Trim(CStr(wsLib.Cells(i, 1).Value))
                If key = "spisok_col_model" Then hasSpisokColModel = True
                If key = "models_col_model_name" Then hasModelsColModelName = True
                If key = "models_col_hrpr" Then hasModelsColHrpr = True
                If key = "models_col_group" Then hasModelsColGroup = True
                If key = "z4" Then hasZ4 = True
            Next i

            Dim Tc13Passed As Boolean
            Tc13Passed = hasSpisokColModel And hasModelsColModelName _
                      And hasModelsColHrpr And hasModelsColGroup And hasZ4
            Dim Tc13Reason As String
            If Not hasSpisokColModel Then Tc13Reason = Tc13Reason & "нет spisok_col_model; "
            If Not hasModelsColModelName Then Tc13Reason = Tc13Reason & "нет models_col_model_name; "
            If Not hasModelsColHrpr Then Tc13Reason = Tc13Reason & "нет models_col_hrpr; "
            If Not hasModelsColGroup Then Tc13Reason = Tc13Reason & "нет models_col_group; "
            If Not hasZ4 Then Tc13Reason = Tc13Reason & "нет z4; "

            AddResult "TC-13", "InitLibName заполнение libname", Tc13Passed, Tc13Reason
        End If
    End If

    Set wsLib = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты ImportVH (TC-14)
' ============================================================
Private Sub RunImportVHTests()
    Dim wsMain As Worksheet
    Dim oldB2 As String

    Debug.Print "--- Mod_Import ImportFromB2_UI Tests ---"

    ' -------------------------------------------------------
    ' TC-14: ImportFromB2_UI — проверка вызова с пустым B2
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets("мэйн")

    If wsMain Is Nothing Then
        AddResult "TC-14", "ImportFromB2_UI с пустым B2", False, "Лист 'мэйн' не найден"
    Else
        ' Сохраняем текущее значение B2
        oldB2 = Trim(CStr(wsMain.Range("B2").Value))

        ' Очищаем B2
        wsMain.Range("B2").Value = ""

        ' Вызываем процедуру — она должна показать MsgBox и выйти без ошибки
        Call Mod_Import.ImportFromB2_UI

        If Err.number <> 0 Then
            AddResult "TC-14", "ImportFromB2_UI с пустым B2", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            AddResult "TC-14", "ImportFromB2_UI с пустым B2", True, ""
        End If

        ' Восстанавливаем B2
        wsMain.Range("B2").Value = oldB2
    End If

    Set wsMain = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Вспомогательные функции
' ============================================================

' Добавляет результат теста в статистику и выводит в Immediate Window
Private Sub AddResult(testId As String, testName As String, _
                      passed As Boolean, Optional failReason As String = "", _
                      Optional skipped As Boolean = False, Optional skipReason As String = "")
    m_Total = m_Total + 1

    If skipped Then
        m_Skipped = m_Skipped + 1
        Debug.Print "[" & testId & "] " & ChrW(&H26A0) & " " & testName & ": SKIP (" & skipReason & ")"
        m_ResultsLog = m_ResultsLog & "[" & testId & "] SKIP: " & testName & " (" & skipReason & ")" & vbCrLf
    ElseIf passed Then
        m_Passed = m_Passed + 1
        Debug.Print "[" & testId & "] " & ChrW(&H2713) & " " & testName & ": PASS"
        m_ResultsLog = m_ResultsLog & "[" & testId & "] PASS: " & testName & vbCrLf
    Else
        m_Failed = m_Failed + 1
        If failReason <> "" Then
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL - " & failReason
            m_ResultsLog = m_ResultsLog & "[" & testId & "] FAIL: " & testName & " - " & failReason & vbCrLf
        Else
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL"
            m_ResultsLog = m_ResultsLog & "[" & testId & "] FAIL: " & testName & vbCrLf
        End If
    End If
End Sub

' Вывод финального отчёта
Private Sub PrintFinalReport()
    Dim ReportMsg As String

    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "  ИТОГОВЫЙ ОТЧЁТ"
    Debug.Print "=============================================="
    Debug.Print "  Всего: " & m_Total
    Debug.Print "  Пройдено: " & m_Passed
    Debug.Print "  Провалено: " & m_Failed
    Debug.Print "  Пропущено: " & m_Skipped
    Debug.Print "=============================================="

    ' Форматирование сообщения для MsgBox
    ReportMsg = "РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ:" & vbCrLf & vbCrLf
    ReportMsg = ReportMsg & "  Всего: " & m_Total & vbCrLf
    ReportMsg = ReportMsg & "  Пройдено: " & m_Passed & vbCrLf
    ReportMsg = ReportMsg & "  Провалено: " & m_Failed & vbCrLf
    ReportMsg = ReportMsg & "  Пропущено: " & m_Skipped & vbCrLf & vbCrLf

    If m_Failed = 0 Then
        ReportMsg = ReportMsg & "Все тесты успешно пройдены!"
    Else
        ReportMsg = ReportMsg & "Обнаружены ошибки! Проверьте Immediate Window для деталей."
    End If

    MsgBox ReportMsg, vbInformation + vbOKOnly, "Mod_FullTestRunner"
End Sub


' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' RunAllTests_UI
' Запускает все тесты (TC-01..TC-13) и показывает результат
' --------------------------------------------------------------------------
Public Sub RunAllTests_UI()
    On Error GoTo ErrHandler

    Call RunAllTests

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RunAllTests_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RunAllTests_UI: " & Err.Description)
End Sub
