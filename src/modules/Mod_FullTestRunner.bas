Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_FullTestRunner
' Назначение: Набор технических тестов для проекта SysW
' Покрытие: TC-01 .. TC-50 (автоматические тесты) + TC-S1..TC-S3
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

    ' Подавляем MsgBox при тестовом запуске (COM-автоматизация)
    Mod_Constants.SilenceMsgBox = True

    Debug.Print "=============================================="
    Debug.Print "  Запуск набора тестов (TC-01..TC-50)"
    Debug.Print "=============================================="
    Debug.Print ""

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: START")

    ' Запуск групп тестов
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunUtilsTests START")
    RunUtilsTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunUtilsTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunLoggerTests START")
    RunLoggerTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunLoggerTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunUtilsEdgeTests START")
    RunUtilsEdgeTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunUtilsEdgeTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunLibNameTests START")
    RunLibNameTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunLibNameTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportVHTests START")
    RunImportVHTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportVHTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBTests START")
    RunModelDBTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunPickWorkTests START")
    RunPickWorkTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunPickWorkTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAutoMatchTests START")
    RunAutoMatchTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAutoMatchTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests START")
    RunSheetOpsTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSheetOpsTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAggregateNameTests START")
    RunAggregateNameTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunAggregateNameTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBReadTests START")
    RunModelDBReadTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunModelDBReadTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSQLiteTests START")
    RunSQLiteTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunSQLiteTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunOrderHeaderTests START")
    RunOrderHeaderTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunOrderHeaderTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportDataTests START")
    RunImportDataTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunImportDataTests END")

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunConstantsTests START")
    RunConstantsTests
    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: RunConstantsTests END")

    ' Финальный отчёт
    PrintFinalReport

    ' Запись результатов в ячейку Z1 листа main для Python
    On Error Resume Next
    WriteResultsToSheet
    If Err.Number <> 0 Then
        Call Mod_Logger.WriteLog("Mod_FullTestRunner", "WriteResultsToSheet error: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0

    ' Восстанавливаем показ MsgBox
    Mod_Constants.SilenceMsgBox = False

    Call Mod_Logger.WriteLog("Mod_FullTestRunner", "RunAllTests: END")
End Sub

' ============================================================
' WriteResultsToSheet — записывает результаты тестов в ячейку Z1
' листа main для чтения из Python (COM-клиента)
' ============================================================
Private Sub WriteResultsToSheet()
    Dim ReportMsg As String
    Dim ws As Worksheet

    ReportMsg = "Total=" & m_Total & ";Passed=" & m_Passed & _
                ";Failed=" & m_Failed & ";Skipped=" & m_Skipped & vbCrLf
    ReportMsg = ReportMsg & m_ResultsLog

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("main")
    If Not ws Is Nothing Then
        ws.Range("Z1").Value = ReportMsg
    End If
End Sub

' ============================================================
' GetTestResults — записывает результаты тестов в ячейку Z1 листа main
' для последующего чтения из Python (COM-клиента)
' ============================================================
Public Sub GetTestResults()
    Dim ReportMsg As String
    Dim ws As Worksheet

    ReportMsg = "Total=" & m_Total & ";Passed=" & m_Passed & _
                ";Failed=" & m_Failed & ";Skipped=" & m_Skipped & vbCrLf
    ReportMsg = ReportMsg & m_ResultsLog

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("main")
    If Not ws Is Nothing Then
        ws.Range("Z1").Value = ReportMsg
    End If
End Sub

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
    Set ws = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_MAIN)
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
    Set ws = Mod_Utils.GetSheetByName(ThisWorkbook, "NONEXISTENT")
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
    LogPath = Mod_Logger.GetLogPath()
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
    OldLogPath = ThisWorkbook.path & "\" & Mod_Constants.LOGS_DIR & "\log_old.txt"

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
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_LIBNAME)

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
    ' TC-14: ImportFromB2_UI — проверка вызова с пустым B4
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    If wsMain Is Nothing Then
        AddResult "TC-14", "ImportFromB2_UI с пустым B4", False, "Лист 'main' не найден"
    Else
        ' Сохраняем текущее значение B4
        oldB2 = Trim(CStr(wsMain.Range("B4").Value))

        ' Очищаем B4
        wsMain.Range("B4").Value = ""

        ' Подавляем MsgBox при тестировании
        Mod_Import.SilenceMsgBox = True

        ' Вызываем процедуру — она должна выйти без ошибки
        Call Mod_Import.ImportFromB2_UI

        Mod_Import.SilenceMsgBox = False

        If Err.number <> 0 Then
            AddResult "TC-14", "ImportFromB2_UI с пустым B4", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            AddResult "TC-14", "ImportFromB2_UI с пустым B4", True, ""
        End If

        ' Восстанавливаем B4 (с отключением событий, чтобы не сработал Worksheet_Change)
        Application.EnableEvents = False
        wsMain.Range("B4").Value = oldB2
        Application.EnableEvents = True
    End If

    Set wsMain = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты ModelDB (TC-31..TC-35)
' ============================================================
Private Sub RunModelDBTests()
    Dim basePath As String
    Dim filePath As String
    Dim groupExists As Boolean
    Dim wb As Workbook

    Debug.Print "--- Mod_ModelDB Tests ---"

    ' @test TC-31
    ' -------------------------------------------------------
    ' TC-31: GetModelDBBasePath возвращает непустую строку
    ' -------------------------------------------------------
    On Error Resume Next
    basePath = Mod_ModelDB.GetModelDBBasePath()
    If Err.number <> 0 Then
        AddResult "TC-31", "GetModelDBBasePath непустая строка", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim pathLenOk As Boolean
        Dim endsWithSlash As Boolean
        pathLenOk = (Len(basePath) > 0)
        endsWithSlash = (Right$(basePath, 1) = "\")
        AddResult "TC-31", "GetModelDBBasePath непустая строка", (pathLenOk And endsWithSlash), _
                  "длина=" & CStr(Len(basePath)) & ", endsWithSlash=" & CStr(endsWithSlash)
    End If
    On Error GoTo 0

    ' @test TC-32
    ' -------------------------------------------------------
    ' TC-32: GetModelGroupFilePath возвращает корректный путь
    ' -------------------------------------------------------
    On Error Resume Next
    filePath = Mod_ModelDB.GetModelGroupFilePath("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-32", "GetModelGroupFilePath путь для UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim endsWithXlsm As Boolean
        Dim endsWithXlsx As Boolean
        endsWithXlsm = (Right$(filePath, 8) = "UAZ.xlsm")
        endsWithXlsx = (Right$(filePath, 8) = "UAZ.xlsx")
        AddResult "TC-32", "GetModelGroupFilePath путь для UAZ", (endsWithXlsm Or endsWithXlsx), _
                  "путь=" & filePath
    End If
    On Error GoTo 0

    ' @test TC-33
    ' -------------------------------------------------------
    ' TC-33: ModelGroupFileExists для существующей группы
    ' -------------------------------------------------------
    On Error Resume Next
    groupExists = Mod_ModelDB.ModelGroupFileExists("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-33", "ModelGroupFileExists UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-33", "ModelGroupFileExists UAZ", groupExists, _
                  "Ожидалось True, получено " & CStr(groupExists)
    End If
    On Error GoTo 0

    ' @test TC-34
    ' -------------------------------------------------------
    ' TC-34: ModelGroupFileExists для несуществующей группы
    ' -------------------------------------------------------
    On Error Resume Next
    groupExists = Mod_ModelDB.ModelGroupFileExists("NonExistentGroup_Test")
    If Err.number <> 0 Then
        AddResult "TC-34", "ModelGroupFileExists несуществующая", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-34", "ModelGroupFileExists несуществующая", (Not groupExists), _
                  "Ожидалось False, получено " & CStr(groupExists)
    End If
    On Error GoTo 0

    ' @test TC-35
    ' -------------------------------------------------------
    ' TC-35: GetModelDataProvider возвращает работающий провайдер
    ' (SQLite-провайдер при наличии SysW.db; чтение тождеств без открытия xlsm)
    ' -------------------------------------------------------
    On Error Resume Next
    Dim prov As IModelDataProvider
    Call Mod_ModelDB.GetModelDataProvider(prov)
    If Err.number <> 0 Then
        AddResult "TC-35", "GetModelDataProvider провайдер", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim provOk As Boolean
        provOk = False
        If Not prov Is Nothing Then
            ' Провайдер должен уметь читать тождества работ (без COM-открытия файла)
            On Error Resume Next
            Dim testCol As Collection
            Set testCol = prov.GetWorkIdentities("UAZ")
            If Err.Number = 0 Then provOk = True
            Err.Clear
            On Error GoTo 0
        End If
        AddResult "TC-35", "GetModelDataProvider провайдер", provOk, _
                  "Провайдер не способен читать тождества из SysW.db"
    End If
    Set prov = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты PickWork (TC-36..TC-38)
' ============================================================
Private Sub RunPickWorkTests()
    Dim groupName As String
    Dim sheetName As String
    Dim wsMain As Worksheet
    Dim oldB14 As String

    Debug.Print "--- Mod_PickWork Tests ---"

    ' @test TC-36
    ' -------------------------------------------------------
    ' TC-36: GetGroupNameFromMain возвращает значение из B14
    ' -------------------------------------------------------
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    ' Сохраняем исходное значение B14 перед тестом
    oldB14 = Trim(CStr(wsMain.Range("B14").Value))

    ' Устанавливаем тестовое значение B14 перед вызовом
    wsMain.Range("B14").Value = "UAZ"

    On Error Resume Next
    groupName = Mod_PickWork.GetGroupNameFromMain()
    If Err.number <> 0 Then
        AddResult "TC-36", "GetGroupNameFromMain чтение B14", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-36", "GetGroupNameFromMain чтение B14", (Len(groupName) > 0), _
                  "Ожидалась непустая строка, получено '" & groupName & "'"
    End If
    On Error GoTo 0

    ' @test TC-37
    ' -------------------------------------------------------
    ' TC-37: GetWorkSheetName возвращает имя листа
    ' -------------------------------------------------------
    On Error Resume Next
    sheetName = Mod_PickWork.GetWorkSheetName("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-37", "GetWorkSheetName имя листа", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-37", "GetWorkSheetName имя листа", (Len(sheetName) > 0), _
                  "Ожидалась непустая строка, получено '" & sheetName & "'"
    End If
    On Error GoTo 0

    ' @test TC-38
    ' -------------------------------------------------------
    ' TC-38: Вызов кнопки РУЧ РАБ не вызывает ошибку
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_PickWork.PickWork_UI
    If Err.number <> 0 Then
        AddResult "TC-38", "PickWork_UI вызов без ошибки", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-38", "PickWork_UI вызов без ошибки", True, ""
    End If
    On Error GoTo 0

    ' Восстанавливаем исходное значение B14 (вместо ClearContents)
    If Len(oldB14) > 0 Then
        wsMain.Range("B14").Value = oldB14
    Else
        wsMain.Range("B14").ClearContents
    End If
    Set wsMain = Nothing

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты AutoMatch (TC-39..TC-44)
' ============================================================
Private Sub RunAutoMatchTests()
    Debug.Print "--- Mod_AutoMatch Tests ---"

    ' @test TC-39
    ' -------------------------------------------------------
    ' TC-39: AutoMatchWorks выполняется без ошибки
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_AutoMatch.AutoMatchWorks
    If Err.number <> 0 Then
        AddResult "TC-39", "AutoMatchWorks без ошибки", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-39", "AutoMatchWorks без ошибки", True, ""
    End If
    On Error GoTo 0

    ' @test TC-40
    ' -------------------------------------------------------
    ' TC-40: AutoMatchParts выполняется без ошибки
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_AutoMatch.AutoMatchParts
    If Err.number <> 0 Then
        AddResult "TC-40", "AutoMatchParts без ошибки", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-40", "AutoMatchParts без ошибки", True, ""
    End If
    On Error GoTo 0

    ' @test TC-41
    ' -------------------------------------------------------
    ' TC-41: HighlightNotFound (Private — пропущен)
    ' -------------------------------------------------------
    AddResult "TC-41", "HighlightNotFound (Private)", True, "", True, _
              "HighlightNotFound — Private процедура, недоступна для прямого вызова"

    ' @test TC-42
    ' -------------------------------------------------------
    ' TC-42: ClearHighlight (Private — пропущен)
    ' -------------------------------------------------------
    AddResult "TC-42", "ClearHighlight (Private)", True, "", True, _
              "ClearHighlight — Private процедура, недоступна для прямого вызова"

    ' @test TC-43
    ' -------------------------------------------------------
    ' TC-43: IsAllFound (Private — пропущен)
    ' -------------------------------------------------------
    AddResult "TC-43", "IsAllFound (Private)", True, "", True, _
              "IsAllFound — Private функция, недоступна для прямого вызова"

    ' @test TC-44
    ' -------------------------------------------------------
    ' TC-44: AutoMatchWorks без сохранения данных (пропущен)
    ' -------------------------------------------------------
    AddResult "TC-44", "AutoMatchWorks без изменения данных", True, "", True, _
              "AutoMatchWorks изменяет данные на листе — тест требует сохранения/восстановления, что небезопасно в автоматическом режиме"

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты SheetOps (TC-15..TC-18, TC-45)
' ============================================================
Private Sub RunSheetOpsTests()
    Dim Result As String
    Dim ws As Worksheet

    Debug.Print "--- Mod_SheetOps Tests ---"

    ' -------------------------------------------------------
    ' TC-15: ExtractNumberFromGRZ 'А123АН77' -> '123'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А123АН77")
    If Err.Number <> 0 Then
        AddResult "TC-15", "ExtractNumberFromGRZ 'А123АН77'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-15", "ExtractNumberFromGRZ 'А123АН77'", (Result = "123"), _
                  "Ожидалось '123', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-16: ExtractNumberFromGRZ 'А12АН34' (2 цифры) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А12АН34")
    If Err.Number <> 0 Then
        AddResult "TC-16", "ExtractNumberFromGRZ 'А12АН34'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-16", "ExtractNumberFromGRZ 'А12АН34'", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-17: ExtractNumberFromGRZ 'А1234АН77' (4 цифры) -> '1234'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("А1234АН77")
    If Err.Number <> 0 Then
        AddResult "TC-17", "ExtractNumberFromGRZ 'А1234АН77'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-17", "ExtractNumberFromGRZ 'А1234АН77'", (Result = "1234"), _
                  "Ожидалось '1234', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-18: ExtractNumberFromGRZ '' (пустая строка) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_SheetOps.ExtractNumberFromGRZ("")
    If Err.Number <> 0 Then
        AddResult "TC-18", "ExtractNumberFromGRZ пустая строка", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-18", "ExtractNumberFromGRZ пустая строка", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-45: SearchSheetByGRZ несуществующий ГРЗ -> Nothing
    ' (Mod_Constants.SilenceMsgBox уже = True в RunAllTests)
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = Mod_SheetOps.SearchSheetByGRZ("НЕСУЩЕСТВУЮЩИЙ")
    If Err.Number <> 0 Then
        AddResult "TC-45", "SearchSheetByGRZ несуществующий ГРЗ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-45", "SearchSheetByGRZ несуществующий ГРЗ", (ws Is Nothing), _
                  "Ожидалось Nothing, лист найден"
    End If
    Set ws = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты AggregateName (TC-19..TC-21)
' ============================================================
Private Sub RunAggregateNameTests()
    Dim Result As String

    Debug.Print "--- Mod_Constants GetAggregateName Tests ---"

    ' -------------------------------------------------------
    ' TC-19: GetAggregateName 'DIAG' -> 'Диагностика'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("DIAG")
    If Err.Number <> 0 Then
        AddResult "TC-19", "GetAggregateName 'DIAG'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-19", "GetAggregateName 'DIAG'", (Result = "Диагностика"), _
                  "Ожидалось 'Диагностика', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-20: GetAggregateName 'TO' -> 'ТО'
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("TO")
    If Err.Number <> 0 Then
        AddResult "TC-20", "GetAggregateName 'TO'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-20", "GetAggregateName 'TO'", (Result = "ТО"), _
                  "Ожидалось 'ТО', получено '" & Result & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-21: GetAggregateName 'XXX' (неизвестный) -> ''
    ' -------------------------------------------------------
    On Error Resume Next
    Result = Mod_Constants.GetAggregateName("XXX")
    If Err.Number <> 0 Then
        AddResult "TC-21", "GetAggregateName 'XXX'", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-21", "GetAggregateName 'XXX'", (Result = ""), _
                  "Ожидалась пустая строка, получено '" & Result & "'"
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты ModelDB Read (TC-22..TC-24)
' ============================================================
Private Sub RunModelDBReadTests()
    ' Чтение моделей выполняется из SysW.db через провайдер БЕЗ открытия
    ' base/models/*.xlsm через COM (исключает COM-зависание).
    Dim provider As IModelDataProvider
    Dim col As Collection

    Debug.Print "--- Mod_ModelDB Read Tests (SQLite) ---"

    On Error Resume Next
    Call Mod_ModelDB.GetModelDataProvider(provider)
    If Err.Number <> 0 Then
        AddResult "TC-22", "GetWorkIdentities UAZ", False, "Ошибка получения провайдера: " & Err.Description
        AddResult "TC-23", "GetPartIdentities UAZ", True, "", True, "Провайдер недоступен"
        AddResult "TC-24", "GetWorks UAZ", True, "", True, "Провайдер недоступен"
        Err.Clear
        Debug.Print ""
        Exit Sub
    End If
    On Error GoTo 0

    If provider Is Nothing Then
        AddResult "TC-22", "GetWorkIdentities UAZ", True, "", True, _
                  "Провайдер не создан (SysW.db/ODBC недоступны)"
        AddResult "TC-23", "GetPartIdentities UAZ", True, "", True, _
                  "Провайдер не создан (SysW.db/ODBC недоступны)"
        AddResult "TC-24", "GetWorks UAZ", True, "", True, _
                  "Провайдер не создан (SysW.db/ODBC недоступны)"
        Debug.Print ""
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' TC-22: GetWorkIdentities 'UAZ' — непустая коллекция WorkIdentity из SysW.db
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = provider.GetWorkIdentities("UAZ")
    If Err.Number <> 0 Then
        AddResult "TC-22", "GetWorkIdentities UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc22Ok As Boolean
        Dim tc22Reason As String
        tc22Ok = False
        If col.Count > 0 Then
            If TypeName(col(1)) = "WorkIdentity" Then
                Dim wi As WorkIdentity
                Set wi = col(1)
                tc22Ok = (Len(wi.OutArticle) > 0) And (Len(wi.Aggregate) > 0)
                If Not tc22Ok Then tc22Reason = "OutArticle/Aggregate пусты"
                Set wi = Nothing
            Else
                tc22Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc22Reason = "Коллекция пуста (нет тождеств работ в SysW.db для UAZ)"
        End If
        AddResult "TC-22", "GetWorkIdentities UAZ", tc22Ok, tc22Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-23: GetPartIdentities 'UAZ' — непустая коллекция PartIdentity из SysW.db
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = provider.GetPartIdentities("UAZ")
    If Err.Number <> 0 Then
        AddResult "TC-23", "GetPartIdentities UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc23Ok As Boolean
        Dim tc23Reason As String
        tc23Ok = False
        If col.Count > 0 Then
            If TypeName(col(1)) = "PartIdentity" Then
                Dim pi As PartIdentity
                Set pi = col(1)
                tc23Ok = (Len(pi.OutArticle) > 0) And (Len(pi.Aggregate) > 0)
                If Not tc23Ok Then tc23Reason = "OutArticle/Aggregate пусты"
                Set pi = Nothing
            Else
                tc23Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc23Reason = "Коллекция пуста (нет тождеств запчастей в SysW.db для UAZ)"
        End If
        AddResult "TC-23", "GetPartIdentities UAZ", tc23Ok, tc23Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-24: GetWorks 'UAZ' — непустая коллекция WorkEntry из SysW.db
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = provider.GetWorks("UAZ", Empty)
    If Err.Number <> 0 Then
        AddResult "TC-24", "GetWorks UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc24Ok As Boolean
        Dim tc24Reason As String
        tc24Ok = False
        If col.Count > 0 Then
            If TypeName(col(1)) = "WorkEntry" Then
                Dim we As WorkEntry
                Set we = col(1)
                tc24Ok = (Len(we.Code) > 0)
                If Not tc24Ok Then tc24Reason = "Code первого элемента пуст"
                Set we = Nothing
            Else
                tc24Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc24Reason = "Коллекция пуста (нет работ в SysW.db для UAZ)"
        End If
        AddResult "TC-24", "GetWorks UAZ", tc24Ok, tc24Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    Set provider = Nothing
    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты SQLite-провайдера (TC-S1..TC-S3, TC-47..TC-50)
' ============================================================
Private Sub RunSQLiteTests()
    Dim provider As IModelDataProvider
    Dim col As Collection

    Debug.Print "--- SQLite Provider Tests ---"

    ' -------------------------------------------------------
    ' TC-S1: Фабрика возвращает работающий провайдер при наличии SysW.db
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_ModelDB.GetModelDataProvider(provider)
    If Err.Number <> 0 Then
        AddResult "TC-S1", "Фабрика GetModelDataProvider", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim s1Ok As Boolean
        s1Ok = (Not provider Is Nothing)
        AddResult "TC-S1", "Фабрика GetModelDataProvider", s1Ok, _
                  "Ожидался провайдер, получен Nothing"
    End If
    On Error GoTo 0

    If provider Is Nothing Then
        AddResult "TC-S2", "GetWorks через провайдер", True, "", True, "Провайдер недоступен"
        AddResult "TC-S3", "Данные мигрированы в SysW.db", True, "", True, "Провайдер недоступен"
        Debug.Print ""
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' TC-S2: GetWorks через провайдер эквивалентен Excel-каталогу
    ' (проверяется чтение работ UAZ из SQLite без открытия xlsm)
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = provider.GetWorks("UAZ", Empty)
    If Err.Number <> 0 Then
        AddResult "TC-S2", "GetWorks через провайдер", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim s2Ok As Boolean
        Dim s2Reason As String
        s2Ok = (col.Count > 0)
        If Not s2Ok Then s2Reason = "Коллекция работ UAZ пуста в SysW.db"
        AddResult "TC-S2", "GetWorks через провайдер", s2Ok, s2Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-S3: Контрольные объёмы конвертера — данные мигрированы
    ' (для каждой группы есть работы; итоговое количество > 0)
    ' -------------------------------------------------------
    On Error Resume Next
    Dim allGroups As Collection
    Set allGroups = provider.GetAllModelGroups()
    If Err.Number <> 0 Then
        AddResult "TC-S3", "Данные мигрированы в SysW.db", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim s3Ok As Boolean
        Dim s3Reason As String
        s3Ok = (allGroups.Count >= 6)   ' ожидается не менее 6 групп моделей
        If Not s3Ok Then s3Reason = "Групп меньше 6: " & CStr(allGroups.Count)
        AddResult "TC-S3", "Данные мигрированы в SysW.db", s3Ok, s3Reason
    End If
    On Error GoTo 0

    ' ============================================================
    ' Расширенное покрытие SQLite-провайдера (Задача 3, v1.0.7):
    ' TC-47 GetParts (JOIN parts_catalog), TC-48 works с дублями,
    ' TC-49 чтение parts_catalog, TC-50 GetMatLibEntries (порядок).
    ' ============================================================

    ' Доступ к ADO-запросам того же соединения через конкретный класс.
    Dim sql As Mod_SQLiteDB
    On Error Resume Next
    Set sql = provider
    On Error GoTo 0

    If sql Is Nothing Then
        ' Активный провайдер — не SQLite (Excel fallback): тесты неприменимы.
        AddResult "TC-47", "GetParts через SQLite (JOIN parts_catalog)", True, "", True, _
                  "SQLite-провайдер недоступен"
        AddResult "TC-48", "works с дублями наименований без схлопывания", True, "", True, _
                  "SQLite-провайдер недоступен"
        AddResult "TC-49", "Чтение parts_catalog", True, "", True, _
                  "SQLite-провайдер недоступен"
        AddResult "TC-50", "GetMatLibEntries: детерминированный порядок", True, "", True, _
                  "SQLite-провайдер недоступен"
        Set provider = Nothing
        Debug.Print ""
        Exit Sub
    End If

    Dim noP As Variant
    noP = Array()

    ' ---------------------------------------------------------------
    ' TC-47: GetParts через SQLite — JOIN parts + parts_catalog
    ' (нормализация v1.0.6). Возвращает [code,name,unit,price,note];
    ' по одной записи на каждую привязку группы.
    ' ---------------------------------------------------------------
    On Error Resume Next
    Dim pc As Collection
    Set pc = provider.GetParts("UAZ", Empty)
    If Err.Number <> 0 Then
        AddResult "TC-47", "GetParts через SQLite (JOIN parts_catalog)", False, _
                  "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc47Ok As Boolean
        Dim tc47Reason As String
        Dim bindCount As Long
        bindCount = Val(sql.ExecuteScalar( _
            "SELECT COUNT(*) FROM parts WHERE group_name = 'UAZ'", noP))
        tc47Ok = (pc.Count > 0) And (pc.Count = bindCount)
        If Not tc47Ok Then
            tc47Reason = "parts.Count=" & CStr(pc.Count) & ", привязок=" & CStr(bindCount)
        Else
            Dim parr As Variant
            parr = pc(1)
            tc47Ok = (UBound(parr) = 4) And (Len(CStr(parr(0))) > 0) And (Len(CStr(parr(1))) > 0)
            If Not tc47Ok Then
                tc47Reason = "неверная форма записи (ожидается [code,name,unit,price,note])"
            End If
        End If
        AddResult "TC-47", "GetParts через SQLite (JOIN parts_catalog)", tc47Ok, tc47Reason
    End If
    Set pc = Nothing
    On Error GoTo 0

    ' ---------------------------------------------------------------
    ' TC-48: Чтение works с дублями наименований — суррогатный PK id
    ' (AUTOINCREMENT) сохраняет все дубли. Сверяем число строк GetWorks
    ' с числом строк таблицы для группы, где присутствуют дубли.
    ' ---------------------------------------------------------------
    On Error Resume Next
    Dim dupGroup As Variant
    dupGroup = sql.ExecuteScalar( _
        "SELECT group_name FROM works " & _
        "GROUP BY group_name, code, name HAVING COUNT(*) > 1 LIMIT 1", noP)
    If Err.Number <> 0 Then
        AddResult "TC-48", "works с дублями наименований без схлопывания", False, _
                  "Ошибка: " & Err.Description
        Err.Clear
    ElseIf IsEmpty(dupGroup) Then
        AddResult "TC-48", "works с дублями наименований без схлопывания", True, "", True, _
                  "В SysW.db нет групп с дублирующимися (code,name) работами"
    Else
        Dim tc48Ok As Boolean
        Dim tc48Reason As String
        Dim rawWorks As Long
        rawWorks = Val(sql.ExecuteScalar( _
            "SELECT COUNT(*) FROM works WHERE group_name = ?", Array(CStr(dupGroup))))
        Dim wc As Collection
        Set wc = provider.GetWorks(CStr(dupGroup), Empty)
        tc48Ok = (wc.Count = rawWorks)
        If Not tc48Ok Then
            tc48Reason = "GetWorks=" & CStr(wc.Count) & ", таблица=" & CStr(rawWorks) & _
                         " (дубли схлопнуты)"
        End If
        AddResult "TC-48", "works с дублями наименований без схлопывания", tc48Ok, tc48Reason
        Set wc = Nothing
    End If
    On Error GoTo 0

    ' ---------------------------------------------------------------
    ' TC-49: Чтение parts_catalog — глобальный уникальный каталог;
    ' все привязки parts ссылаются на существующие коды каталога.
    ' ---------------------------------------------------------------
    On Error Resume Next
    Dim catalogCount As Long
    catalogCount = Val(sql.ExecuteScalar( _
        "SELECT COUNT(*) FROM parts_catalog", noP))
    Dim orphanCount As Long
    orphanCount = Val(sql.ExecuteScalar( _
        "SELECT COUNT(*) FROM parts p LEFT JOIN parts_catalog c ON c.code = p.part_code " & _
        "WHERE c.code IS NULL", noP))
    If Err.Number <> 0 Then
        AddResult "TC-49", "Чтение parts_catalog", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc49Ok As Boolean
        Dim tc49Reason As String
        tc49Ok = (catalogCount > 0) And (orphanCount = 0)
        If Not tc49Ok Then
            tc49Reason = "каталог=" & CStr(catalogCount) & ", битых ссылок=" & CStr(orphanCount)
        End If
        AddResult "TC-49", "Чтение parts_catalog", tc49Ok, tc49Reason
    End If
    On Error GoTo 0

    ' ---------------------------------------------------------------
    ' TC-50: GetMatLibEntries через SQLite — возврат записей и
    ' детерминированный порядок (ORDER BY target_type, target_code);
    ' потребитель берёт первую запись нужного типа (mod_part для ЗЧ).
    ' ---------------------------------------------------------------
    On Error Resume Next
    Dim mlRS As Object
    Set mlRS = sql.ExecuteQuery( _
        "SELECT group_name, entry_code FROM matlib_entries " & _
        "WHERE target_type = 'mod_part' LIMIT 1", noP)
    If Err.Number <> 0 Then
        AddResult "TC-50", "GetMatLibEntries: детерминированный порядок", False, _
                  "Ошибка: " & Err.Description
        Err.Clear
    ElseIf mlRS Is Nothing Then
        AddResult "TC-50", "GetMatLibEntries: детерминированный порядок", True, "", True, _
                  "В SysW.db нет записей matlib_entries типа mod_part"
    ElseIf mlRS.EOF Then
        AddResult "TC-50", "GetMatLibEntries: детерминированный порядок", True, "", True, _
                  "В SysW.db нет записей matlib_entries типа mod_part"
    Else
        Dim mlGroup As String
        Dim mlCode As String
        mlGroup = CStr(mlRS.Fields(0).Value)
        mlCode = CStr(mlRS.Fields(1).Value)
        mlRS.Close

        Dim meCol1 As Collection
        Set meCol1 = provider.GetMatLibEntries(mlGroup, mlCode)
        Dim meCol2 As Collection
        Set meCol2 = provider.GetMatLibEntries(mlGroup, mlCode)

        Dim seq1 As String
        Dim seq2 As String
        Dim firstNeed1 As Long
        Dim firstNeed2 As Long
        Dim k As Long
        firstNeed1 = 0
        firstNeed2 = 0
        For k = 1 To meCol1.Count
            Dim earr1 As Variant
            earr1 = meCol1(k)
            seq1 = seq1 & UCase$(CStr(earr1(0))) & "|"
            If firstNeed1 = 0 And UCase$(CStr(earr1(0))) = "MOD_PART" Then firstNeed1 = k
        Next k
        For k = 1 To meCol2.Count
            Dim earr2 As Variant
            earr2 = meCol2(k)
            seq2 = seq2 & UCase$(CStr(earr2(0))) & "|"
            If firstNeed2 = 0 And UCase$(CStr(earr2(0))) = "MOD_PART" Then firstNeed2 = k
        Next k

        Dim tc50Ok As Boolean
        Dim tc50Reason As String
        tc50Ok = (meCol1.Count > 0) And (seq1 = seq2) And (firstNeed1 > 0) And _
                 (firstNeed1 = firstNeed2)
        If Not tc50Ok Then
            tc50Reason = "count=" & CStr(meCol1.Count) & ", детерминизм=" & CStr(seq1 = seq2) & _
                         ", first1=" & CStr(firstNeed1) & ", first2=" & CStr(firstNeed2)
        End If
        AddResult "TC-50", "GetMatLibEntries: детерминированный порядок", tc50Ok, tc50Reason
        Set meCol1 = Nothing
        Set meCol2 = Nothing
    End If
    Set mlRS = Nothing
    On Error GoTo 0

    Set provider = Nothing
    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты OrderHeader (TC-25..TC-28)
' ============================================================
Private Sub RunOrderHeaderTests()
    Dim wsMain As Worksheet
    Dim wsSpisok As Worksheet
    Dim wsModels As Worksheet
    Dim orderNum As Variant
    Dim Result As Boolean
    Dim Header As OrderHeader
    Dim savedB5toB17 As Variant
    Dim modelsLastRowBefore As Long
    Dim modelsLastRowAfter As Long
    Dim i As Long

    Debug.Print "--- Mod_OrderHeader Tests ---"

    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)
    Set wsSpisok = ThisWorkbook.Sheets(Mod_Constants.SHEET_SPISOK)
    Set wsModels = ThisWorkbook.Sheets(Mod_Constants.SHEET_MODELS)
    On Error GoTo 0

    If wsMain Is Nothing Or wsSpisok Is Nothing Or wsModels Is Nothing Then
        AddResult "TC-25", "FillHeaderFromOrder существующий заказ", True, "", True, _
                  "Лист main/spisok/models не найден"
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", True, "", True, _
                  "Лист main/spisok/models не найден"
        AddResult "TC-27", "FindOrder существующий заказ", True, "", True, _
                  "Лист spisok не найден"
        AddResult "TC-28", "FindOrder несуществующий заказ", True, "", True, _
                  "Лист spisok не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' Получаем реальный номер заказа: первый непустой номер из столбца A spisok
    orderNum = Empty
    For i = 2 To wsSpisok.Cells(wsSpisok.Rows.Count, 1).End(xlUp).Row
        If Not IsEmpty(wsSpisok.Cells(i, 1).Value) Then
            orderNum = wsSpisok.Cells(i, 1).Value
            Exit For
        End If
    Next i

    If IsEmpty(orderNum) Then
        AddResult "TC-25", "FillHeaderFromOrder существующий заказ", True, "", True, _
                  "В spisok нет ни одного номера заказа"
        AddResult "TC-27", "FindOrder существующий заказ", True, "", True, _
                  "В spisok нет ни одного номера заказа"
    Else
        ' =====================================================
        ' TC-25: FillHeaderFromOrder с существующим заказом
        ' =====================================================
        ' Сохраняем B5:B17 листа main
        savedB5toB17 = wsMain.Range("B5:B17").Value

        ' Сохраняем последнюю строку листа models (для детекции автодобавления)
        modelsLastRowBefore = wsModels.Cells(wsModels.Rows.Count, 1).End(xlUp).Row

        On Error Resume Next
        Application.EnableEvents = False
        Result = Mod_OrderHeader.FillHeaderFromOrder(orderNum)
        Application.EnableEvents = True
        If Err.Number <> 0 Then
            AddResult "TC-25", "FillHeaderFromOrder существующий заказ", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim tc25Ok As Boolean
            Dim tc25Reason As String
            tc25Ok = Result And (Len(Trim(CStr(wsMain.Range("B5").Value))) > 0)
            If Not Result Then
                tc25Reason = "Функция вернула False"
            ElseIf Len(Trim(CStr(wsMain.Range("B5").Value))) = 0 Then
                tc25Reason = "B5 не заполнена"
            End If
            AddResult "TC-25", "FillHeaderFromOrder существующий заказ", tc25Ok, tc25Reason
        End If
        On Error GoTo 0

        ' Восстанавливаем B5:B17
        Application.EnableEvents = False
        wsMain.Range("B5:B17").Value = savedB5toB17
        Application.EnableEvents = True

        ' Если FillHeaderFromOrder добавил модель в models — удаляем добавленную строку
        modelsLastRowAfter = wsModels.Cells(wsModels.Rows.Count, 1).End(xlUp).Row
        If modelsLastRowAfter > modelsLastRowBefore Then
            wsModels.Rows(modelsLastRowAfter).Delete
        End If

        ' =====================================================
        ' TC-27: FindOrder с существующим заказом
        ' =====================================================
        On Error Resume Next
        Result = Mod_OrderHeader.FindOrder(CStr(orderNum), Header)
        If Err.Number <> 0 Then
            AddResult "TC-27", "FindOrder существующий заказ", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim tc27Ok As Boolean
            Dim tc27Reason As String
            tc27Ok = Result And (Len(Trim(CStr(Header.OrderNumber))) > 0) _
                          And (Len(Trim(CStr(Header.ModelName))) > 0)
            If Not Result Then
                tc27Reason = "Функция вернула False"
            ElseIf Len(Trim(CStr(Header.OrderNumber))) = 0 Then
                tc27Reason = "Header.OrderNumber пуст"
            ElseIf Len(Trim(CStr(Header.ModelName))) = 0 Then
                tc27Reason = "Header.ModelName пуст"
            End If
            AddResult "TC-27", "FindOrder существующий заказ", tc27Ok, tc27Reason
        End If
        On Error GoTo 0
    End If

    ' =====================================================
    ' TC-26: FillHeaderFromOrder с несуществующим заказом
    ' =====================================================
    savedB5toB17 = wsMain.Range("B5:B17").Value

    On Error Resume Next
    Application.EnableEvents = False
    Result = Mod_OrderHeader.FillHeaderFromOrder(999999)
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc26Ok As Boolean
        Dim tc26Reason As String
        tc26Ok = (Not Result) And IsEmpty(wsMain.Range("B5").Value)
        If Result Then
            tc26Reason = "Функция вернула True для несуществующего заказа"
        ElseIf Not IsEmpty(wsMain.Range("B5").Value) Then
            tc26Reason = "B5 не очищена"
        End If
        AddResult "TC-26", "FillHeaderFromOrder несуществующий заказ", tc26Ok, tc26Reason
    End If
    On Error GoTo 0

    ' Восстанавливаем B5:B17
    Application.EnableEvents = False
    wsMain.Range("B5:B17").Value = savedB5toB17
    Application.EnableEvents = True

    ' =====================================================
    ' TC-28: FindOrder с несуществующим заказом
    ' =====================================================
    On Error Resume Next
    Result = Mod_OrderHeader.FindOrder("999999", Header)
    If Err.Number <> 0 Then
        AddResult "TC-28", "FindOrder несуществующий заказ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-28", "FindOrder несуществующий заказ", (Not Result), _
                  "Ожидалось False, получено " & CStr(Result)
    End If
    On Error GoTo 0

    Set wsMain = Nothing
    Set wsSpisok = Nothing
    Set wsModels = Nothing

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты ImportData (TC-29, TC-30)
' ============================================================
Private Sub RunImportDataTests()
    Dim wsMain As Worksheet
    Dim wsTemp As Worksheet
    Dim savedLtoN As Variant
    Dim savedXtoAA As Variant
    Dim tempName As String
    Dim lastRow As Long

    Debug.Print "--- Mod_Import Tests ---"

    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)
    On Error GoTo 0

    If wsMain Is Nothing Then
        AddResult "TC-29", "ImportDataToMain перенос данных", True, "", True, _
                  "Лист main не найден"
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", True, "", True, _
                  "Лист main не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' =====================================================
    ' TC-29: ImportDataToMain с временным листом-источником
    ' =====================================================
    ' Сохраняем L:N и X:AA листа main
    lastRow = Application.WorksheetFunction.Max( _
        wsMain.Cells(wsMain.Rows.Count, 12).End(xlUp).Row, _
        wsMain.Cells(wsMain.Rows.Count, 24).End(xlUp).Row)
    If lastRow < 4 Then lastRow = 4
    savedLtoN = wsMain.Range("L4:N" & lastRow).Value
    savedXtoAA = wsMain.Range("X4:AA" & lastRow).Value

    ' Создаём временный лист
    tempName = "TC29_Temp_" & Format$(Timer * 1000, "0")
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = tempName
    On Error GoTo 0

    If wsTemp Is Nothing Then
        AddResult "TC-29", "ImportDataToMain перенос данных", False, "Не удалось создать временный лист"
    Else
        ' --- Таблица "Выполненные работы" ---
        wsTemp.Cells(1, 1).Value = "Выполненные работы"
        wsTemp.Cells(2, 1).Value = "№"
        wsTemp.Cells(2, 2).Value = "№ кат."
        wsTemp.Cells(2, 3).Value = "Наименование"
        wsTemp.Cells(2, 4).Value = "Кол. оп."
        wsTemp.Cells(2, 5).Value = "Цена"
        wsTemp.Cells(2, 6).Value = "Норма"
        wsTemp.Cells(2, 7).Value = "н/ч"
        wsTemp.Cells(2, 8).Value = "Всего"
        wsTemp.Cells(2, 9).Value = "в т.ч. НДС"
        wsTemp.Cells(3, 1).Value = "1"
        wsTemp.Cells(3, 2).Value = "2"
        wsTemp.Cells(3, 3).Value = "3"
        wsTemp.Cells(3, 4).Value = "4"
        wsTemp.Cells(3, 5).Value = "5"
        wsTemp.Cells(3, 6).Value = "6"
        wsTemp.Cells(3, 7).Value = "7"
        wsTemp.Cells(3, 8).Value = "8"
        wsTemp.Cells(3, 9).Value = "9"
        ' Данные работ: D=Наименование, I=Всего, M=в т.ч. НДС
        wsTemp.Cells(4, 4).Value = "Тестовая работа 1"
        wsTemp.Cells(4, 9).Value = 100
        wsTemp.Cells(4, 13).Value = 20
        wsTemp.Cells(5, 4).Value = "Тестовая работа 2"
        wsTemp.Cells(5, 9).Value = 200
        wsTemp.Cells(5, 13).Value = 40
        wsTemp.Cells(6, 4).Value = "Итого работ"

        ' --- Таблица "Расходная накладная" ---
        wsTemp.Cells(8, 1).Value = "Расходная накладная"
        wsTemp.Cells(9, 1).Value = "№"
        wsTemp.Cells(9, 2).Value = "№ кат."
        wsTemp.Cells(9, 3).Value = "Наименование"
        wsTemp.Cells(9, 4).Value = "Кол-во"
        wsTemp.Cells(9, 5).Value = "Ед.изм."
        wsTemp.Cells(9, 6).Value = "Цена"
        wsTemp.Cells(9, 7).Value = "Всего"
        wsTemp.Cells(9, 8).Value = "в т.ч. НДС"
        wsTemp.Cells(10, 1).Value = "1"
        wsTemp.Cells(10, 2).Value = "2"
        wsTemp.Cells(10, 3).Value = "3"
        wsTemp.Cells(10, 4).Value = "4"
        wsTemp.Cells(10, 5).Value = "5"
        wsTemp.Cells(10, 6).Value = "6"
        wsTemp.Cells(10, 7).Value = "7"
        wsTemp.Cells(10, 8).Value = "8"
        ' Данные материалов (маппинг v1.0.6, docs/table.md разд. 3.2):
        ' B=№ кат.→X, C=Наименование→Y, D=Кол-во→Z, G=Всего→AA, M=в т.ч. НДС
        wsTemp.Cells(11, 2).Value = "ТК-001"
        wsTemp.Cells(11, 3).Value = "Тестовая запчасть 1"
        wsTemp.Cells(11, 4).Value = 2
        wsTemp.Cells(11, 7).Value = 50
        wsTemp.Cells(11, 13).Value = 10
        wsTemp.Cells(12, 2).Value = "ТК-002"
        wsTemp.Cells(12, 3).Value = "Тестовая запчасть 2"
        wsTemp.Cells(12, 4).Value = 3
        wsTemp.Cells(12, 7).Value = 60
        wsTemp.Cells(12, 13).Value = 12
        wsTemp.Cells(13, 2).Value = "Итого"

        ' Подавляем MsgBox для Mod_Import
        Mod_Import.SilenceMsgBox = True

        On Error Resume Next
        Call Mod_Import.ImportDataToMain(wsTemp)
        If Err.Number <> 0 Then
            AddResult "TC-29", "ImportDataToMain перенос данных", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            ' ВРЕМЕННАЯ ЗАГЛУШКА TC-29 (по указанию юзера):
            ' Прямая проверка X4=№ кат. некорректна: входящие запчасти могут иметь
            ' ПУСТОЙ артикул (№ кат.), но не пустое наименование. Подбор ЗЧ идёт
            ' сначала по артикулу, затем по наименованию; подбор работ — только
            ' по входящему наименованию. Корректная проверка будет восстановлена
            ' после согласования бизнес-логики.
            AddResult "TC-29", "ImportDataToMain перенос данных", True, "", True, _
                      "Временная заглушка: проверка прямого переноса запчастей отложена"
        End If
        On Error GoTo 0

        Mod_Import.SilenceMsgBox = False

        ' Удаляем временный лист ([S3])
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If

    ' Восстанавливаем L:N и X:AA листа main
    Application.EnableEvents = False
    wsMain.Range("L4:N" & lastRow).Value = savedLtoN
    wsMain.Range("X4:AA" & lastRow).Value = savedXtoAA
    Application.EnableEvents = True

    ' =====================================================
    ' TC-30: ImportSheet с несуществующим ГРЗ
    ' =====================================================
    Mod_Import.SilenceMsgBox = True

    On Error Resume Next
    Call Mod_Import.ImportSheet("НЕСУЩЕСТВУЮЩИЙ")
    If Err.Number <> 0 Then
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-30", "ImportSheet несуществующий ГРЗ", True, ""
    End If
    On Error GoTo 0

    Mod_Import.SilenceMsgBox = False

    Set wsMain = Nothing
    Set wsTemp = Nothing

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты Constants (TC-46)
' ============================================================
Private Sub RunConstantsTests()
    Dim wsLib As Worksheet
    Dim lastRowBefore As Long
    Dim lastRowAfter As Long
    Dim addedRow As Long
    Dim i As Long

    Debug.Print "--- Mod_Constants AddWorkEntry Tests ---"

    On Error Resume Next
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_LIBNAME)
    On Error GoTo 0

    If wsLib Is Nothing Then
        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", True, "", True, _
                  "Лист libname не найден"
        Debug.Print ""
        Exit Sub
    End If

    ' Сохраняем последнюю строку до теста
    lastRowBefore = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row

    ' Вызываем AddWorkEntry
    On Error Resume Next
    Call Mod_Constants.AddWorkEntry()
    If Err.Number <> 0 Then
        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        lastRowAfter = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row
        addedRow = lastRowAfter

        ' Проверяем, что запись work.xlsm присутствует в последней строке
        Dim tc46Ok As Boolean
        Dim tc46Reason As String
        tc46Ok = (Trim(CStr(wsLib.Cells(addedRow, 1).Value)) = "work.xlsm")
        If Not tc46Ok Then
            tc46Reason = "Последняя строка не содержит work.xlsm: '" & _
                         CStr(wsLib.Cells(addedRow, 1).Value) & "'"
        End If

        ' Проверяем идемпотентность: повторный вызов не должен добавить дубликат
        If tc46Ok Then
            Call Mod_Constants.AddWorkEntry()
            Dim lastRowAfter2 As Long
            lastRowAfter2 = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row
            If lastRowAfter2 > lastRowAfter Then
                tc46Ok = False
                tc46Reason = "Повторный вызов добавил дубликат (идемпотентность нарушена)"
            End If
        End If

        AddResult "TC-46", "AddWorkEntry добавление work.xlsm", tc46Ok, tc46Reason

        ' Восстанавливаем libname: если строка была добавлена тестом — удаляем её
        If lastRowAfter > lastRowBefore Then
            Application.DisplayAlerts = False
            wsLib.Rows(lastRowAfter).Delete
            Application.DisplayAlerts = True
        End If
    End If
    On Error GoTo 0

    Set wsLib = Nothing

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
    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "  ИТОГОВЫЙ ОТЧЁТ"
    Debug.Print "=============================================="
    Debug.Print "  Всего: " & m_Total
    Debug.Print "  Пройдено: " & m_Passed
    Debug.Print "  Провалено: " & m_Failed
    Debug.Print "  Пропущено: " & m_Skipped
    Debug.Print "=============================================="
End Sub


' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' RunAllTests_UI
' Запускает все тесты (TC-01..TC-46) и показывает результат
' --------------------------------------------------------------------------
Public Sub RunAllTests_UI()
    On Error GoTo ErrHandler

    Call RunAllTests

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RunAllTests_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RunAllTests_UI: " & Err.Description)
End Sub
