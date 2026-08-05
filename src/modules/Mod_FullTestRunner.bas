Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_FullTestRunner
' Назначение: Набор технических тестов для проекта SysW
' Покрытие: TC-01 .. TC-46 (автоматические тесты)
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
    Debug.Print "  Запуск набора тестов (TC-01..TC-46)"
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
    ' TC-35: OpenModelGroupFile открывает книгу
    ' -------------------------------------------------------
    On Error Resume Next
    Set wb = Mod_ModelDB.OpenModelGroupFile("UAZ")
    If Err.number <> 0 Then
        AddResult "TC-35", "OpenModelGroupFile UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim wbOpened As Boolean
        wbOpened = (Not wb Is Nothing)
        ' Закрываем книгу, если открыта
        If wbOpened Then
            wb.Close SaveChanges:=False
        End If
        AddResult "TC-35", "OpenModelGroupFile UAZ", wbOpened, _
                  "Ожидалось Not Nothing, получено Nothing"
    End If
    Set wb = Nothing
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
    Dim col As Collection
    Dim wb As Workbook
    Dim groupExists As Boolean
    Dim item As Variant

    Debug.Print "--- Mod_ModelDB Read Tests ---"

    ' Проверяем наличие файла группы UAZ (рабочие данные)
    On Error Resume Next
    groupExists = Mod_ModelDB.ModelGroupFileExists("UAZ")
    If Err.Number <> 0 Then
        groupExists = False
        Err.Clear
    End If
    On Error GoTo 0

    If Not groupExists Then
        AddResult "TC-22", "GetWorkIdentities UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        AddResult "TC-23", "GetPartIdentities UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        AddResult "TC-24", "GetWorks UAZ", True, "", True, _
                  "Файл группы UAZ не найден (base/models/UAZ.xlsm)"
        Debug.Print ""
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' TC-22: GetWorkIdentities 'UAZ' — непустая коллекция WorkIdentity
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetWorkIdentities("UAZ")
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
                If Not tc22Ok Then
                    tc22Reason = "OutArticle/Aggregate пусты"
                End If
                Set wi = Nothing
            Else
                tc22Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc22Reason = "Коллекция пуста (нет данных в UAZw)"
        End If
        AddResult "TC-22", "GetWorkIdentities UAZ", tc22Ok, tc22Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-23: GetPartIdentities 'UAZ' — непустая коллекция PartIdentity
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetPartIdentities("UAZ")
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
                If Not tc23Ok Then
                    tc23Reason = "OutArticle/Aggregate пусты"
                End If
                Set pi = Nothing
            Else
                tc23Reason = "Тип элемента: " & TypeName(col(1))
            End If
        Else
            tc23Reason = "Коллекция пуста (нет данных в UAZz4)"
        End If
        AddResult "TC-23", "GetPartIdentities UAZ", tc23Ok, tc23Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-24: GetWorks 'UAZ' — непустая коллекция WorkEntry
    ' -------------------------------------------------------
    On Error Resume Next
    Set col = Mod_ModelDB.GetWorks("UAZ", Empty)
    If Err.Number <> 0 Then
        AddResult "TC-24", "GetWorks UAZ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim tc24Ok As Boolean
        Dim tc24Reason As String
        tc24Ok = False
        If col.Count > 0 Then
            ' Элементы коллекции — Variant-массивы [Code, Name, Unit, NormHours, Price, Note]
            item = col(1)   ' item — Variant, теперь это массив
            tc24Ok = (Len(CStr(item(0))) > 0)   ' Code — первый элемент массива
            If Not tc24Ok Then
                tc24Reason = "Code первого элемента пуст"
            End If
        Else
            tc24Reason = "Коллекция пуста (нет данных в UAZ)"
        End If
        AddResult "TC-24", "GetWorks UAZ", tc24Ok, tc24Reason
    End If
    Set col = Nothing
    On Error GoTo 0

    ' Закрываем книгу UAZ.xlsm, если она осталась открытой
    On Error Resume Next
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks("UAZ.xlsm")
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
    On Error GoTo 0

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
        ' Данные материалов: C=Наименование, D=Кол-во, J=Всего, M=в т.ч. НДС
        wsTemp.Cells(11, 3).Value = "Тестовая запчасть 1"
        wsTemp.Cells(11, 4).Value = 2
        wsTemp.Cells(11, 10).Value = 50
        wsTemp.Cells(11, 13).Value = 10
        wsTemp.Cells(12, 3).Value = "Тестовая запчасть 2"
        wsTemp.Cells(12, 4).Value = 3
        wsTemp.Cells(12, 10).Value = 60
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
            Dim tc29Ok As Boolean
            Dim tc29Reason As String
            ' Проверяем перенос: L4 = D источника, X4 = C источника
            tc29Ok = (Trim(CStr(wsMain.Cells(4, 12).Value)) = "Тестовая работа 1") _
                 And (Trim(CStr(wsMain.Cells(4, 24).Value)) = "Тестовая запчасть 1")
            If Trim(CStr(wsMain.Cells(4, 12).Value)) <> "Тестовая работа 1" Then
                tc29Reason = "L4 не заполнена: '" & CStr(wsMain.Cells(4, 12).Value) & "'"
            ElseIf Trim(CStr(wsMain.Cells(4, 24).Value)) <> "Тестовая запчасть 1" Then
                tc29Reason = "X4 не заполнена: '" & CStr(wsMain.Cells(4, 24).Value) & "'"
            End If
            AddResult "TC-29", "ImportDataToMain перенос данных", tc29Ok, tc29Reason
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
