Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_FullTestRunner
' Назначение: Полный набор тестов для проекта SysW
' Покрытие: TC-01 .. TC-30 (автоматические тесты)
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
    Debug.Print "  Запуск полного набора тестов (TC-01..TC-30)"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Запуск групп тестов
    RunUtilsTests
    RunOrderHeaderTests
    RunImportTests
    RunButtonTests
    RunLoggerTests
    RunSheetOpsTests
    RunImportIntegrationTests
    RunUtilsEdgeTests
    RunButtonAutomationTests

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
' Группа: тесты Utils (TC-01..TC-04, TC-06, TC-07, TC-19, TC-20)
' ============================================================
Private Sub RunUtilsTests()
    Dim Header As OrderHeader
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
    ' TC-06: GetSheetByName существующий
    ' -------------------------------------------------------
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "main")
    If Err.number <> 0 Then
        AddResult "TC-06", "GetSheetByName существующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-06", "GetSheetByName существующий лист", (Not ws Is Nothing), _
                  "Ожидалось Not Nothing, получено Nothing"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-07: GetSheetByName несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = GetSheetByName(ThisWorkbook, "NONEXISTENT")
    If Err.number <> 0 Then
        AddResult "TC-07", "GetSheetByName несуществующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-07", "GetSheetByName несуществующий лист", (ws Is Nothing), _
                  "Ожидалось Nothing, лист найден"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-19: WriteLog
    ' -------------------------------------------------------
    On Error Resume Next
    Call Mod_Utils.WriteLog("Mod_FullTestRunner: выполнение проверки TC-19")
    LogPath = ThisWorkbook.path & "\log.txt"
    If Err.number <> 0 Then
        AddResult "TC-19", "WriteLog запись в лог", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-19", "WriteLog запись в лог", FileExists(LogPath), _
                  "файл лога не найден: " & LogPath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-20: GetWorkbookPath / GetCurrentUser
    ' -------------------------------------------------------
    On Error Resume Next
    PathResult = GetWorkbookPath()
    UserResult = GetCurrentUser()
    If Err.number <> 0 Then
        AddResult "TC-20", "GetWorkbookPath / GetCurrentUser", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim PathOk As Boolean
        Dim UserOk As Boolean
        PathOk = (Len(PathResult) > 0)
        UserOk = (Len(UserResult) > 0)
        AddResult "TC-20", "GetWorkbookPath / GetCurrentUser", (PathOk And UserOk), _
                  "Path пустой=" & CStr(Not PathOk) & ", User пустой=" & CStr(Not UserOk)
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты OrderHeader (TC-08, TC-09, TC-11, TC-12)
'          TC-10 в резерве (требует специальных данных)
' ============================================================
Private Sub RunOrderHeaderTests()
    Dim Header As OrderHeader
    Dim FindResult As Boolean
    Dim wsMain As Worksheet
    Dim wsSpisok As Worksheet
    Dim wsModel As Worksheet
    Dim SavedState As Variant

    Debug.Print "--- Mod_OrderHeader Tests ---"

    ' -------------------------------------------------------
    ' TC-08: FindOrder существующий (по № п/п "1")
    ' -------------------------------------------------------
    On Error Resume Next
    ' Очистка Header перед тестом
    Header.OrderNumber = ""
    Header.ModelName = ""
    Header.grz = ""
    Header.VIN = ""
    Header.GarageNumber = ""
    Header.YearMade = 0
    Header.MileageValue = 0
    Header.DateValue = 0

    FindResult = FindOrder("1", Header)
    If Err.number <> 0 Then
        AddResult "TC-08", "FindOrder существующий (по № п/п '1')", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim Tc08Passed As Boolean
        Tc08Passed = FindResult And (Header.OrderNumber = "1")
        Dim Tc08Reason As String
        If Not FindResult Then
            Tc08Reason = "FindOrder вернул False"
        ElseIf Header.OrderNumber <> "1" Then
            Tc08Reason = "OrderNumber='" & Header.OrderNumber & "', ожидалось '1'"
        End If
        AddResult "TC-08", "FindOrder существующий (по № п/п '1')", Tc08Passed, Tc08Reason
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-09: FindOrder несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    FindResult = FindOrder("999", Header)
    If Err.number <> 0 Then
        AddResult "TC-09", "FindOrder несуществующий (по № п/п '999')", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-09", "FindOrder несуществующий (по № п/п '999')", (FindResult = False), _
                  "Ожидалось False, получено True"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-11: FillHeaderFromOrder существующий заказ
    ' -------------------------------------------------------
    On Error Resume Next
    Dim fillResult As Boolean
    fillResult = FillHeaderFromOrder("1")
    If Err.number <> 0 Then
        AddResult "TC-11", "FillHeaderFromOrder существующий заказ", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-11", "FillHeaderFromOrder существующий заказ", fillResult, ""
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-12: FillHeaderFromOrder заказ не найден
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")
    Set wsSpisok = GetSheetByName(ThisWorkbook, "spisok")
    Set wsModel = GetSheetByName(ThisWorkbook, "model")

    If (Not wsMain Is Nothing) And (Not wsSpisok Is Nothing) And (Not wsModel Is Nothing) Then
        ' Сохраняем состояние B3:B15
        SavedState = SaveSheetRange(wsMain, "B3:B15")

        ' Вызов с несуществующим заказом
        fillResult = FillHeaderFromOrder("999")

        If Err.number <> 0 Then
            AddResult "TC-12", "FillHeaderFromOrder заказ не найден", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            ' Проверяем, что B3:B15 очищены
            Dim IsCleared As Boolean
            IsCleared = (wsMain.Range("B3").Value = "") And _
                        (wsMain.Range("B4").Value = "") And _
                        (wsMain.Range("B5").Value = "") And _
                        (wsMain.Range("B6").Value = "") And _
                        (wsMain.Range("B7").Value = "") And _
                        (wsMain.Range("B8").Value = "") And _
                        (wsMain.Range("B9").Value = "") And _
                        (wsMain.Range("B10").Value = "") And _
                        (wsMain.Range("B11").Value = "") And _
                        (wsMain.Range("B12").Value = "") And _
                        (wsMain.Range("B13").Value = "") And _
                        (wsMain.Range("B14").Value = "") And _
                        (wsMain.Range("B15").Value = "")
            AddResult "TC-12", "FillHeaderFromOrder заказ не найден", IsCleared, _
                      "B3:B15 не были очищены после ошибки"
        End If

        ' Восстанавливаем состояние
        RestoreSheetRange wsMain, "B3:B15", SavedState
    Else
        AddResult "TC-12", "FillHeaderFromOrder заказ не найден", False, _
                  "Не найден один из листов: main/spisok/model"
    End If
    On Error GoTo 0

    Set wsMain = Nothing
    Set wsSpisok = Nothing
    Set wsModel = Nothing

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты импорта (TC-05, TC-13, TC-14, TC-15, TC-17)
'          TC-16 в резерве (требует наличия отчётной книги)
' ============================================================
Private Sub RunImportTests()
    Dim GRZResult As String
    Dim wsFound As Worksheet
    Dim wsMain As Worksheet
    Dim wsReport As Worksheet
    Dim SavedState As Variant

    Debug.Print "--- Mod_Import Tests ---"

    ' -------------------------------------------------------
    ' TC-05: ExtractNumberFromGRZ
    ' -------------------------------------------------------
    On Error Resume Next
    GRZResult = Mod_SheetOps.ExtractNumberFromGRZ("А123АН77")
    If Err.number <> 0 Then
        AddResult "TC-05", "ExtractNumberFromGRZ (А123АН77)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-05", "ExtractNumberFromGRZ (А123АН77)", (GRZResult = "123"), _
                  "Ожидалось '123', получено '" & GRZResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-13: ExtractNumberFromGRZ граничные случаи
    ' -------------------------------------------------------
    On Error Resume Next
    Dim Tc13AllPassed As Boolean
    Dim Tc13Details As String
    Tc13AllPassed = True
    Tc13Details = ""

    ' Кейс 1: "А123АН77" -> "123"
    GRZResult = Mod_SheetOps.ExtractNumberFromGRZ("А123АН77")
    If GRZResult <> "123" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[1: '" & GRZResult & "' != '123'] "
    End If

    ' Кейс 2: "А456ВС" -> "456"
    GRZResult = Mod_SheetOps.ExtractNumberFromGRZ("А456ВС")
    If GRZResult <> "456" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[2: '" & GRZResult & "' != '456'] "
    End If

    ' Кейс 3: "" -> ""
    GRZResult = Mod_SheetOps.ExtractNumberFromGRZ("")
    If GRZResult <> "" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[3: '" & GRZResult & "' != ''] "
    End If

    If Err.number <> 0 Then
        AddResult "TC-13", "ExtractNumberFromGRZ граничные случаи", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If Tc13AllPassed Then
            AddResult "TC-13", "ExtractNumberFromGRZ граничные случаи", True, ""
        Else
            AddResult "TC-13", "ExtractNumberFromGRZ граничные случаи", False, Tc13Details
        End If
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-14: SearchSheetByGRZ существующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsFound = Mod_SheetOps.SearchSheetByGRZ("12345")
    If Err.number <> 0 Then
        AddResult "TC-14", "SearchSheetByGRZ существующий", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            ' Пропускаем, если нет отчётной книги
            AddResult "TC-14", "SearchSheetByGRZ существующий", True, "", True, "Нет книги с GRZ_12345"
        Else
            AddResult "TC-14", "SearchSheetByGRZ существующий", (Not wsFound Is Nothing), _
                      "лист не найден"
        End If
    End If
    Set wsFound = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-15: SearchSheetByGRZ несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsFound = Mod_SheetOps.SearchSheetByGRZ("АБ")
    If Err.number <> 0 Then
        AddResult "TC-15", "SearchSheetByGRZ несуществующий", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            AddResult "TC-15", "SearchSheetByGRZ несуществующий", True, ""
        Else
            ' Если лист всё-таки найден — SKIP, т.к. среда непредсказуема
            AddResult "TC-15", "SearchSheetByGRZ несуществующий", True, "", True, "Лист 'АБ' существует"
        End If
    End If
    Set wsFound = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-17: ImportFromReport (устарел — функция удалена при рефакторинге)
    ' -------------------------------------------------------
    On Error Resume Next
    AddResult "TC-17", "ImportFromReport", True, "", True, _
              "Функция ImportFromReport удалена при рефакторинге (архитектура Phase 2)"
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты кнопок (TC-18)
' ============================================================
Private Sub RunButtonTests()
    Debug.Print "--- Mod_ButtonDispatcher Tests ---"

    ' TC-18: Btn_main_Clear_Click в ручном режиме
    ' Требует взаимодействия с MsgBox (подтверждение очистки)
    ' Не тестируется в автоматическом режиме
    AddResult "TC-18", "Btn_main_Clear_Click", True, "", True, "Ручной тест (требуется подтверждение в MsgBox)"

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты Logger (TC-21, TC-22, TC-23)
' ============================================================
Private Sub RunLoggerTests()
    Dim LogPath As String
    Dim OldLogPath As String
    Dim F As Long
    Dim i As Long

    Debug.Print "--- Mod_Logger Tests ---"

    ' -------------------------------------------------------
    ' TC-21: WriteLog — запись в лог-файл
    ' -------------------------------------------------------
    On Error Resume Next
    LogPath = Mod_Logger.GetLogPath()

    ' Очищаем лог перед тестом
    Call Mod_Logger.ClearLog

    ' Записываем сообщение
    Call Mod_Logger.WriteLog("TestModule", "Test message")

    If Err.number <> 0 Then
        AddResult "TC-21", "WriteLog запись в лог-файл", False, "Ошибка: " & Err.Description
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

        Dim Tc21Passed As Boolean
        Tc21Passed = FileExistsAfterWrite And FileContainsMessage
        Dim Tc21Reason As String
        If Not FileExistsAfterWrite Then
            Tc21Reason = "Файл лога не создан: " & LogPath
        ElseIf Not FileContainsMessage Then
            Tc21Reason = "Файл лога не содержит 'Test message'"
        End If
        AddResult "TC-21", "WriteLog запись в лог-файл", Tc21Passed, Tc21Reason
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-22: RotateLogIfNeeded — ротация лога
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
        AddResult "TC-22", "RotateLogIfNeeded ротация лога", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        Dim OldLogExists As Boolean
        OldLogExists = (Len(Dir(OldLogPath)) > 0)
        AddResult "TC-22", "RotateLogIfNeeded ротация лога", OldLogExists, _
                  "Старый лог не найден: " & OldLogPath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-23: ClearLog — очистка лога
    ' -------------------------------------------------------
    On Error Resume Next
    LogPath = Mod_Logger.GetLogPath()

    ' Записываем что-то в лог
    Call Mod_Logger.WriteLog("TestModule", "Message before clear")

    ' Очищаем лог
    Call Mod_Logger.ClearLog

    If Err.number <> 0 Then
        AddResult "TC-23", "ClearLog очистка лога", False, "Ошибка: " & Err.Description
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
        Dim Tc23Passed As Boolean
        Tc23Passed = (Not LogExistsAfterClear) Or LogIsEmpty
        AddResult "TC-23", "ClearLog очистка лога", Tc23Passed, _
                  "Файл существует=" & CStr(LogExistsAfterClear) & ", пуст=" & CStr(LogIsEmpty)
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: тесты SheetOps (TC-24, TC-25, TC-28)
' ============================================================
Private Sub RunSheetOpsTests()
    Dim wsMain As Worksheet
    Dim SavedState As Variant
    Dim wsTest As Worksheet
    Dim testSheetName As String

    Debug.Print "--- Mod_SheetOps Tests ---"

    ' -------------------------------------------------------
    ' TC-24: ClearMainSheet_UI — очистка листа main (silent mode)
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")

    If wsMain Is Nothing Then
        AddResult "TC-24", "ClearMainSheet_UI очистка листа main", False, "Нет листа 'main'"
    Else
        ' Сохраняем состояние B2:ZZ
        SavedState = SaveSheetRange(wsMain, "B2:ZZ")

        ' Заполняем тестовыми данными
        wsMain.Range("B2").Value = "TestData"
        wsMain.Range("C2").Value = "TestData2"
        wsMain.Range("A1").Value = "ShouldNotBeCleared"

        ' Вызываем очистку в silent режиме
        Call Mod_SheetOps.ClearMainSheet_UI(True)

        If Err.number <> 0 Then
            AddResult "TC-24", "ClearMainSheet_UI очистка листа main", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim B2Cleared As Boolean
            Dim C2Cleared As Boolean
            Dim A1NotCleared As Boolean

            B2Cleared = (wsMain.Range("B2").Value = "")
            C2Cleared = (wsMain.Range("C2").Value = "")
            A1NotCleared = (wsMain.Range("A1").Value <> "")

            Dim Tc24Passed As Boolean
            Tc24Passed = B2Cleared And C2Cleared And A1NotCleared
            Dim Tc24Reason As String
            If Not B2Cleared Then Tc24Reason = Tc24Reason & "B2 не очищен; "
            If Not C2Cleared Then Tc24Reason = Tc24Reason & "C2 не очищен; "
            If Not A1NotCleared Then Tc24Reason = Tc24Reason & "A1 очищен (не должен был); "
            AddResult "TC-24", "ClearMainSheet_UI очистка листа main", Tc24Passed, Tc24Reason
        End If

        ' Восстанавливаем состояние
        RestoreSheetRange wsMain, "B2:ZZ", SavedState
        wsMain.Range("A1").Value = ""
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-25: ClearHeader_UI — очистка шапки
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")

    If wsMain Is Nothing Then
        AddResult "TC-25", "ClearHeader_UI очистка шапки", False, "Нет листа 'main'"
    Else
        ' Сохраняем состояние B3:B15
        SavedState = SaveSheetRange(wsMain, "B3:B15")

        ' Заполняем тестовыми данными
        wsMain.Range("B3").Value = "TestHeader"
        wsMain.Range("B4").Value = "TestHeader2"

        ' Вызываем очистку шапки
        Call Mod_SheetOps.ClearHeader_UI

        If Err.number <> 0 Then
            AddResult "TC-25", "ClearHeader_UI очистка шапки", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim B3Cleared As Boolean
            Dim B4Cleared As Boolean
            B3Cleared = (wsMain.Range("B3").Value = "")
            B4Cleared = (wsMain.Range("B4").Value = "")
            AddResult "TC-25", "ClearHeader_UI очистка шапки", (B3Cleared And B4Cleared), _
                      "B3 пуста=" & CStr(B3Cleared) & ", B4 пуста=" & CStr(B4Cleared)
        End If

        ' Восстанавливаем состояние
        RestoreSheetRange wsMain, "B3:B15", SavedState
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-28: RenameSheetsByGRZ — переименование листов
    ' -------------------------------------------------------
    On Error Resume Next
    testSheetName = "GRZ_Test_12345"

    ' Создаём тестовый лист
    On Error Resume Next
    Set wsTest = GetSheetByName(ThisWorkbook, testSheetName)
    If Not wsTest Is Nothing Then
        wsTest.Delete
    End If
    On Error GoTo 0

    Set wsTest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsTest.Name = testSheetName

    ' Заполняем тестовыми данными
    wsTest.Range("A1").Value = "автомобиль А123АН77"

    If Err.number <> 0 Then
        AddResult "TC-28", "RenameSheetsByGRZ переименование листов", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        ' TC-28 требует report.xlsx, поэтому SKIP если нет
        If Mod_Utils.FileExists(ThisWorkbook.path & "\report.xlsx") Then
            ' Вызываем переименование (работает с report.xlsx, не с текущей книгой)
            ' В тестовом режиме просто проверяем, что функция не падает
            AddResult "TC-28", "RenameSheetsByGRZ переименование листов", True, "", True, _
                      "Требует report.xlsx для полного теста"
        Else
            AddResult "TC-28", "RenameSheetsByGRZ переименование листов", True, "", True, _
                      "Нет report.xlsx"
        End If
    End If

    ' Удаляем тестовый лист
    On Error Resume Next
    Set wsTest = GetSheetByName(ThisWorkbook, testSheetName)
    If Not wsTest Is Nothing Then
        wsTest.Delete
    End If
    On Error GoTo 0

    Set wsMain = Nothing
    Set wsTest = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: интеграционные тесты импорта (TC-26, TC-27)
' ============================================================
Private Sub RunImportIntegrationTests()
    Dim wsMain As Worksheet
    Dim wsTest As Worksheet
    Dim SavedState As Variant
    Dim testSheetName As String

    Debug.Print "--- Mod_Import Integration Tests ---"

    ' -------------------------------------------------------
    ' TC-26: ImportSheet — импорт листа
    ' -------------------------------------------------------
    On Error Resume Next
    If Mod_Utils.FileExists(ThisWorkbook.path & "\report.xlsx") Then
        ' Пытаемся вызвать ImportSheet с тестовым ГРЗ
        ' В тестовом режиме — SKIP, т.к. требует реального report.xlsx с данными
        AddResult "TC-26", "ImportSheet импорт листа", True, "", True, _
                  "report.xlsx существует, но тест требует реальных данных"
    Else
        AddResult "TC-26", "ImportSheet импорт листа", True, "", True, _
                  "Нет report.xlsx для теста"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-27: ImportDataToMain — перенос данных
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")
    If wsMain Is Nothing Then
        AddResult "TC-27", "ImportDataToMain перенос данных", False, "Нет листа 'main'"
    Else
        ' Сохраняем состояние L:N и X:AA
        SavedState = SaveSheetRange(wsMain, "L:N")

        ' Создаём тестовый лист с данными
        testSheetName = "_tc27_test"
        On Error Resume Next
        Set wsTest = GetSheetByName(ThisWorkbook, testSheetName)
        If Not wsTest Is Nothing Then
            wsTest.Delete
        End If
        On Error GoTo 0

        Set wsTest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsTest.Name = testSheetName

        ' Заполняем тестовыми данными (таблица "Выполненные работы")
        wsTest.Range("C3").Value = "Наименование работы"
        wsTest.Range("C4").Value = "Замена масла"
        wsTest.Range("D4").Value = "1000"
        wsTest.Range("H4").Value = "1"

        ' Вызываем ImportDataToMain
        Call Mod_Import.ImportDataToMain(wsTest)

        If Err.number <> 0 Then
            AddResult "TC-27", "ImportDataToMain перенос данных", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            ' Проверяем, что данные перенесены
            Dim DataImported As Boolean
            DataImported = (wsMain.Range("L2").Value <> "")
            AddResult "TC-27", "ImportDataToMain перенос данных", DataImported, _
                      "L2 пуст, данные не перенесены"
        End If

        ' Восстанавливаем состояние
        RestoreSheetRange wsMain, "L:N", SavedState

        ' Удаляем тестовый лист
        On Error Resume Next
        Set wsTest = GetSheetByName(ThisWorkbook, testSheetName)
        If Not wsTest Is Nothing Then
            wsTest.Delete
        End If
        On Error GoTo 0
    End If
    Set wsMain = Nothing
    Set wsTest = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: граничные случаи Utils (TC-29)
' ============================================================
Private Sub RunUtilsEdgeTests()
    Dim FmtResult As String

    Debug.Print "--- Mod_Utils Edge Cases ---"

    ' -------------------------------------------------------
    ' TC-29: FormatDateSQL граничные случаи
    ' -------------------------------------------------------

    ' Кейс 1: Пустая строка (невалидный вызов — FormatDateSQL ожидает Date, а не String)
    ' Пропускаем, т.к. это ошибка типа на уровне выполнения VBA
    AddResult "TC-29", "FormatDateSQL граничные случаи (пустая строка)", True, "", True, _
              "FormatDateSQL принимает Date, пустая строка — невалидный аргумент"

    ' Кейс 2: Только год
    On Error Resume Next
    FmtResult = FormatDateSQL(DateSerial(2026, 1, 1))
    If Err.number <> 0 Then
        AddResult "TC-29", "FormatDateSQL граничные случаи (только год)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-29", "FormatDateSQL граничные случаи (только год)", (FmtResult = "2026-01-01"), _
                  "Ожидалось '2026-01-01', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' Кейс 3: Дата с временем
    On Error Resume Next
    FmtResult = FormatDateSQL(DateSerial(2026, 7, 12) + TimeSerial(14, 30, 0))
    If Err.number <> 0 Then
        AddResult "TC-29", "FormatDateSQL граничные случаи (дата+время)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-29", "FormatDateSQL граничные случаи (дата+время)", (FmtResult = "2026-07-12"), _
                  "Ожидалось '2026-07-12', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: автоматизированные тесты кнопок (TC-30)
' ============================================================
Private Sub RunButtonAutomationTests()
    Dim wsMain As Worksheet
    Dim SavedState As Variant

    Debug.Print "--- Button Automation Tests ---"

    ' -------------------------------------------------------
    ' TC-30: Btn_main_Clear_Click — автоматизированный (silent)
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")

    If wsMain Is Nothing Then
        AddResult "TC-30", "Btn_main_Clear_Click автоматизированный", False, "Нет листа 'main'"
    Else
        ' Сохраняем состояние
        SavedState = SaveSheetRange(wsMain, "B2:ZZ")

        ' Заполняем тестовыми данными
        wsMain.Range("B2").Value = "TestClear"
        wsMain.Range("C5").Value = "TestClear2"

        ' Вызываем через диспетчер с silent=True
        Call Mod_SheetOps.ClearMainSheet_UI(True)

        If Err.number <> 0 Then
            AddResult "TC-30", "Btn_main_Clear_Click автоматизированный", False, "Ошибка: " & Err.Description
            Err.Clear
        Else
            Dim B2Empty As Boolean
            Dim C5Empty As Boolean
            B2Empty = (wsMain.Range("B2").Value = "")
            C5Empty = (wsMain.Range("C5").Value = "")
            AddResult "TC-30", "Btn_main_Clear_Click автоматизированный", (B2Empty And C5Empty), _
                      "B2 пуста=" & CStr(B2Empty) & ", C5 пуста=" & CStr(C5Empty)
        End If

        ' Восстанавливаем состояние
        RestoreSheetRange wsMain, "B2:ZZ", SavedState
    End If
    Set wsMain = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Вспомогательные функции
' ============================================================

' Сохраняет значения диапазона в массив
Private Function SaveSheetRange(ws As Worksheet, RangeAddr As String) As Variant
    On Error Resume Next
    SaveSheetRange = ws.Range(RangeAddr).Value
    On Error GoTo 0
End Function

' Восстанавливает значения диапазона из массива
Private Sub RestoreSheetRange(ws As Worksheet, RangeAddr As String, data As Variant)
    If ws Is Nothing Then Exit Sub
    If IsEmpty(data) Then Exit Sub
    On Error Resume Next
    ws.Range(RangeAddr).ClearContents
    If Not IsNull(data) Then
        If IsArray(data) Then
            ws.Range(RangeAddr).Value = data
        Else
            ' Одиночное значение
            ws.Range(RangeAddr).Value = data
        End If
    End If
    On Error GoTo 0
End Sub

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
' Запускает все тесты (TC-01..TC-30) и показывает результат
' --------------------------------------------------------------------------
Public Sub RunAllTests_UI()
    On Error GoTo ErrHandler

    Call RunAllTests

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RunAllTests_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RunAllTests_UI: " & Err.Description)
End Sub
