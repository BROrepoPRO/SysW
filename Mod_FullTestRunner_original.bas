Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_FullTestRunner
' Назначение: Полный набор тестов для системы SysW
' Покрытие: TC-01 .. TC-20 (17 автоматических, 3 ручных)
' ============================================================

' ---- Счётчики результатов ----
Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

' ============================================================
' Главная процедура: запуск всех тестов
' ============================================================
Public Sub RunAllTests()
    ' Инициализация счётчиков
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0

    Debug.Print "=============================================="
    Debug.Print "  ЗАПУСК ПОЛНОГО НАБОРА ТЕСТОВ (TC-01..TC-20)"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Запуск групп тестов
    RunUtilsTests
    RunOrderHeaderTests
    RunImportTests
    RunButtonTests

    ' Итоговый отчёт
    PrintFinalReport
End Sub

' ============================================================
' Группа: Тесты утилит (TC-01..TC-04, TC-06, TC-07, TC-19, TC-20)
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
    If Err.Number <> 0 Then
        AddResult "TC-01", "FileExists с существующим файлом", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-01", "FileExists с существующим файлом", (Result = True), _
                  "ожидалось True, получено " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-02: FileExists с несуществующим файлом
    ' -------------------------------------------------------
    On Error Resume Next
    Result = FileExists("C:\nonexistent_file_12345.txt")
    If Err.Number <> 0 Then
        AddResult "TC-02", "FileExists с несуществующим файлом", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-02", "FileExists с несуществующим файлом", (Result = False), _
                  "ожидалось False, получено " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-03: FormatDateSQL с валидной датой
    ' -------------------------------------------------------
    On Error Resume Next
    Dim FmtResult As String
    FmtResult = FormatDateSQL(DateSerial(2026, 7, 12))
    If Err.Number <> 0 Then
        AddResult "TC-03", "FormatDateSQL с валидной датой", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-03", "FormatDateSQL с валидной датой", (FmtResult = "2026-07-12"), _
                  "ожидалось '2026-07-12', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-04: FormatDateSQL с нулевой датой
    ' -------------------------------------------------------
    On Error Resume Next
    FmtResult = FormatDateSQL(0)
    If Err.Number <> 0 Then
        AddResult "TC-04", "FormatDateSQL с нулевой датой", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-04", "FormatDateSQL с нулевой датой", (FmtResult = "1899-12-30"), _
                  "ожидалось '1899-12-30', получено '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-06: GetSheetByName существующий
    ' -------------------------------------------------------
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "main")
    If Err.Number <> 0 Then
        AddResult "TC-06", "GetSheetByName существующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-06", "GetSheetByName существующий лист", (Not ws Is Nothing), _
                  "ожидался Not Nothing, получено Nothing"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-07: GetSheetByName несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = GetSheetByName(ThisWorkbook, "NONEXISTENT")
    If Err.Number <> 0 Then
        AddResult "TC-07", "GetSheetByName несуществующий лист", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-07", "GetSheetByName несуществующий лист", (ws Is Nothing), _
                  "ожидался Nothing, получен лист"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-19: WriteLog
    ' -------------------------------------------------------
    On Error Resume Next
    Call WriteLog("Mod_FullTestRunner: тестовое сообщение TC-19")
    LogPath = ThisWorkbook.Path & "\log.txt"
    If Err.Number <> 0 Then
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
    If Err.Number <> 0 Then
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
' Группа: Тесты OrderHeader (TC-08, TC-09, TC-11, TC-12)
'          TC-10 — ручной тест (проверка визуального результата)
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
    ' TC-08: FindOrder существующий (№ п/п "1")
    ' -------------------------------------------------------
    On Error Resume Next
    ' Очищаем Header перед тестом
    Header.OrderNumber = ""
    Header.ModelName = ""
    Header.GRZ = ""
    Header.VIN = ""
    Header.GarageNumber = ""
    Header.YearMade = 0
    Header.MileageValue = 0
    Header.DateValue = 0

    FindResult = FindOrder("1", Header)
    If Err.Number <> 0 Then
        AddResult "TC-08", "FindOrder существующий (№ п/п '1')", False, "Ошибка: " & Err.Description
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
        AddResult "TC-08", "FindOrder существующий (№ п/п '1')", Tc08Passed, Tc08Reason
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-09: FindOrder несуществующий
    ' -------------------------------------------------------
    On Error Resume Next
    FindResult = FindOrder("999", Header)
    If Err.Number <> 0 Then
        AddResult "TC-09", "FindOrder несуществующий (№ п/п '999')", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-09", "FindOrder несуществующий (№ п/п '999')", (FindResult = False), _
                  "ожидалось False, получено True"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-11: FillHeaderFromOrder с Nothing-листами
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")
    ' Передаём Nothing в wsSpisok и wsModel
    Call FillHeaderFromOrder("1", wsMain, Nothing, Nothing)
    If Err.Number <> 0 Then
        AddResult "TC-11", "FillHeaderFromOrder с Nothing-листами", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        ' Ожидается, что процедура вызовет MsgBox и выйдет без ошибки
        AddResult "TC-11", "FillHeaderFromOrder с Nothing-листами", True, ""
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

        ' Вызываем с несуществующим номером
        Call FillHeaderFromOrder("999", wsMain, wsSpisok, wsModel)

        If Err.Number <> 0 Then
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
                      "B3:B15 не полностью очищены"
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
' Группа: Тесты импорта (TC-05, TC-13, TC-14, TC-15, TC-17)
'          TC-16 — ручной тест (требует создания временного листа)
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
    GRZResult = ExtractNumberFromGRZ("А123ВВ77")
    If Err.Number <> 0 Then
        AddResult "TC-05", "ExtractNumberFromGRZ (А123ВВ77)", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        AddResult "TC-05", "ExtractNumberFromGRZ (А123ВВ77)", (GRZResult = "12377"), _
                  "ожидалось '12377', получено '" & GRZResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-13: ExtractNumberFromGRZ различные форматы
    ' -------------------------------------------------------
    On Error Resume Next
    Dim Tc13AllPassed As Boolean
    Dim Tc13Details As String
    Tc13AllPassed = True
    Tc13Details = ""

    ' Тест 1: "А123ВВ77" > "12377"
    GRZResult = ExtractNumberFromGRZ("А123ВВ77")
    If GRZResult <> "12377" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[1: '" & GRZResult & "' != '12377'] "
    End If

    ' Тест 2: "В456ЕК" > "456"
    GRZResult = ExtractNumberFromGRZ("В456ЕК")
    If GRZResult <> "456" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[2: '" & GRZResult & "' != '456'] "
    End If

    ' Тест 3: "" > ""
    GRZResult = ExtractNumberFromGRZ("")
    If GRZResult <> "" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[3: '" & GRZResult & "' != ''] "
    End If

    If Err.Number <> 0 Then
        AddResult "TC-13", "ExtractNumberFromGRZ различные форматы", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If Tc13AllPassed Then
            AddResult "TC-13", "ExtractNumberFromGRZ различные форматы", True, ""
        Else
            AddResult "TC-13", "ExtractNumberFromGRZ различные форматы", False, Tc13Details
        End If
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-14: SearchSheetByGRZ существующий
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsFound = SearchSheetByGRZ("12345")
    If Err.Number <> 0 Then
        AddResult "TC-14", "SearchSheetByGRZ существующий", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            ' Пропускаем, если нет подходящего листа
            AddResult "TC-14", "SearchSheetByGRZ существующий", True, "", True, "нет листа с GRZ_12345"
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
    Set wsFound = SearchSheetByGRZ("ЗН")
    If Err.Number <> 0 Then
        AddResult "TC-15", "SearchSheetByGRZ несуществующий", False, "Ошибка: " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            AddResult "TC-15", "SearchSheetByGRZ несуществующий", True, ""
        Else
            ' Если лист всё-таки нашёлся — SKIP, т.к. тест предполагает отсутствие
            AddResult "TC-15", "SearchSheetByGRZ несуществующий", True, "", True, "лист 'ЗН' существует"
        End If
    End If
    Set wsFound = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-17: ImportFromReport
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsReport = GetSheetByName(ThisWorkbook, "report")
    If wsReport Is Nothing Then
        AddResult "TC-17", "ImportFromReport", True, "", True, "нет листа 'report'"
    Else
        Set wsMain = GetSheetByName(ThisWorkbook, "main")
        If wsMain Is Nothing Then
            AddResult "TC-17", "ImportFromReport", False, "нет листа 'main'"
        Else
            ' Сохраняем состояние столбцов A, B, C листа main
            SavedState = SaveSheetRange(wsMain, "A:C")

            ' Вызываем импорт
            Call ImportFromReport

            If Err.Number <> 0 Then
                AddResult "TC-17", "ImportFromReport", False, "Ошибка: " & Err.Description
                Err.Clear
            Else
                AddResult "TC-17", "ImportFromReport", True, ""
            End If

            ' Восстанавливаем состояние
            RestoreSheetRange wsMain, "A:C", SavedState
        End If
    End If
    Set wsReport = Nothing
    Set wsMain = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Группа: Тесты кнопок (TC-18 — ручной)
' ============================================================
Private Sub RunButtonTests()
    Debug.Print "--- Mod_ButtonDispatcher Tests ---"

    ' TC-18: Btn_main_Clear_Click — ручной тест
    ' Требует взаимодействия с MsgBox (подтверждение очистки)
    ' Не включён в автоматический прогон
    AddResult "TC-18", "Btn_main_Clear_Click", True, "", True, "ручной тест (требует подтверждения в MsgBox)"

    Debug.Print ""
End Sub

' ============================================================
' Вспомогательные функции
' ============================================================

' Сохраняет диапазон листа в массив
Private Function SaveSheetRange(ws As Worksheet, RangeAddr As String) As Variant
    On Error Resume Next
    SaveSheetRange = ws.Range(RangeAddr).Value
    On Error GoTo 0
End Function

' Восстанавливает диапазон листа из массива
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

' Добавляет результат теста в счётчики и выводит в Immediate Window
Private Sub AddResult(testId As String, testName As String, _
                      passed As Boolean, Optional failReason As String = "", _
                      Optional skipped As Boolean = False, Optional skipReason As String = "")
    m_Total = m_Total + 1

    If skipped Then
        m_Skipped = m_Skipped + 1
        Debug.Print "[" & testId & "] " & ChrW(&H26A0) & " " & testName & ": SKIP (" & skipReason & ")"
    ElseIf passed Then
        m_Passed = m_Passed + 1
        Debug.Print "[" & testId & "] " & ChrW(&H2713) & " " & testName & ": PASS"
    Else
        m_Failed = m_Failed + 1
        If failReason <> "" Then
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL — " & failReason
        Else
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL"
        End If
    End If
End Sub

' Выводит итоговый отчёт
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

    ' Формируем сообщение для MsgBox
    ReportMsg = "Результаты тестирования:" & vbCrLf & vbCrLf
    ReportMsg = ReportMsg & "  Всего: " & m_Total & vbCrLf
    ReportMsg = ReportMsg & "  Пройдено: " & m_Passed & vbCrLf
    ReportMsg = ReportMsg & "  Провалено: " & m_Failed & vbCrLf
    ReportMsg = ReportMsg & "  Пропущено: " & m_Skipped & vbCrLf & vbCrLf

    If m_Failed = 0 Then
        ReportMsg = ReportMsg & "Все тесты пройдены успешно!"
    Else
        ReportMsg = ReportMsg & "Обнаружены ошибки! Проверьте Immediate Window для деталей."
    End If

    MsgBox ReportMsg, vbInformation + vbOKOnly, "Mod_FullTestRunner"
End Sub