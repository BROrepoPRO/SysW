Attribute VB_Name = "Mod_MinimalTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_MinimalTestRunner
' Назначение: Минимальный набор тестов TC-01..TC-09
' ============================================================

' Главная процедура запуска тестов
Public Sub RunAllTests()
    Dim TotalTests As Long
    Dim PassedTests As Long
    Dim FailedTests As Long
    
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    
    Debug.Print "========================================="
    Debug.Print "Запуск тестов Mod_MinimalTestRunner"
    Debug.Print "========================================="
    
    ' TC-01: FileExists с существующим файлом
    TotalTests = TotalTests + 1
    If Test_FileExists_Existing() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-01: FileExists с существующим файлом"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-01: FileExists с существующим файлом"
    End If
    
    ' TC-02: FileExists с несуществующим файлом
    TotalTests = TotalTests + 1
    If Test_FileExists_NonExisting() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-02: FileExists с несуществующим файлом"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-02: FileExists с несуществующим файлом"
    End If
    
    ' TC-03: FormatDateSQL корректная дата
    TotalTests = TotalTests + 1
    If Test_FormatDateSQL_Valid() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-03: FormatDateSQL корректная дата"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-03: FormatDateSQL корректная дата"
    End If
    
    ' TC-04: FormatDateSQL с пустой датой
    TotalTests = TotalTests + 1
    If Test_FormatDateSQL_Empty() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-04: FormatDateSQL с пустой датой"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-04: FormatDateSQL с пустой датой"
    End If
    
    ' TC-05: ExtractNumberFromGRZ
    TotalTests = TotalTests + 1
    If Test_ExtractNumberFromGRZ() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-05: ExtractNumberFromGRZ"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-05: ExtractNumberFromGRZ"
    End If
    
    ' TC-06: GetSheetByName существующий лист
    TotalTests = TotalTests + 1
    If Test_GetSheetByName_Existing() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-06: GetSheetByName существующий лист"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-06: GetSheetByName существующий лист"
    End If
    
    ' TC-07: GetSheetByName несуществующий лист
    TotalTests = TotalTests + 1
    If Test_GetSheetByName_NonExisting() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-07: GetSheetByName несуществующий лист"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-07: GetSheetByName несуществующий лист"
    End If
    
    ' TC-08: FindOrder с существующим заказом
    TotalTests = TotalTests + 1
    If Test_FindOrder_Existing() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-08: FindOrder с существующим заказом"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-08: FindOrder с существующим заказом"
    End If
    
    ' TC-09: FindOrder с несуществующим заказом
    TotalTests = TotalTests + 1
    If Test_FindOrder_NonExisting() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] TC-09: FindOrder с несуществующим заказом"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] TC-09: FindOrder с несуществующим заказом"
    End If
    
    ' Итог
    Debug.Print "========================================="
    Debug.Print "Итого: " & TotalTests & " тестов, " & PassedTests & " пройдено, " & FailedTests & " падений"
    Debug.Print "========================================="
    
    MsgBox "Тестирование завершено." & vbCrLf & _
           "Всего: " & TotalTests & ", пройдено: " & PassedTests & ", падений: " & FailedTests, _
           vbInformation, "Результаты тестов"
End Sub

' TC-01: FileExists с существующим файлом
Private Function Test_FileExists_Existing() As Boolean
    On Error Resume Next
    Test_FileExists_Existing = FileExists("C:\Windows\notepad.exe")
    If Err.Number <> 0 Then Test_FileExists_Existing = False
    On Error GoTo 0
End Function

' TC-02: FileExists с несуществующим файлом
Private Function Test_FileExists_NonExisting() As Boolean
    On Error Resume Next
    Test_FileExists_NonExisting = Not FileExists("C:\Windows\ nonexistent_file_12345.tmp")
    If Err.Number <> 0 Then Test_FileExists_NonExisting = False
    On Error GoTo 0
End Function

' TC-03: FormatDateSQL корректная дата
Private Function Test_FormatDateSQL_Valid() As Boolean
    On Error Resume Next
    Dim Result As String
    Result = FormatDateSQL(DateSerial(2026, 7, 12))
    Test_FormatDateSQL_Valid = (Result = "2026-07-12")
    If Err.Number <> 0 Then Test_FormatDateSQL_Valid = False
    On Error GoTo 0
End Function

' TC-04: FormatDateSQL с пустой датой
Private Function Test_FormatDateSQL_Empty() As Boolean
    On Error Resume Next
    Dim Result As String
    Result = FormatDateSQL(0)
    Test_FormatDateSQL_Empty = (Result = "1899-12-30")
    If Err.Number <> 0 Then Test_FormatDateSQL_Empty = False
    On Error GoTo 0
End Function

' TC-05: ExtractNumberFromGRZ
Private Function Test_ExtractNumberFromGRZ() As Boolean
    On Error Resume Next
    Dim Result As String
    Result = ExtractNumberFromGRZ("А123ВВ77")
    Test_ExtractNumberFromGRZ = (Result = "12377")
    If Err.Number <> 0 Then Test_ExtractNumberFromGRZ = False
    On Error GoTo 0
End Function

' TC-06: GetSheetByName существующий лист
Private Function Test_GetSheetByName_Existing() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "main")
    Test_GetSheetByName_Existing = Not (ws Is Nothing)
    If Err.Number <> 0 Then Test_GetSheetByName_Existing = False
    On Error GoTo 0
End Function

' TC-07: GetSheetByName несуществующий лист
Private Function Test_GetSheetByName_NonExisting() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "NONEXISTENT_SHEET_12345")
    Test_GetSheetByName_NonExisting = (ws Is Nothing)
    If Err.Number <> 0 Then Test_GetSheetByName_NonExisting = False
    On Error GoTo 0
End Function

' TC-08: FindOrder с существующим заказом
Private Function Test_FindOrder_Existing() As Boolean
    On Error Resume Next
    Dim Header As OrderHeader
    Test_FindOrder_Existing = FindOrder("ЗН-001", Header)
    If Err.Number <> 0 Then Test_FindOrder_Existing = False
    On Error GoTo 0
End Function

' TC-09: FindOrder с несуществующим заказом
Private Function Test_FindOrder_NonExisting() As Boolean
    On Error Resume Next
    Dim Header As OrderHeader
    Test_FindOrder_NonExisting = Not FindOrder("NONEXISTENT_ORDER_999", Header)
    If Err.Number <> 0 Then Test_FindOrder_NonExisting = False
    On Error GoTo 0
End Function
