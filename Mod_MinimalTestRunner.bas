Attribute VB_Name = "Mod_MinimalTestRunner"
Option Explicit

' ============================================================
' Модуль: Mod_MinimalTestRunner
' Назначение: минимальный набор тестов для проверки модулей
' ============================================================

' Главная процедура запуска тестов
Public Sub RunAllTests()
    Dim TotalTests As Long
    Dim PassedTests As Long
    Dim FailedTests As Long
    
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    
    ' Тест 1: проверка FileExists
    TotalTests = TotalTests + 1
    If Test_FileExists() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] Тест FileExists пройден"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] Тест FileExists не пройден"
    End If
    
    ' Тест 2: проверка FormatDateSQL
    TotalTests = TotalTests + 1
    If Test_FormatDateSQL() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] Тест FormatDateSQL пройден"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] Тест FormatDateSQL не пройден"
    End If
    
    ' Тест 3: проверка FindOrder
    TotalTests = TotalTests + 1
    If Test_FindOrder() Then
        PassedTests = PassedTests + 1
        Debug.Print "  [OK] Тест FindOrder пройден"
    Else
        FailedTests = FailedTests + 1
        Debug.Print "  [FAIL] Тест FindOrder не пройден"
    End If
    
    ' Итог
    Debug.Print "========================================="
    Debug.Print "ИТОГО: " & TotalTests & " тестов, " & PassedTests & " пройдено, " & FailedTests & " упало"
    Debug.Print "========================================="
    
    MsgBox "Тестирование завершено." & vbCrLf & _
           "Всего: " & TotalTests & ", Пройдено: " & PassedTests & ", Упало: " & FailedTests, _
           vbInformation, "Результаты тестов"
End Sub

' Тест: FileExists
Private Function Test_FileExists() As Boolean
    On Error Resume Next
    ' Проверяем, что функция существует и возвращает Boolean
    Dim Result As Boolean
    Result = FileExists("C:\Windows\notepad.exe")
    Test_FileExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Тест: FormatDateSQL
Private Function Test_FormatDateSQL() As Boolean
    On Error Resume Next
    Dim Result As String
    Result = FormatDateSQL(DateSerial(2026, 7, 12))
    Test_FormatDateSQL = (Result = "2026-07-12")
    On Error GoTo 0
End Function

' Тест: FindOrder
Private Function Test_FindOrder() As Boolean
    On Error Resume Next
    Dim Header As OrderHeader
    Test_FindOrder = FindOrder("ЗН-001", Header)
    On Error GoTo 0
End Function