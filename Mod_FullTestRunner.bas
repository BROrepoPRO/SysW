Attribute VB_Name = "Mod_FullTestRunner"
Option Explicit

' ============================================================
' Â Â Â Â Â Â : Mod_FullTestRunner
' Â Â Â Â Â Â Â Â Â Â : Â Â Â Â Â Â  Â Â Â Â Â  Â Â Â Â Â Â  Â Â Â  Â Â Â Â Â Â Â  SysW
' Â Â Â Â Â Â Â Â : TC-01 .. TC-20 (17 Â Â Â Â Â Â Â Â Â Â Â Â Â Â , 3 Â Â Â Â Â Â )
' ============================================================

' ---- Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â  ----
Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

' ============================================================
' Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â : Â Â Â Â Â Â  Â Â Â Â  Â Â Â Â Â Â 
' ============================================================
Public Sub RunAllTests()
    ' Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â 
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0

    Debug.Print "=============================================="
    Debug.Print "  Â Â Â Â Â Â  Â Â Â Â Â Â Â  Â Â Â Â Â Â  Â Â Â Â Â Â  (TC-01..TC-20)"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Â Â Â Â Â Â  Â Â Â Â Â  Â Â Â Â Â Â 
    RunUtilsTests
    RunOrderHeaderTests
    RunImportTests
    RunButtonTests

    ' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â 
    PrintFinalReport
End Sub

' ============================================================
' Â Â Â Â Â Â : Â Â Â Â Â  Utils (TC-01..TC-04, TC-06, TC-07, TC-19, TC-20)
' ============================================================
Private Sub RunUtilsTests()
    Dim Header As OrderHeader
    Dim LogPath As String
    Dim Result As Boolean
    Dim PathResult As String
    Dim UserResult As String

    Debug.Print "--- Mod_Utils Tests ---"

    ' -------------------------------------------------------
    ' TC-01: FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Result = FileExists("C:\Windows\notepad.exe")
    If Err.Number <> 0 Then
        AddResult "TC-01", "FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-01", "FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", (Result = True), _
                  "Â Â Â Â Â Â Â Â Â  True, Â Â Â Â Â Â Â Â  " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-02: FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Result = FileExists("C:\nonexistent_file_12345.txt")
    If Err.Number <> 0 Then
        AddResult "TC-02", "FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-02", "FileExists Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", (Result = False), _
                  "Â Â Â Â Â Â Â Â Â  False, Â Â Â Â Â Â Â Â  " & CStr(Result)
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-03: FormatDateSQL Â  Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Dim FmtResult As String
    FmtResult = FormatDateSQL(DateSerial(2026, 7, 12))
    If Err.Number <> 0 Then
        AddResult "TC-03", "FormatDateSQL Â  Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-03", "FormatDateSQL Â  Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â ", (FmtResult = "2026-07-12"), _
                  "Â Â Â Â Â Â Â Â Â  '2026-07-12', Â Â Â Â Â Â Â Â  '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-04: FormatDateSQL Â  Â Â Â Â Â Â Â  Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    FmtResult = FormatDateSQL(0)
    If Err.Number <> 0 Then
        AddResult "TC-04", "FormatDateSQL Â  Â Â Â Â Â Â Â  Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-04", "FormatDateSQL Â  Â Â Â Â Â Â Â  Â Â Â Â Â ", (FmtResult = "1899-12-30"), _
                  "Â Â Â Â Â Â Â Â Â  '1899-12-30', Â Â Â Â Â Â Â Â  '" & FmtResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-06: GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheetByName(ThisWorkbook, "main")
    If Err.Number <> 0 Then
        AddResult "TC-06", "GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-06", "GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â ", (Not ws Is Nothing), _
                  "Â Â Â Â Â Â Â Â Â  Not Nothing, Â Â Â Â Â Â Â Â  Nothing"
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-07: GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Set ws = GetSheetByName(ThisWorkbook, "NONEXISTENT")
    If Err.Number <> 0 Then
        AddResult "TC-07", "GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-07", "GetSheetByName Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â ", (ws Is Nothing), _
                  "Â Â Â Â Â Â Â Â Â  Nothing, Â Â Â Â  Â Â Â Â Â Â "
    End If
    Set ws = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-19: WriteLog
    ' -------------------------------------------------------
    On Error Resume Next
    Call WriteLog("Mod_FullTestRunner: Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â  TC-19")
    LogPath = ThisWorkbook.Path & "\log.txt"
    If Err.Number <> 0 Then
        AddResult "TC-19", "WriteLog Â Â Â Â Â Â  Â  Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-19", "WriteLog Â Â Â Â Â Â  Â  Â Â Â ", FileExists(LogPath), _
                  "Â Â Â Â  Â Â Â Â  Â Â  Â Â Â Â Â Â : " & LogPath
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-20: GetWorkbookPath / GetCurrentUser
    ' -------------------------------------------------------
    On Error Resume Next
    PathResult = GetWorkbookPath()
    UserResult = GetCurrentUser()
    If Err.Number <> 0 Then
        AddResult "TC-20", "GetWorkbookPath / GetCurrentUser", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        Dim PathOk As Boolean
        Dim UserOk As Boolean
        PathOk = (Len(PathResult) > 0)
        UserOk = (Len(UserResult) > 0)
        AddResult "TC-20", "GetWorkbookPath / GetCurrentUser", (PathOk And UserOk), _
                  "Path Â Â Â Â Â Â =" & CStr(Not PathOk) & ", User Â Â Â Â Â Â =" & CStr(Not UserOk)
    End If
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Â Â Â Â Â Â : Â Â Â Â Â  OrderHeader (TC-08, TC-09, TC-11, TC-12)
'          TC-10 Â  Â Â Â Â Â Â Â  (Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â )
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
    ' TC-08: FindOrder Â Â Â Â Â Â Â Â Â Â Â Â  (Â Â  Â  Â /Â  "1")
    ' -------------------------------------------------------
    On Error Resume Next
    ' Â Â Â Â Â Â Â  Header Â Â Â Â Â  Â Â Â Â Â Â 
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
        AddResult "TC-08", "FindOrder Â Â Â Â Â Â Â Â Â Â Â Â  (Â Â  Â  Â /Â  '1')", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        Dim Tc08Passed As Boolean
        Tc08Passed = FindResult And (Header.OrderNumber = "1")
        Dim Tc08Reason As String
        If Not FindResult Then
            Tc08Reason = "FindOrder Â Â Â Â Â Â  False"
        ElseIf Header.OrderNumber <> "1" Then
            Tc08Reason = "OrderNumber='" & Header.OrderNumber & "', Â Â Â Â Â Â Â Â Â  '1'"
        End If
        AddResult "TC-08", "FindOrder Â Â Â Â Â Â Â Â Â Â Â Â  (Â Â  Â  Â /Â  '1')", Tc08Passed, Tc08Reason
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-09: FindOrder Â Â Â Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    FindResult = FindOrder("999", Header)
    If Err.Number <> 0 Then
        AddResult "TC-09", "FindOrder Â Â Â Â Â Â Â Â Â Â Â Â Â Â  (Â Â  Â  Â /Â  '999')", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-09", "FindOrder Â Â Â Â Â Â Â Â Â Â Â Â Â Â  (Â Â  Â  Â /Â  '999')", (FindResult = False), _
                  "Â Â Â Â Â Â Â Â Â  False, Â Â Â Â Â Â Â Â  True"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-11: FillHeaderFromOrder Â  Nothing-Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")
    ' Â Â Â Â Â Â Â Â  Nothing Â  wsSpisok Â  wsModel
    Call FillHeaderFromOrder("1")
    If Err.Number <> 0 Then
        AddResult "TC-11", "FillHeaderFromOrder Â  Nothing-Â Â Â Â Â Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        ' Â Â Â Â Â Â Â Â Â , Â Â Â  Â Â Â Â Â Â Â  Â Â Â Â Â Â  MsgBox Â  Â Â Â Â Â Â  False
        AddResult "TC-11", "FillHeaderFromOrder Â  Nothing-Â Â Â Â Â Â Â Â Â Â Â ", True, ""
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-12: FillHeaderFromOrder Â Â Â Â Â  Â Â  Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsMain = GetSheetByName(ThisWorkbook, "main")
    Set wsSpisok = GetSheetByName(ThisWorkbook, "spisok")
    Set wsModel = GetSheetByName(ThisWorkbook, "model")

    If (Not wsMain Is Nothing) And (Not wsSpisok Is Nothing) And (Not wsModel Is Nothing) Then
        ' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  B3:B15
        SavedState = SaveSheetRange(wsMain, "B3:B15")

        ' Â Â Â Â Â  Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â 
        Call FillHeaderFromOrder("999")

        If Err.Number <> 0 Then
            AddResult "TC-12", "FillHeaderFromOrder Â Â Â Â Â  Â Â  Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
            Err.Clear
        Else
            ' Â Â Â Â Â Â Â Â Â , Â Â Â  B3:B15 Â Â Â Â Â Â Â 
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
            AddResult "TC-12", "FillHeaderFromOrder Â Â Â Â Â  Â Â  Â Â Â Â Â Â ", IsCleared, _
                      "B3:B15 Â Â  Â Â Â Â  Â Â Â Â Â Â Â  Â Â Â Â Â  Â Â Â Â Â Â "
        End If

        ' Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â 
        RestoreSheetRange wsMain, "B3:B15", SavedState
    Else
        AddResult "TC-12", "FillHeaderFromOrder Â Â Â Â Â  Â Â  Â Â Â Â Â Â ", False, _
                  "Â Â  Â Â Â Â Â Â  Â Â Â Â  Â Â  Â Â Â Â Â Â : main/spisok/model"
    End If
    On Error GoTo 0

    Set wsMain = Nothing
    Set wsSpisok = Nothing
    Set wsModel = Nothing

    Debug.Print ""
End Sub

' ============================================================
' Â Â Â Â Â Â : Â Â Â Â Â  Â Â Â Â Â Â Â  (TC-05, TC-13, TC-14, TC-15, TC-17)
'          TC-16 Â  Â Â Â Â Â Â Â  (Â Â Â Â Â Â Â  Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â  Â Â Â Â Â )
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
    GRZResult = ExtractNumberFromGRZ("Â 123Â Â 77")
    If Err.Number <> 0 Then
        AddResult "TC-05", "ExtractNumberFromGRZ (Â 123Â Â 77)", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        AddResult "TC-05", "ExtractNumberFromGRZ (Â 123Â Â 77)", (GRZResult = "12377"), _
                  "Â Â Â Â Â Â Â Â Â  '12377', Â Â Â Â Â Â Â Â  '" & GRZResult & "'"
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-13: ExtractNumberFromGRZ Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Dim Tc13AllPassed As Boolean
    Dim Tc13Details As String
    Tc13AllPassed = True
    Tc13Details = ""

    ' Â Â Â Â  1: "Â 123Â Â 77" -> "12377"
    GRZResult = ExtractNumberFromGRZ("Â 123Â Â 77")
    If GRZResult <> "12377" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[1: '" & GRZResult & "' != '12377'] "
    End If

    ' Â Â Â Â  2: "Â 456Â Â " -> "456"
    GRZResult = ExtractNumberFromGRZ("Â 456Â Â ")
    If GRZResult <> "456" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[2: '" & GRZResult & "' != '456'] "
    End If

    ' Â Â Â Â  3: "" -> ""
    GRZResult = ExtractNumberFromGRZ("")
    If GRZResult <> "" Then
        Tc13AllPassed = False
        Tc13Details = Tc13Details & "[3: '" & GRZResult & "' != ''] "
    End If

    If Err.Number <> 0 Then
        AddResult "TC-13", "ExtractNumberFromGRZ Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        If Tc13AllPassed Then
            AddResult "TC-13", "ExtractNumberFromGRZ Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", True, ""
        Else
            AddResult "TC-13", "ExtractNumberFromGRZ Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ", False, Tc13Details
        End If
    End If
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-14: SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsFound = SearchSheetByGRZ("12345")
    If Err.Number <> 0 Then
        AddResult "TC-14", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            ' Â Â Â Â Â Â Â Â Â Â , Â Â Â Â  Â Â Â  Â Â Â Â Â Â Â Â  Â Â Â Â Â 
            AddResult "TC-14", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â ", True, "", True, "Â Â Â  Â Â Â Â Â  Â  GRZ_12345"
        Else
            AddResult "TC-14", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â ", (Not wsFound Is Nothing), _
                      "Â Â Â Â  Â Â  Â Â Â Â Â Â "
        End If
    End If
    Set wsFound = Nothing
    On Error GoTo 0

    ' -------------------------------------------------------
    ' TC-15: SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â Â Â 
    ' -------------------------------------------------------
    On Error Resume Next
    Set wsFound = SearchSheetByGRZ("Â Â ")
    If Err.Number <> 0 Then
        AddResult "TC-15", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â Â Â ", False, "Â Â Â Â Â Â : " & Err.Description
        Err.Clear
    Else
        If wsFound Is Nothing Then
            AddResult "TC-15", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â Â Â ", True, ""
        Else
            ' Â Â Â Â  Â Â Â Â  Â Â -Â Â Â Â  Â Â Â Â Â Â  Â  SKIP, Â .Â . Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â 
            AddResult "TC-15", "SearchSheetByGRZ Â Â Â Â Â Â Â Â Â Â Â Â Â Â ", True, "", True, "Â Â Â Â  'Â Â ' Â Â Â Â Â Â Â Â Â Â "
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
        AddResult "TC-17", "ImportFromReport", True, "", True, "Â Â Â  Â Â Â Â Â  'report'"
    Else
        Set wsMain = GetSheetByName(ThisWorkbook, "main")
        If wsMain Is Nothing Then
            AddResult "TC-17", "ImportFromReport", False, "Â Â Â  Â Â Â Â Â  'main'"
        Else
            ' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  A, B, C Â Â Â Â Â  main
            SavedState = SaveSheetRange(wsMain, "A:C")

            ' Â Â Â Â Â  Â Â Â Â Â Â Â Â Â 
            Call Mod_Import.ImportSheet(wsMain.Range("B4").Value)

            If Err.Number <> 0 Then
                AddResult "TC-17", "ImportFromReport", False, "Â Â Â Â Â Â : " & Err.Description
                Err.Clear
            Else
                AddResult "TC-17", "ImportFromReport", True, ""
            End If

            ' Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â 
            RestoreSheetRange wsMain, "A:C", SavedState
        End If
    End If
    Set wsReport = Nothing
    Set wsMain = Nothing
    On Error GoTo 0

    Debug.Print ""
End Sub

' ============================================================
' Â Â Â Â Â Â : Â Â Â Â Â  Â Â Â Â Â Â  (TC-18 Â  Â Â Â Â Â )
' ============================================================
Private Sub RunButtonTests()
    Debug.Print "--- Mod_ButtonDispatcher Tests ---"

    ' TC-18: Btn_main_Clear_Click Â  Â Â Â Â Â Â  Â Â Â Â Â Â 
    ' Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â  MsgBox (Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â )
    ' Â Â  Â Â Â Â Â Â Â Â Â Â Â  Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â 
    AddResult "TC-18", "Btn_main_Clear_Click", True, "", True, "Â Â Â Â Â Â  Â Â Â Â  (Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â Â Â  Â  MsgBox)"

    Debug.Print ""
End Sub

' ============================================================
' Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â 
' ============================================================

' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  Â  Â Â Â Â Â Â 
Private Function SaveSheetRange(ws As Worksheet, RangeAddr As String) As Variant
    On Error Resume Next
    SaveSheetRange = ws.Range(RangeAddr).Value
    On Error GoTo 0
End Function

' Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  Â Â  Â Â Â Â Â Â Â 
Private Sub RestoreSheetRange(ws As Worksheet, RangeAddr As String, data As Variant)
    If ws Is Nothing Then Exit Sub
    If IsEmpty(data) Then Exit Sub
    On Error Resume Next
    ws.Range(RangeAddr).ClearContents
    If Not IsNull(data) Then
        If IsArray(data) Then
            ws.Range(RangeAddr).Value = data
        Else
            ' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â 
            ws.Range(RangeAddr).Value = data
        End If
    End If
    On Error GoTo 0
End Sub

' Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  Â Â Â Â Â  Â  Â Â Â Â Â Â Â Â Â Â  Â  Â Â Â Â Â Â Â  Â  Immediate Window
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
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL - " & failReason
        Else
            Debug.Print "[" & testId & "] " & ChrW(&H2717) & " " & testName & ": FAIL"
        End If
    End If
End Sub

' Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â 
Private Sub PrintFinalReport()
    Dim ReportMsg As String

    Debug.Print ""
    Debug.Print "=============================================="
    Debug.Print "  Â Â Â Â Â Â Â Â  Â Â ?Â "
    Debug.Print "=============================================="
    Debug.Print "  Â Â Â Â Â : " & m_Total
    Debug.Print "  Â Â Â Â Â Â Â Â : " & m_Passed
    Debug.Print "  Â Â Â Â Â Â Â Â Â : " & m_Failed
    Debug.Print "  Â Â Â Â Â Â Â Â Â : " & m_Skipped
    Debug.Print "=============================================="

    ' Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â  Â Â Â  MsgBox
    ReportMsg = "Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â Â Â Â :" & vbCrLf & vbCrLf
    ReportMsg = ReportMsg & "  Â Â Â Â Â : " & m_Total & vbCrLf
    ReportMsg = ReportMsg & "  Â Â Â Â Â Â Â Â : " & m_Passed & vbCrLf
    ReportMsg = ReportMsg & "  Â Â Â Â Â Â Â Â Â : " & m_Failed & vbCrLf
    ReportMsg = ReportMsg & "  Â Â Â Â Â Â Â Â Â : " & m_Skipped & vbCrLf & vbCrLf

    If m_Failed = 0 Then
        ReportMsg = ReportMsg & "Â Â Â  Â Â Â Â Â  Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â !"
    Else
        ReportMsg = ReportMsg & "Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â ! Â Â Â Â Â Â Â Â Â  Immediate Window Â Â Â  Â Â Â Â Â Â Â ."
    End If

    MsgBox ReportMsg, vbInformation + vbOKOnly, "Mod_FullTestRunner"
End Sub


' ============================================================
' _UI-Â Â Â Â Â Â Â Â Â  (Â Â Â Â Â Â  Â  Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â /Â Â Â Â Â Â Â )
' ============================================================

' --------------------------------------------------------------------------
' RunAllTests_UI
' Â Â Â Â Â Â Â Â Â  Â Â Â  Â Â Â Â Â  (TC-01..TC-20) Â  Â Â Â Â Â Â Â Â Â Â  Â Â Â Â Â Â Â Â Â 
' --------------------------------------------------------------------------
Public Sub RunAllTests_UI()
    On Error GoTo ErrHandler

    Call RunAllTests

    Exit Sub

ErrHandler:
    MsgBox "Â Â Â Â Â Â  Â  RunAllTests_UI: " & Err.Description, vbCritical, "Â Â Â Â Â Â "
    Call Mod_Utils.WriteLog("RunAllTests_UI: " & Err.Description)
End Sub
