Attribute VB_Name = "Mod_OrderHeader"
Option Explicit

' ============================================================
' Модуль: Mod_OrderHeader
' Назначение: Заполнение заголовка заказа-наряда
' ============================================================

' DEPRECATED: use Mod_Constants.SPISOK_COL_*
Private Const SPISOK_COL_MODEL As Long = 2
Private Const SPISOK_COL_GRZ As Long = 3
Private Const SPISOK_COL_VIN As Long = 4
Private Const SPISOK_COL_GARAGE As Long = 5
Private Const SPISOK_COL_YEAR As Long = 6
Private Const SPISOK_COL_MILEAGE As Long = 7
Private Const SPISOK_COL_DATE As Long = 8
Private Const SPISOK_COL_NOTE As Long = 10

' DEPRECATED: use Mod_Constants.MODEL_COL_*
Private Const MODEL_COL_GROUP As Long = 2
Private Const MODEL_COL_PRICE As Long = 3
Private Const MODEL_COL_WORKS_ORIG As Long = 4
Private Const MODEL_COL_WORKS_MOD As Long = 5

' ============================================================
' Тип для хранения данных заказа-наряда
' Поля соответствуют столбцам листа "spisok":
' A=№ п/п, B=Модель, C=ГРЗ, D=VIN, E=гараж.№, F=год вып., G=пробег, H=дата
' ============================================================
Public Type OrderHeader
    OrderNumber As String    ' № п/п (колонка A)
    ModelName As String      ' Модель (колонка B)
    grz As String            ' ГРЗ/госномер (колонка C)
    VIN As String            ' VIN (колонка D)
    GarageNumber As String   ' Гаражный № (колонка E)
    YearMade As Integer      ' Год выпуска (колонка F)
    MileageValue As Long     ' Пробег (колонка G)
    DateValue As Date        ' Дата (колонка H)
End Type

' --------------------------------------------------------------------------
' FillHeaderFromOrder
' Заполняет B3:B15 на листе main данными из spisok и model по номеру заказа
' --------------------------------------------------------------------------
Public Function FillHeaderFromOrder(orderNum As Variant) As Boolean
    Dim wsSpisok As Worksheet
    Dim wsModel As Worksheet
    Dim wsMain As Worksheet
    Dim FoundRow As Range
    Dim ModelRow As Range
    Dim ModelCode As Variant

    Set wsMain = ThisWorkbook.Sheets("main")
    Set wsSpisok = ThisWorkbook.Sheets("spisok")
    Set wsModel = ThisWorkbook.Sheets("model")

    ' Очистка B3:B15
    wsMain.Range("B3:B15").ClearContents

    ' Проверка, что OrderNum — число
    If Not IsNumeric(orderNum) Then
        Call Mod_Logger.WriteLog("Mod_OrderHeader", "FillHeaderFromOrder: Номер заказа должен быть числом!")
        FillHeaderFromOrder = False
        Exit Function
    End If

    ' Поиск заказа в столбце A листа spisok (точное совпадение)
    Set FoundRow = wsSpisok.Columns(1).Find(What:=orderNum, LookAt:=xlWhole)
    If FoundRow Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_OrderHeader", "FillHeaderFromOrder: Заказ с номером " & orderNum & " не найден!")
        wsMain.Range("B3:B15").ClearContents
        FillHeaderFromOrder = False
        Exit Function
    End If

    ' B3 = "00" & значение B2 & "-20"
    wsMain.Cells(3, 2).Value = "00" & CStr(orderNum) & "-20"

    ' B4:B10 из столбцов B–H найденной строки spisok
    wsMain.Cells(4, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_MODEL).Value   ' Модель
    wsMain.Cells(5, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_GRZ).Value     ' ГРЗ
    wsMain.Cells(6, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_VIN).Value     ' VIN
    wsMain.Cells(7, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_GARAGE).Value  ' Гараж.№
    wsMain.Cells(8, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_YEAR).Value    ' Год вып.
    wsMain.Cells(9, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_MILEAGE).Value ' Пробег
    wsMain.Cells(10, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_DATE).Value   ' Дата

    ' Код модели из столбца I
    ModelCode = FoundRow.Cells(1, 9).Value

    ' Поиск кода модели в столбце A листа model (начиная со строки 3)
    If Not IsNull(ModelCode) And ModelCode <> "" Then
        Dim lastModelRow As Long
        lastModelRow = wsModel.Cells(wsModel.Rows.count, 1).End(xlUp).Row
        Set ModelRow = wsModel.Range("A3:A" & lastModelRow).Find(What:=ModelCode, LookAt:=xlWhole)
        If Not ModelRow Is Nothing And ModelRow.Row >= 3 Then
            wsMain.Cells(11, 2).Value = ModelRow.Cells(1, Mod_Constants.MODEL_COL_PRICE).Value      ' Цена н/ч
            wsMain.Cells(12, 2).Value = ModelRow.Cells(1, Mod_Constants.MODEL_COL_GROUP).Value      ' Группа
            wsMain.Cells(13, 2).Value = ModelRow.Cells(1, Mod_Constants.MODEL_COL_WORKS_ORIG).Value ' Работы исх
            wsMain.Cells(14, 2).Value = ModelRow.Cells(1, Mod_Constants.MODEL_COL_WORKS_MOD).Value  ' Работы мод
        End If
    End If

    ' B15 = примечание из столбца J spisok
    wsMain.Cells(15, 2).Value = FoundRow.Cells(1, Mod_Constants.SPISOK_COL_NOTE).Value

    FillHeaderFromOrder = True
End Function

' --------------------------------------------------------------------------
' FindOrder
' Поиск заказа по номеру (столбец A) на листе spisok
' Заполняет структуру OrderHeader полями A–H
' --------------------------------------------------------------------------
Public Function FindOrder(ByVal orderNum As String, ByRef Header As OrderHeader) As Boolean
    Dim ws As Worksheet
    Dim FoundCell As Range

    Set ws = GetSheetByName(ThisWorkbook, "spisok")
    If ws Is Nothing Then
        FindOrder = False
        Exit Function
    End If

    ' Поиск в столбце A (номер заказа)
    Set FoundCell = ws.Columns("A").Find(What:=orderNum, LookAt:=xlWhole, LookIn:=xlValues)

    If Not FoundCell Is Nothing Then
        ' Заполнение структуры из найденной строки листа spisok:
        ' A=№ заказа, B=Модель, C=ГРЗ, D=VIN, E=Гараж.№, F=Год вып., G=Пробег, H=Дата
        Header.OrderNumber = FoundCell.Value                    ' A: № заказа
        Header.ModelName = FoundCell.Offset(0, 1).Value         ' B: Модель
        Header.grz = FoundCell.Offset(0, 2).Value               ' C: ГРЗ
        Header.VIN = FoundCell.Offset(0, 3).Value               ' D: VIN
        Header.GarageNumber = FoundCell.Offset(0, 4).Value      ' E: Гараж. №
        Header.YearMade = Val(FoundCell.Offset(0, 5).Value)     ' F: Год вып.
        Header.MileageValue = Val(FoundCell.Offset(0, 6).Value) ' G: Пробег
        Header.DateValue = FoundCell.Offset(0, 7).Value         ' H: Дата
        FindOrder = True
    Else
        FindOrder = False
    End If
End Function

' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' FillHeaderFromOrder_UI
' Запрашивает номер заказа из B2, вызывает FillHeaderFromOrder,
' показывает результат пользователю
' --------------------------------------------------------------------------
Public Sub FillHeaderFromOrder_UI()
    On Error GoTo ErrHandler

    Dim orderNum As Variant
    orderNum = ThisWorkbook.Sheets("main").Range("B2").Value

    If IsEmpty(orderNum) Or orderNum = "" Then
        MsgBox "Введите номер заказа в ячейку B2!", vbExclamation, "Предупреждение"
        Exit Sub
    End If

    If Not FillHeaderFromOrder(orderNum) Then
        MsgBox "Не удалось заполнить шапку заказа.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    MsgBox "Шапка заказа заполнена.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в FillHeaderFromOrder_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_OrderHeader", "FillHeaderFromOrder_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' FindOrder_UI
' Запрашивает номер заказа через InputBox, вызывает FindOrder,
' форматирует и показывает результат
' --------------------------------------------------------------------------
Public Sub FindOrder_UI()
    On Error GoTo ErrHandler

    Dim orderNum As String
    Dim Header As OrderHeader
    Dim Result As Boolean
    Dim msg As String

    orderNum = InputBox("Введите номер заказа для поиска:", "Поиск заказа")

    If orderNum = "" Then
        Exit Sub
    End If

    Result = FindOrder(orderNum, Header)

    If Result Then
        msg = "Заказ найден:" & vbCrLf & vbCrLf
        msg = msg & "Номер: " & Header.OrderNumber & vbCrLf
        msg = msg & "Модель: " & Header.ModelName & vbCrLf
        msg = msg & "ГРЗ: " & Header.grz & vbCrLf
        msg = msg & "VIN: " & Header.VIN & vbCrLf
        msg = msg & "Гараж.№: " & Header.GarageNumber & vbCrLf
        msg = msg & "Год вып.: " & Header.YearMade & vbCrLf
        msg = msg & "Пробег: " & Header.MileageValue & vbCrLf
        msg = msg & "Дата: " & Header.DateValue
        MsgBox msg, vbInformation, "Результат поиска"
    Else
        MsgBox "Заказ с номером '" & orderNum & "' не найден.", vbExclamation, "Результат поиска"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в FindOrder_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("FindOrder_UI: " & Err.Description)
End Sub
