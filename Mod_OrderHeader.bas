Attribute VB_Name = "Mod_OrderHeader"
Option Explicit

' ============================================================
' Модуль: Mod_OrderHeader
' Назначение: автоматическое заполнение шапки заказ-наряда
' ============================================================

' Константы для столбцов
Private Const COL_ORDER_NUM As String = "B"
Private Const COL_CLIENT As String = "C"
Private Const COL_CAR_MODEL As String = "D"
Private Const COL_CAR_PLATE As String = "E"
Private Const COL_MILEAGE As String = "F"
Private Const COL_DATE_IN As String = "G"
Private Const COL_DATE_OUT As String = "H"
Private Const COL_STATUS As String = "I"

' Обработчик изменения ячейки (вызывается из листа)
Public Sub OnWorksheetChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Intersect(Target, Range(COL_ORDER_NUM & ":" & COL_ORDER_NUM)) Is Nothing Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    
    Dim OrderNum As String
    OrderNum = Trim(Target.Value)
    
    If OrderNum = "" Then Exit Sub
    
    ' Поиск заказа в SQLite
    Dim Header As OrderHeader
    If FindOrder(OrderNum, Header) Then
        ' Заполняем шапку
        Application.EnableEvents = False
        Range(COL_CLIENT & Target.Row).Value = Header.ClientName
        Range(COL_CAR_MODEL & Target.Row).Value = Header.CarModel
        Range(COL_CAR_PLATE & Target.Row).Value = Header.CarPlate
        Range(COL_MILEAGE & Target.Row).Value = Header.Mileage
        Range(COL_DATE_IN & Target.Row).Value = Header.DateIn
        Range(COL_DATE_OUT & Target.Row).Value = Header.DateOut
        Range(COL_STATUS & Target.Row).Value = Header.Status
        Application.EnableEvents = True
    End If
End Sub

' Функция: поиск заказа в SQLite по номеру
Private Function FindOrder(ByVal OrderNum As String, ByRef Header As OrderHeader) As Boolean
    ' Заглушка — в реальности здесь будет SQL-запрос к SQLite
    ' Пока возвращаем тестовые данные
    If OrderNum = "ЗН-001" Then
        Header.OrderNumber = OrderNum
        Header.ClientName = "Иванов И.И."
        Header.CarModel = "Toyota Camry"
        Header.CarPlate = "А123ВС77"
        Header.Mileage = 45000
        Header.DateIn = DateSerial(2026, 7, 1)
        Header.DateOut = DateSerial(2026, 7, 5)
        Header.Status = "В работе"
        FindOrder = True
    Else
        FindOrder = False
    End If
End Function