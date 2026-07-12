Attribute VB_Name = "Mod_OrderHeader"
Option Explicit

' ============================================================
' Модуль: Mod_OrderHeader
' Назначение: Автоматическое заполнение полей заказа-наряда
' ============================================================

' Главная процедура: заполнение заголовка из найденного заказа
Public Sub FillHeaderFromOrder(ByVal OrderNum As String, _
                               ByVal wsMain As Worksheet, _
                               ByVal wsSpisok As Worksheet, _
                               ByVal wsModel As Worksheet)
    Dim FoundCell As Range
    Dim ModelCode As String
    Dim ModelFound As Range
    
    ' Проверка листов
    If wsSpisok Is Nothing Or wsModel Is Nothing Then
        MsgBox "Не найдены служебные листы (spisok, model).", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Поиск по .Text в колонке A листа spisok
    Set FoundCell = wsSpisok.Columns("A").Find(What:=OrderNum, LookAt:=xlWhole, LookIn:=xlValues)
    
    If Not FoundCell Is Nothing Then
        ' Заполняем B3:B15
        wsMain.Range("B3").Value = FoundCell.Value                                   ' Номер заказа
        wsMain.Range("B4").Value = FoundCell.Offset(0, 1).Value                      ' Клиент
        wsMain.Range("B5").Value = FoundCell.Offset(0, 2).Value                      ' Автомобиль
        wsMain.Range("B6").Value = FoundCell.Offset(0, 3).Value                      ' Госномер
        wsMain.Range("B7").Value = FoundCell.Offset(0, 4).Value                      ' Пробег
        wsMain.Range("B8").Value = FoundCell.Offset(0, 5).Value                      ' Дата заезда
        wsMain.Range("B9").Value = FoundCell.Offset(0, 6).Value                      ' Дата выезда
        wsMain.Range("B10").Value = FoundCell.Offset(0, 7).Value                     ' Статус
        
        ' Поиск модели по коду
        ModelCode = FoundCell.Offset(0, 8).Value                                     ' Код модели
        If ModelCode <> "" Then
            Set ModelFound = wsModel.Columns("A").Find(What:=ModelCode, LookAt:=xlWhole, LookIn:=xlValues)
            If Not ModelFound Is Nothing Then
                wsMain.Range("B11").Value = ModelFound.Value                         ' Модель
                wsMain.Range("B12").Value = ModelFound.Offset(0, 1).Value            ' Цвет
                wsMain.Range("B13").Value = ModelFound.Offset(0, 2).Value            ' Год выпуска
                wsMain.Range("B14").Value = ModelFound.Offset(0, 3).Value            ' VIN
            End If
        End If
        
        wsMain.Range("B15").Value = FoundCell.Offset(0, 9).Value                     ' Примечание
    Else
        ' Очищаем B3:B15
        wsMain.Range("B3:B15").ClearContents
    End If
End Sub

' Публичная функция для тестов: поиск заказа
Public Function FindOrder(ByVal OrderNum As String, ByRef Header As OrderHeader) As Boolean
    Dim ws As Worksheet
    Dim FoundCell As Range
    
    Set ws = GetSheetByName(ThisWorkbook, "spisok")
    If ws Is Nothing Then
        FindOrder = False
        Exit Function
    End If
    
    Set FoundCell = ws.Columns("A").Find(What:=OrderNum, LookAt:=xlWhole, LookIn:=xlValues)
    
    If Not FoundCell Is Nothing Then
        Header.OrderNumber = FoundCell.Value
        Header.ClientName = FoundCell.Offset(0, 1).Value
        Header.CarModel = FoundCell.Offset(0, 2).Value
        Header.CarPlate = FoundCell.Offset(0, 3).Value
        Header.Mileage = Val(FoundCell.Offset(0, 4).Value)
        Header.DateIn = FoundCell.Offset(0, 5).Value
        Header.DateOut = FoundCell.Offset(0, 6).Value
        Header.Status = FoundCell.Offset(0, 7).Value
        FindOrder = True
    Else
        FindOrder = False
    End If
End Function
