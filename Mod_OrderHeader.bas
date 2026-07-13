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

    ' Поиск по № п/п в колонке A листа spisok
    Set FoundCell = wsSpisok.Columns("A").Find(What:=OrderNum, LookAt:=xlWhole, LookIn:=xlValues)

    If Not FoundCell Is Nothing Then
        ' Заполняем B3:B15
        ' Маппинг в соответствии с реальной структурой листа spisok:
        ' A=№ п/п, B=Модель, C=ГРЗ, D=VIN, E=гараж.№, F=год вып., G=пробег, H=дата
        ' Номер заказ-наряда формируется по формуле: "00"&B2&"-20"
        wsMain.Range("B3").Value = "00" & wsMain.Range("B2").Value & "-20"           ' Номер заказ-наряда
        wsMain.Range("B4").Value = FoundCell.Offset(0, 1).Value                      ' B: Модель
        wsMain.Range("B5").Value = FoundCell.Offset(0, 2).Value                      ' C: ГРЗ
        wsMain.Range("B6").Value = FoundCell.Offset(0, 3).Value                      ' D: VIN
        wsMain.Range("B7").Value = FoundCell.Offset(0, 4).Value                      ' E: гараж. №
        wsMain.Range("B8").Value = FoundCell.Offset(0, 5).Value                      ' F: год вып.
        wsMain.Range("B9").Value = FoundCell.Offset(0, 6).Value                      ' G: пробег
        wsMain.Range("B10").Value = FoundCell.Offset(0, 7).Value                     ' H: дата

        ' Поиск модели по коду
        ModelCode = FoundCell.Offset(0, 8).Value                                     ' I: Код модели
        If ModelCode <> "" Then
            Set ModelFound = wsModel.Columns("A").Find(What:=ModelCode, LookAt:=xlWhole, LookIn:=xlValues)
            If Not ModelFound Is Nothing Then
                wsMain.Range("B11").Value = ModelFound.Value                         ' Модель
                wsMain.Range("B12").Value = ModelFound.Offset(0, 1).Value            ' Цвет
                wsMain.Range("B13").Value = ModelFound.Offset(0, 2).Value            ' Год выпуска
                wsMain.Range("B14").Value = ModelFound.Offset(0, 3).Value            ' VIN
            End If
        End If

        wsMain.Range("B15").Value = FoundCell.Offset(0, 9).Value                     ' J: Примечание
    Else
        ' Очищаем B3:B15
        wsMain.Range("B3:B15").ClearContents
    End If
End Sub

' Публичная функция для тестов: поиск заказа
' Ищет по № п/п (колонка A) на листе spisok
Public Function FindOrder(ByVal OrderNum As String, ByRef Header As OrderHeader) As Boolean
    Dim ws As Worksheet
    Dim FoundCell As Range

    Set ws = GetSheetByName(ThisWorkbook, "spisok")
    If ws Is Nothing Then
        FindOrder = False
        Exit Function
    End If

    ' Поиск по колонке A (№ п/п)
    Set FoundCell = ws.Columns("A").Find(What:=OrderNum, LookAt:=xlWhole, LookIn:=xlValues)

    If Not FoundCell Is Nothing Then
        ' Маппинг в соответствии с реальной структурой листа spisok:
        ' A=№ п/п, B=Модель, C=ГРЗ, D=VIN, E=гараж.№, F=год вып., G=пробег, H=дата
        Header.OrderNumber = FoundCell.Value                    ' A: № п/п
        Header.ModelName = FoundCell.Offset(0, 1).Value         ' B: Модель
        Header.GRZ = FoundCell.Offset(0, 2).Value               ' C: ГРЗ
        Header.VIN = FoundCell.Offset(0, 3).Value               ' D: VIN
        Header.GarageNumber = FoundCell.Offset(0, 4).Value      ' E: гараж. №
        Header.YearMade = Val(FoundCell.Offset(0, 5).Value)     ' F: год вып.
        Header.MileageValue = Val(FoundCell.Offset(0, 6).Value) ' G: пробег
        Header.DateValue = FoundCell.Offset(0, 7).Value         ' H: дата
        FindOrder = True
    Else
        FindOrder = False
    End If
End Function