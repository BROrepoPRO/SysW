Attribute VB_Name = "Mod_ButtonDispatcher"
Option Explicit

' ============================================================================
' Модуль: Mod_ButtonDispatcher
' Назначение: Диспетчеризация нажатий на кнопки листа
' Автор: SourceCraft
' Дата: 2026-07-09
' ============================================================================

' ============================================================================
' Обработчики кнопок листа main
' ============================================================================

' --------------------------------------------------------------------------
' Btn_main_HeaderFill - заполнение шапки заказ-наряда
' Вызывает Mod_OrderHeader.FillHeaderFromOrder по значению из main!B2,
' передавая main, spisok, model.
' --------------------------------------------------------------------------
Public Sub Btn_main_HeaderFill()
    Const PROC_NAME As String = "Btn_main_HeaderFill"
    
    On Error GoTo ErrHandler
    
    ' Отключаем обновление экрана и события
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Получаем лист main
    Dim wsMain As Worksheet
    Set wsMain = Mod_Utils.GetSheetByName("main")
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Получаем лист spisok
    Dim wsSpisok As Worksheet
    Set wsSpisok = Mod_Utils.GetSheetByName("spisok")
    If wsSpisok Is Nothing Then
        MsgBox "Лист 'spisok' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Получаем лист model
    Dim wsModel As Worksheet
    Set wsModel = Mod_Utils.GetSheetByName("model")
    If wsModel Is Nothing Then
        MsgBox "Лист 'model' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Читаем значение заказ-наряда из ячейки B2
    Dim orderValue As Variant
    orderValue = wsMain.Range("B2").Value
    
    ' Проверяем, что значение не пусто
    If IsEmpty(orderValue) Then
        MsgBox "Ячейка B2 на листе 'main' пуста. Введите номер заказ-наряда.", _
               vbInformation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Вызываем функцию заполнения шапки
    Dim result As Boolean
    result = Mod_OrderHeader.FillHeaderFromOrder(orderValue, wsMain, wsSpisok, wsModel)
    
    If result Then
        MsgBox "Шапка заказ-наряда успешно заполнена.", vbInformation, PROC_NAME
    End If
    
CleanUp:
    ' Восстанавливаем настройки
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_ClearHeader - очистка шапки (ячейки B3:B15 на листе main)
' --------------------------------------------------------------------------
Public Sub Btn_main_ClearHeader()
    Const PROC_NAME As String = "Btn_main_ClearHeader"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wsMain As Worksheet
    Set wsMain = Mod_Utils.GetSheetByName("main")
    
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Очищаем ячейки B3:B15
    wsMain.Range("B3:B15").ClearContents
    
    MsgBox "Шапка заказ-наряда очищена.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_ClearAll - очистка всех данных заказ-наряда на листе main
' Очищает B3:B15 и таблицу заказ/накладной (строки 17-200).
' --------------------------------------------------------------------------
Public Sub Btn_main_ClearAll()
    Const PROC_NAME As String = "Btn_main_ClearAll"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wsMain As Worksheet
    Set wsMain = Mod_Utils.GetSheetByName("main")
    
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Очищаем шапку
    wsMain.Range("B3:B15").ClearContents
    
    ' Очищаем таблицу заказ/накладной (строки 17-200)
    wsMain.Range("A17:Z200").ClearContents
    
    MsgBox "Все данные заказ-наряда на листе 'main' очищены.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' ============================================================================
' Вспомогательные функции для Btn_main_AutoSearch
' ============================================================================

' --------------------------------------------------------------------------
' Функция: FindSheetsByGRZ
' Ищет в книге wbReport все видимые листы, содержащие grzValue в имени.
' Возвращает коллекцию найденных листов.
' --------------------------------------------------------------------------
Private Function FindSheetsByGRZ(ByVal wbReport As Workbook, _
                                 ByVal grzValue As String) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    Dim sheetName As String
    
    For Each ws In wbReport.Sheets
        If ws.Visible = xlSheetVisible Then
            sheetName = ws.Name
            If InStr(1, sheetName, grzValue, vbTextCompare) > 0 Then
                result.Add ws
            End If
        End If
    Next ws
    
    Set FindSheetsByGRZ = result
End Function

' --------------------------------------------------------------------------
' Функция: ShowSheetSelectionDialog
' Показывает диалог со списком найденных листов и выбором действия.
' Возвращает: vbYes - переименовать, vbNo - переместить, vbCancel - копировать.
' --------------------------------------------------------------------------
Private Function ShowSheetSelectionDialog(ByVal foundSheets As Collection, _
                                          ByVal grzValue As String) As VbMsgBoxResult
    Dim listMsg As String
    Dim ws As Worksheet
    Dim i As Long
    
    listMsg = "Найдены листы с ГРЗ """ & grzValue & """:" & vbCrLf & vbCrLf
    
    i = 1
    For Each ws In foundSheets
        listMsg = listMsg & i & ". " & ws.Name & vbCrLf
        i = i + 1
    Next ws
    
    listMsg = listMsg & vbCrLf & "Выберите действие:" & vbCrLf & vbCrLf & _
              "Да - переименовать лист в report.xlsx" & vbCrLf & _
              "Нет - переместить лист в текущую книгу" & vbCrLf & _
              "Отмена - скопировать лист в текущую книгу"
    
    ShowSheetSelectionDialog = MsgBox(listMsg, vbYesNoCancel + vbQuestion, "Btn_main_AutoSearch")
End Function

' --------------------------------------------------------------------------
' Функция: RenameSheetInReport
' Переименовывает лист targetWs в книге report.xlsx.
' --------------------------------------------------------------------------
Private Sub RenameSheetInReport(ByVal targetWs As Worksheet)
    Dim newName As String
    newName = InputBox("Введите новое имя для листа """ & targetWs.Name & """ в книге report.xlsx:", _
                       "Переименование листа", targetWs.Name)
    If Len(newName) > 0 Then
        On Error Resume Next
        targetWs.Name = newName
        If Err.Number <> 0 Then
            MsgBox "Не удалось переименовать лист. Возможно, имя """ & newName & """ уже используется.", _
                   vbExclamation, "RenameSheetInReport"
        Else
            MsgBox "Лист переименован в """ & newName & """ в книге report.xlsx.", _
                   vbInformation, "RenameSheetInReport"
        End If
        On Error GoTo 0
    End If
End Sub

' --------------------------------------------------------------------------
' Функция: GenerateUniqueSheetName
' Генерирует уникальное имя листа на основе baseName, проверяя
' существование листов с таким именем в текущей книге.
' --------------------------------------------------------------------------
Private Function GenerateUniqueSheetName(ByVal baseName As String) As String
    Dim result As String
    Dim existingWs As Worksheet
    Dim suffix As Long
    
    result = baseName
    Set existingWs = Mod_Utils.GetSheetByName(result)
    suffix = 1
    Do While Not existingWs Is Nothing
        result = baseName & "_" & suffix
        Set existingWs = Mod_Utils.GetSheetByName(result)
        suffix = suffix + 1
    Loop
    
    GenerateUniqueSheetName = result
End Function

' --------------------------------------------------------------------------
' Функция: MoveSheetToCurrentBook
' Перемещает лист targetWs в текущую книгу с именем на основе grzValue.
' --------------------------------------------------------------------------
Private Sub MoveSheetToCurrentBook(ByVal targetWs As Worksheet, _
                                   ByVal grzValue As String)
    Dim moveName As String
    moveName = GenerateUniqueSheetName(grzValue)
    
    targetWs.Move Before:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = moveName
    MsgBox "Лист перемещен в текущую книгу с наименованием """ & moveName & """.", _
           vbInformation, "MoveSheetToCurrentBook"
End Sub

' --------------------------------------------------------------------------
' Функция: CopySheetToCurrentBook
' Копирует лист targetWs в текущую книгу с именем на основе grzValue.
' --------------------------------------------------------------------------
Private Sub CopySheetToCurrentBook(ByVal targetWs As Worksheet, _
                                   ByVal grzValue As String)
    Dim copyName As String
    copyName = GenerateUniqueSheetName(grzValue)
    
    targetWs.Copy Before:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = copyName
    MsgBox "Лист скопирован в текущую книгу с наименованием """ & copyName & """.", _
           vbInformation, "CopySheetToCurrentBook"
End Sub

' --------------------------------------------------------------------------
' Btn_main_AutoSearch - поиск в report.xlsx по номеру ГРЗ,
' открытие и копирование листа на лист B4 листа main.
' Далее работа с листом поиска: переименование, перемещение, копирование.
' --------------------------------------------------------------------------
Public Sub Btn_main_AutoSearch()
    Const PROC_NAME As String = "Btn_main_AutoSearch"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Получаем лист main
    Dim wsMain As Worksheet
    Set wsMain = Mod_Utils.GetSheetByName("main")
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Получаем ГРЗ из ячейки B4
    Dim grzValue As String
    grzValue = Trim(CStr(wsMain.Range("B4").Value))
    
    If Len(grzValue) = 0 Then
        MsgBox "Ячейка B4 (ГРЗ) на листе 'main' пуста. Введите соответствующий регистрационный знак.", _
               vbInformation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Определяем путь к report.xlsx
    Dim reportPath As String
    reportPath = Mod_Utils.GetWorkbookPath() & "report.xlsx"
    
    ' Проверяем существование файла
    If Not Mod_Utils.FileExists(reportPath) Then
        MsgBox "Файл report.xlsx не найден по пути:" & vbCrLf & reportPath, _
               vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Открываем report.xlsx
    Dim wbReport As Workbook
    Set wbReport = Workbooks.Open(reportPath, ReadOnly:=False, UpdateLinks:=0)
    
    ' Ищем листы по ГРЗ
    Dim foundSheets As Collection
    Set foundSheets = FindSheetsByGRZ(wbReport, grzValue)
    
    ' Проверяем, сколько нашли
    If foundSheets.Count = 0 Then
        MsgBox "В файле report.xlsx не найдено ни одного листа, содержащего """ & grzValue & """ в названии.", _
               vbInformation, PROC_NAME
        wbReport.Close SaveChanges:=False
        GoTo CleanUp
    End If
    
    ' Показываем диалог выбора
    Dim userChoice As VbMsgBoxResult
    userChoice = ShowSheetSelectionDialog(foundSheets, grzValue)
    
    ' Используем первый лист (для упрощения логики)
    Dim targetWs As Worksheet
    Set targetWs = foundSheets(1)
    
    If foundSheets.Count > 1 Then
        MsgBox "Найдено несколько листов. Будет использован первый: """ & targetWs.Name & """." & vbCrLf & _
               "Для выбора другого листа откройте файл вручную.", vbInformation, PROC_NAME
    End If
    
    ' Выполняем выбранное действие
    Select Case userChoice
        Case vbYes
            RenameSheetInReport targetWs
        Case vbNo
            MoveSheetToCurrentBook targetWs, grzValue
        Case vbCancel
            CopySheetToCurrentBook targetWs, grzValue
    End Select
    
    ' Закрываем report.xlsx с сохранением
    wbReport.Close SaveChanges:=True
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_Import - импорт данных из report.xlsx
' Вызывает Mod_Import.ImportFromReport для импорта данных по ГРЗ.
' Управление ScreenUpdating/EnableEvents внутри ImportFromReport.
' --------------------------------------------------------------------------
Public Sub Btn_main_Import()
    Const PROC_NAME As String = "Btn_main_Import"
    
    On Error GoTo ErrHandler
    
    ' Вызываем главную процедуру импорта
    Mod_Import.ImportFromReport
    
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
End Sub

' --------------------------------------------------------------------------
' Btn_main_PickWorkAuto - автоматический подбор работ (заглушка)
' --------------------------------------------------------------------------
Public Sub Btn_main_PickWorkAuto()
    Const PROC_NAME As String = "Btn_main_PickWorkAuto"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Автоматический подбор работ пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_PickWorkManual - ручной подбор работ (заглушка)
' --------------------------------------------------------------------------
Public Sub Btn_main_PickWorkManual()
    Const PROC_NAME As String = "Btn_main_PickWorkManual"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Ручной подбор работ пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_PickPartAuto - автоматический подбор запчастей (заглушка)
' --------------------------------------------------------------------------
Public Sub Btn_main_PickPartAuto()
    Const PROC_NAME As String = "Btn_main_PickPartAuto"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Автоматический подбор запчастей пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_PickPartManual - ручной подбор запчастей (заглушка)
' --------------------------------------------------------------------------
Public Sub Btn_main_PickPartManual()
    Const PROC_NAME As String = "Btn_main_PickPartManual"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Ручной подбор запчастей пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_NZToText - преобразование цифр в текст в ячейках листа НЗ
' Вызывает Mod_ToText.toTEXT (если модуль существует и доступен).
' --------------------------------------------------------------------------
Public Sub Btn_main_NZToText()
    Const PROC_NAME As String = "Btn_main_NZToText"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Модуль Mod_ToText не установлен. Заглушка.
    Debug.Print "toTEXT: модуль не установлен"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_Export - экспорт листов "НЗ" и "Счет" в отдельные файлы
' Имена файлов: НЗ00<номер B2>-20.xlsx и Счет00<номер B2>-20.xlsx
' --------------------------------------------------------------------------
Public Sub Btn_main_Export()
    Const PROC_NAME As String = "Btn_main_Export"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Получаем лист main
    Dim wsMain As Worksheet
    Set wsMain = Mod_Utils.GetSheetByName("main")
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Получаем номер заказ-наряда из B2
    Dim orderValue As Variant
    orderValue = wsMain.Range("B2").Value
    
    If IsEmpty(orderValue) Then
        MsgBox "Ячейка B2 (номер заказ-наряда) пуста. Заполните сначала.", _
               vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    Dim orderStr As String
    orderStr = CStr(orderValue)
    
    ' Определяем путь сохранения
    Dim basePath As String
    basePath = Mod_Utils.GetWorkbookPath()
    
    Dim fileZN As String
    Dim fileSchet As String
    fileZN = basePath & "НЗ00" & orderStr & "-20.xlsx"
    fileSchet = basePath & "Счет00" & orderStr & "-20.xlsx"
    
    ' Проверяем наличие листов
    Dim wsZN As Worksheet
    Dim wsSchet As Worksheet
    Dim hasZN As Boolean
    Dim hasSchet As Boolean
    
    Set wsZN = Mod_Utils.GetSheetByName("НЗ")
    hasZN = Not (wsZN Is Nothing)
    
    Set wsSchet = Mod_Utils.GetSheetByName("Счет")
    hasSchet = Not (wsSchet Is Nothing)
    
    If Not hasZN And Not hasSchet Then
        MsgBox "Листы 'НЗ' и 'Счет' не найдены в текущей книге. Экспорт невозможен.", _
               vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' Экспортируем лист "НЗ"
    If hasZN Then
        wsZN.Copy
        ActiveWorkbook.SaveAs fileName:=fileZN, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=True
    End If
    
    ' Экспортируем лист "Счет"
    If hasSchet Then
        wsSchet.Copy
        ActiveWorkbook.SaveAs fileName:=fileSchet, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=True
    End If
    
    Dim exportMsg As String
    exportMsg = "Экспорт завершен:" & vbCrLf
    If hasZN Then exportMsg = exportMsg & "  - НЗ00" & orderStr & "-20.xlsx" & vbCrLf
    If hasSchet Then exportMsg = exportMsg & "  - Счет00" & orderStr & "-20.xlsx" & vbCrLf
    exportMsg = exportMsg & vbCrLf & "Файлы сохранены в: " & basePath
    
    MsgBox exportMsg, vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_main_Report - формирование сводного отчета в report.xlsx из листа report
' Заготовка: требуется уточнение и доработка логики.
' --------------------------------------------------------------------------
Public Sub Btn_main_Report()
    Const PROC_NAME As String = "Btn_main_Report"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim headers As String
    headers = "Номер п/п, Номер з/н, Марка/Модель, ГРЗ, VIN, Кузов, Год вып., Цвет, " & _
              "Пробег, Кузов/Салон, Тип кузова, Цвет кузова, Тип/Модель, Цвет кузова, " & _
              "Двигатель, № дв., Номер шасси, Кузов, № куз., Номер двигателя, " & _
              "Мощность, %, Объем, %, Норма расхода, %"
    
    MsgBox "Функция формирования сводного отчета находится в разработке." & vbCrLf & _
           vbCrLf & "Планируемые колонки:" & vbCrLf & headers, _
           vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' ============================================================================
' Обработчики кнопок для листов z4, work и дополнительных
' ============================================================================

' --------------------------------------------------------------------------
' Btn_<лист>_Action1 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_z4_Action1()
    Const PROC_NAME As String = "Btn_z4_Action1"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие z4 Action1 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_<лист>_Action2 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_z4_Action2()
    Const PROC_NAME As String = "Btn_z4_Action2"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие z4 Action2 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_<лист>_Action3 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_z4_Action3()
    Const PROC_NAME As String = "Btn_z4_Action3"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие z4 Action3 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_work_Action1 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_work_Action1()
    Const PROC_NAME As String = "Btn_work_Action1"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие work Action1 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_work_Action2 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_work_Action2()
    Const PROC_NAME As String = "Btn_work_Action2"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие work Action2 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub

' --------------------------------------------------------------------------
' Btn_work_Action3 - заглушка
' --------------------------------------------------------------------------
Public Sub Btn_work_Action3()
    Const PROC_NAME As String = "Btn_work_Action3"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    MsgBox "Действие work Action3 пока в разработке.", vbInformation, PROC_NAME
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanUp
End Sub