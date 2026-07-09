Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================================
' Модуль: Mod_Import
' Назначение: Импорт данных из файла report.xlsx в книгу заказ-наряда
' Автор: SourceCraft
' Дата: 2026-07-09
' ============================================================================

' ============================================================================
' Вспомогательные приватные функции
' ============================================================================

' --------------------------------------------------------------------------
' Функция: FindFirstCellContaining
' Ищет на листе ws первую ячейку (сверху вниз, слева направо),
' содержащую подстроку searchText (без учёта регистра).
' Поиск ведётся в UsedRange, но не более чем по maxRows строкам.
' Возвращает Nothing, если не найдено.
' --------------------------------------------------------------------------
Private Function FindFirstCellContaining(ByVal ws As Worksheet, _
                                         ByVal searchText As String, _
                                         Optional ByVal maxRows As Long = 30) As Range
    Dim rng As Range
    Dim cell As Range
    Dim rowLimit As Long
    Dim i As Long, j As Long
    
    ' Определяем область поиска: UsedRange, но не более maxRows строк
    Set rng = ws.UsedRange
    If rng Is Nothing Then
        Set FindFirstCellContaining = Nothing
        Exit Function
    End If
    
    rowLimit = rng.Rows.Count
    If rowLimit > maxRows Then rowLimit = maxRows
    
    ' Поиск перебором (первая сверху, слева направо)
    For i = 1 To rowLimit
        For j = 1 To rng.Columns.Count
            Set cell = ws.Cells(rng.Row + i - 1, rng.Column + j - 1)
            If VarType(cell.Value) = vbString Then
                If InStr(1, CStr(cell.Value), searchText, vbTextCompare) > 0 Then
                    Set FindFirstCellContaining = cell
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    Set FindFirstCellContaining = Nothing
End Function

' --------------------------------------------------------------------------
' Функция: FindHeaderInColumnC
' Ищет в столбце C листа ws ячейку, содержащую подстроку searchText
' (без учёта регистра). Поиск сверху вниз, не более maxRows строк.
' Возвращает Nothing, если не найдено.
' --------------------------------------------------------------------------
Private Function FindHeaderInColumnC(ByVal ws As Worksheet, _
                                     ByVal searchText As String, _
                                     Optional ByVal maxRows As Long = 100) As Range
    Dim i As Long
    Dim cellVal As String
    
    For i = 1 To maxRows
        If VarType(ws.Cells(i, 3).Value) = vbString Then
            cellVal = CStr(ws.Cells(i, 3).Value)
            If InStr(1, cellVal, searchText, vbTextCompare) > 0 Then
                Set FindHeaderInColumnC = ws.Cells(i, 3)
                Exit Function
            End If
        End If
    Next i
    
    Set FindHeaderInColumnC = Nothing
End Function

' --------------------------------------------------------------------------
' Функция: FindRowWithTextInColumnC
' Ищет в столбце C листа ws строку, начиная с startRow, содержащую searchText.
' Возвращает номер строки или 0, если не найдено.
' --------------------------------------------------------------------------
Private Function FindRowWithTextInColumnC(ByVal ws As Worksheet, _
                                          ByVal searchText As String, _
                                          ByVal startRow As Long, _
                                          Optional ByVal maxRows As Long = 100) As Long
    Dim i As Long
    Dim cellVal As String
    
    For i = startRow To startRow + maxRows
        If VarType(ws.Cells(i, 3).Value) = vbString Then
            cellVal = CStr(ws.Cells(i, 3).Value)
            If InStr(1, cellVal, searchText, vbTextCompare) > 0 Then
                FindRowWithTextInColumnC = i
                Exit Function
            End If
        End If
    Next i
    
    FindRowWithTextInColumnC = 0
End Function

' --------------------------------------------------------------------------
' Функция: GetLastNonEmptyRowInRange
' Возвращает номер последней непустой строки в указанном диапазоне столбцов.
' --------------------------------------------------------------------------
Private Function GetLastNonEmptyRowInRange(ByVal ws As Worksheet, _
                                           ByVal colStart As Long, _
                                           ByVal colEnd As Long) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim j As Long
    
    lastRow = 0
    For j = colStart To colEnd
        i = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        If i > lastRow Then lastRow = i
    Next j
    
    GetLastNonEmptyRowInRange = lastRow
End Function

' ============================================================================
' Публичные функции
' ============================================================================

' --------------------------------------------------------------------------
' Функция: ExtractNumberFromGRZ
' Извлекает из строки grz первую непрерывную последовательность
' из 3 или 4 цифр подряд.
'
' Примеры:
'   "Автомобиль : Лада Веста г/н Е 833 ОЕ/15" -> "833"
'   "А0663 95" -> "0663"
'
' Возвращает пустую строку, если последовательность не найдена или равна "0".
' --------------------------------------------------------------------------
Public Function ExtractNumberFromGRZ(ByVal grz As String) As String
    Dim i As Long
    Dim digitSeq As String
    Dim ch As String
    
    digitSeq = ""
    
    For i = 1 To Len(grz)
        ch = Mid(grz, i, 1)
        If ch >= "0" And ch <= "9" Then
            digitSeq = digitSeq & ch
        Else
            ' Проверяем накопленную последовательность
            If Len(digitSeq) = 3 Or Len(digitSeq) = 4 Then
                ' Проверяем, что это не "0"
                If digitSeq <> "0" Then
                    ExtractNumberFromGRZ = digitSeq
                    Exit Function
                End If
            End If
            digitSeq = ""
        End If
    Next i
    
    ' Проверяем остаток в конце строки
    If Len(digitSeq) = 3 Or Len(digitSeq) = 4 Then
        If digitSeq <> "0" Then
            ExtractNumberFromGRZ = digitSeq
            Exit Function
        End If
    End If
    
    ExtractNumberFromGRZ = ""
End Function

' --------------------------------------------------------------------------
' Функция: RenameSheetsByGRZ
' Открывает книгу report.xlsx, переименовывает видимые листы
' (кроме "report" и "spisok") в цифровой код, извлечённый из ГРЗ.
'
' Параметры:
'   reportPath - полный путь к файлу report.xlsx
'
' Возвращает True, если хотя бы один лист был переименован.
' --------------------------------------------------------------------------
Public Function RenameSheetsByGRZ(ByVal reportPath As String) As Boolean
    Const PROC_NAME As String = "RenameSheetsByGRZ"
    
    Dim wbReport As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim grzNumber As String
    Dim renamedCount As Long
    Dim sheetName As String
    Dim existingWs As Worksheet
    
    On Error GoTo ErrHandler
    
    renamedCount = 0
    
    ' Проверяем существование файла
    If Not Mod_Utils.FileExists(reportPath) Then
        MsgBox "Файл report.xlsx не найден по пути:" & vbCrLf & reportPath, _
               vbExclamation, PROC_NAME
        RenameSheetsByGRZ = False
        Exit Function
    End If
    
    ' Открываем книгу в режиме записи, невидимо
    Set wbReport = Workbooks.Open(reportPath, ReadOnly:=False, UpdateLinks:=0)
    wbReport.Windows(1).Visible = False
    
    ' Перебираем все видимые листы
    For Each ws In wbReport.Sheets
        If ws.Visible = xlSheetVisible Then
            sheetName = ws.Name
            
            ' Пропускаем листы "report" и "spisok" (без учёта регистра)
            If StrComp(sheetName, "report", vbTextCompare) <> 0 And _
               StrComp(sheetName, "spisok", vbTextCompare) <> 0 Then
                
                ' Ищем первую ячейку, содержащую "автомобиль"
                Set cell = FindFirstCellContaining(ws, "автомобиль", 30)
                
                If Not cell Is Nothing Then
                    ' Извлекаем номер из полного значения ячейки
                    grzNumber = ExtractNumberFromGRZ(CStr(cell.Value))
                    
                    If Len(grzNumber) > 0 Then
                        ' Если лист с таким именем уже существует, удаляем его
                        Application.DisplayAlerts = False
                        On Error Resume Next
                        Set existingWs = Nothing
                        Set existingWs = wbReport.Sheets(grzNumber)
                        If Not existingWs Is Nothing Then
                            existingWs.Delete
                        End If
                        Err.Clear
                        On Error GoTo ErrHandler
                        Application.DisplayAlerts = True
                        
                        ' Переименовываем лист
                        ws.Name = grzNumber
                        renamedCount = renamedCount + 1
                    End If
                End If
            End If
        End If
    Next ws
    
    ' Сохраняем и закрываем
    wbReport.Close SaveChanges:=True
    
    Application.DisplayAlerts = True
    RenameSheetsByGRZ = (renamedCount > 0)
    Exit Function
    
ErrHandler:
    ' Восстанавливаем DisplayAlerts
    Application.DisplayAlerts = True
    
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    
    ' Пытаемся закрыть книгу, если она открыта
    On Error Resume Next
    If Not wbReport Is Nothing Then
        Application.DisplayAlerts = False
        wbReport.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    RenameSheetsByGRZ = False
End Function

' --------------------------------------------------------------------------
' Функция: SearchSheetByGRZ
' В открытой книге reportWorkbook ищет лист с именем, равным grzNumber.
'
' Возвращает объект Worksheet или Nothing.
' --------------------------------------------------------------------------
Public Function SearchSheetByGRZ(ByVal reportWorkbook As Workbook, _
                                 ByVal grzNumber As String) As Worksheet
    On Error Resume Next
    Set SearchSheetByGRZ = reportWorkbook.Sheets(grzNumber)
    If Err.Number <> 0 Then Set SearchSheetByGRZ = Nothing
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' Функция: ImportSheet
' Копирует sourceSheet перед листом "main" в mainBook.
' Переименовывает скопированный лист в newSheetName.
' Не удаляет существующий лист с таким именем (Excel создаст суффикс).
'
' Возвращает True при успехе.
' --------------------------------------------------------------------------
Public Function ImportSheet(ByVal sourceSheet As Worksheet, _
                            ByVal mainBook As Workbook, _
                            ByVal newSheetName As String) As Boolean
    Const PROC_NAME As String = "ImportSheet"
    
    On Error GoTo ErrHandler
    
    ' Получаем ссылку на лист "main" в целевой книге
    Dim wsMain As Worksheet
    Set wsMain = Nothing
    On Error Resume Next
    Set wsMain = mainBook.Sheets("main")
    On Error GoTo ErrHandler
    
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в целевой книге.", vbExclamation, PROC_NAME
        ImportSheet = False
        Exit Function
    End If
    
    ' Копируем лист перед листом "main"
    sourceSheet.Copy Before:=wsMain
    
    ' Переименовываем скопированный лист
    ' После Copy активным становится новый лист - используем ActiveSheet
    On Error Resume Next
    ActiveSheet.Name = newSheetName
    If Err.Number <> 0 Then
        ' Если имя занято, Excel уже создал суффикс - это нормально
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ImportSheet = True
    Exit Function
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    ImportSheet = False
End Function

' --------------------------------------------------------------------------
' Функция: ImportIncomingDocNumber
' Заглушка. Возвращает пустую строку.
' --------------------------------------------------------------------------
Public Function ImportIncomingDocNumber(ByVal sourceSheet As Worksheet) As String
    ImportIncomingDocNumber = ""
End Function

' --------------------------------------------------------------------------
' Функция: ImportDataToMain
' Импортирует данные из sourceSheet на mainSheet.
'
' Таблица "Выполненные работы":
'   - Ищем заголовок "Наименование" в столбце C
'   - Собираем строки данных до строки "Итого работ"
'   - C -> L (Наименование), D -> M (Кол-во оп.), H -> N (Всего)
'
' Таблица "Расходная накладная":
'   - Ищем "Расходная накладная" или второе "Наименование" в столбце C
'   - Собираем строки данных до пустой строки
'   - B -> X (№ кат.), C -> Y (Наименование), D -> Z (Кол-во), G -> AA (Всего)
' --------------------------------------------------------------------------
Public Sub ImportDataToMain(ByVal sourceSheet As Worksheet, _
                            ByVal mainSheet As Worksheet)
    Const PROC_NAME As String = "ImportDataToMain"
    
    Dim lastRowL As Long
    Dim lastRowX As Long
    Dim i As Long
    Dim dataRow As Long
    Dim destRow As Long
    
    On Error GoTo ErrHandler
    
    ' --- Очистка целевых диапазонов ---
    ' Определяем последнюю непустую строку для L:N
    lastRowL = GetLastNonEmptyRowInRange(mainSheet, 12, 14) ' L=12, N=14
    If lastRowL >= 2 Then
        mainSheet.Range("L2:N" & lastRowL).ClearContents
    End If
    
    ' Определяем последнюю непустую строку для X:AA
    lastRowX = GetLastNonEmptyRowInRange(mainSheet, 24, 27) ' X=24, AA=27
    If lastRowX >= 2 Then
        mainSheet.Range("X2:AA" & lastRowX).ClearContents
    End If
    
    ' --- Импорт таблицы "Выполненные работы" ---
    Dim headerRowWorks As Long
    Dim итогоRowWorks As Long
    
    ' Ищем заголовок "Наименование" в столбце C
    headerRowWorks = FindRowWithTextInColumnC(sourceSheet, "Наименование", 1, 100)
    
    If headerRowWorks > 0 Then
        ' Ищем строку "Итого работ" после заголовка
        итогоRowWorks = FindRowWithTextInColumnC(sourceSheet, "Итого работ", headerRowWorks + 1, 100)
        
        If итогоRowWorks > 0 Then
            ' Собираем строки данных между заголовком и "Итого работ"
            destRow = 2
            For dataRow = headerRowWorks + 1 To итогоRowWorks - 1
                ' Проверяем, что строка не пуста (хотя бы одна ячейка не пуста)
                If Len(Trim(CStr(sourceSheet.Cells(dataRow, 3).Value))) > 0 Then
                    ' Наименование (C) -> L
                    mainSheet.Cells(destRow, 12).Value = sourceSheet.Cells(dataRow, 3).Value
                    ' Кол-во оп. (D) -> M
                    mainSheet.Cells(destRow, 13).Value = sourceSheet.Cells(dataRow, 4).Value
                    ' Всего (H) -> N
                    mainSheet.Cells(destRow, 14).Value = sourceSheet.Cells(dataRow, 8).Value
                    destRow = destRow + 1
                End If
            Next dataRow
        Else
            MsgBox "Не найдена строка 'Итого работ' на листе '" & sourceSheet.Name & "'.", _
                   vbExclamation, PROC_NAME
        End If
    Else
        MsgBox "Не найден заголовок таблицы 'Выполненные работы' на листе '" & sourceSheet.Name & "'.", _
               vbExclamation, PROC_NAME
    End If
    
    ' --- Импорт таблицы "Расходная накладная" ---
    Dim расходRow As Long
    Dim headerRowParts As Long
    Dim partsStartRow As Long
    
    ' Сначала ищем "Расходная накладная"
    расходRow = FindRowWithTextInColumnC(sourceSheet, "Расходная накладная", 1, 200)
    
    If расходRow > 0 Then
        ' Ищем заголовок "Наименование" после строки "Расходная накладная"
        headerRowParts = FindRowWithTextInColumnC(sourceSheet, "Наименование", расходRow + 1, 50)
    Else
        ' Ищем второе вхождение "Наименование" в столбце C
        Dim firstНаименование As Long
        firstНаименование = FindRowWithTextInColumnC(sourceSheet, "Наименование", 1, 100)
        If firstНаименование > 0 Then
            headerRowParts = FindRowWithTextInColumnC(sourceSheet, "Наименование", firstНаименование + 1, 100)
        Else
            headerRowParts = 0
        End If
    End If
    
    If headerRowParts > 0 Then
        ' Собираем строки данных после заголовка до пустой строки
        destRow = 2
        For dataRow = headerRowParts + 1 To headerRowParts + 200
            ' Проверяем, что строка не пуста (хотя бы одна ячейка в B:D не пуста)
            If Len(Trim(CStr(sourceSheet.Cells(dataRow, 2).Value))) = 0 And _
               Len(Trim(CStr(sourceSheet.Cells(dataRow, 3).Value))) = 0 And _
               Len(Trim(CStr(sourceSheet.Cells(dataRow, 4).Value))) = 0 Then
                Exit For ' Дошли до пустой строки
            End If
            
            ' № кат. (B) -> X
            mainSheet.Cells(destRow, 24).Value = sourceSheet.Cells(dataRow, 2).Value
            ' Наименование (C) -> Y
            mainSheet.Cells(destRow, 25).Value = sourceSheet.Cells(dataRow, 3).Value
            ' Кол-во (D) -> Z
            mainSheet.Cells(destRow, 26).Value = sourceSheet.Cells(dataRow, 4).Value
            ' Всего (G) -> AA
            mainSheet.Cells(destRow, 27).Value = sourceSheet.Cells(dataRow, 7).Value
            
            destRow = destRow + 1
        Next dataRow
    Else
        MsgBox "Не найдена таблица 'Расходная накладная' на листе '" & sourceSheet.Name & "'.", _
               vbExclamation, PROC_NAME
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
End Sub

' ============================================================================
' Главная процедура
' ============================================================================

' --------------------------------------------------------------------------
' Процедура: ImportFromReport
' Главная процедура импорта, вызывается по кнопке "ИМПОРТ".
'
' Последовательность:
'   1. Отключает ScreenUpdating и EnableEvents
'   2. Вызывает RenameSheetsByGRZ
'   3. Получает ГРЗ из main!B4, извлекает номер
'   4. Открывает report.xlsx только для чтения
'   5. Ищет лист по номеру ГРЗ
'   6. Формирует имя нового листа: <номер заказа>M
'   7. Копирует лист в текущую книгу
'   8. Импортирует данные на лист main
'   9. Закрывает report.xlsx
'   10. Восстанавливает настройки
' --------------------------------------------------------------------------
Public Sub ImportFromReport()
    Const PROC_NAME As String = "ImportFromReport"
    
    Dim wsMain As Worksheet
    Dim grzValue As String
    Dim grzNumber As String
    Dim reportPath As String
    Dim wbReport As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheetName As String
    Dim importedSheet As Worksheet
    Dim orderNumber As Variant
    
    On Error GoTo ErrHandler
    
    ' Отключаем обновление экрана и события
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Получаем лист main
    Set wsMain = Nothing
    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets("main")
    On Error GoTo ErrHandler
    
    If wsMain Is Nothing Then
        MsgBox "Лист 'main' не найден в текущей книге.", vbExclamation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' --- Шаг 1: Переименование листов в report.xlsx ---
    reportPath = ThisWorkbook.Path
    If Right(reportPath, 1) <> "\" Then reportPath = reportPath & "\"
    reportPath = reportPath & "report.xlsx"
    
    If Not RenameSheetsByGRZ(reportPath) Then
        MsgBox "Не удалось переименовать листы в файле report.xlsx." & vbCrLf & _
               "Возможно, файл не найден или на листах не обнаружены ГРЗ.", _
               vbInformation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' --- Шаг 2: Получение ГРЗ ---
    grzValue = Trim(CStr(wsMain.Range("B4").Value))
    
    If Len(grzValue) = 0 Then
        MsgBox "Ячейка B4 (ГРЗ) на листе 'main' пуста.", vbInformation, PROC_NAME
        GoTo CleanUp
    End If
    
    grzNumber = ExtractNumberFromGRZ(grzValue)
    
    If Len(grzNumber) = 0 Then
        MsgBox "Не удалось извлечь номер из ГРЗ: """ & grzValue & """", _
               vbInformation, PROC_NAME
        GoTo CleanUp
    End If
    
    ' --- Шаг 3: Открытие report.xlsx только для чтения ---
    Set wbReport = Workbooks.Open(reportPath, ReadOnly:=True, UpdateLinks:=0)
    wbReport.Windows(1).Visible = False
    
    ' --- Шаг 4: Поиск листа по номеру ГРЗ ---
    Set sourceSheet = SearchSheetByGRZ(wbReport, grzNumber)
    
    If sourceSheet Is Nothing Then
        MsgBox "В файле report.xlsx не найден лист с номером """ & grzNumber & """", _
               vbInformation, PROC_NAME
        GoTo CloseReport
    End If
    
    ' --- Шаг 5: Формирование имени нового листа ---
    orderNumber = wsMain.Range("B2").Value
    If IsEmpty(orderNumber) Then
        MsgBox "Ячейка B2 (номер заказа) на листе 'main' пуста.", _
               vbInformation, PROC_NAME
        GoTo CloseReport
    End If
    newSheetName = CStr(orderNumber) & "M"
    
    ' --- Шаг 6: Копирование листа ---
    If Not ImportSheet(sourceSheet, ThisWorkbook, newSheetName) Then
        MsgBox "Не удалось скопировать лист.", vbExclamation, PROC_NAME
        GoTo CloseReport
    End If
    
    ' --- Шаг 7: Получение ссылки на скопированный лист ---
    On Error Resume Next
    Set importedSheet = ThisWorkbook.Sheets(newSheetName)
    If Err.Number <> 0 Then
        ' Если точное имя не найдено (Excel добавил суффикс),
        ' ищем последний добавленный лист
        Set importedSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' --- Шаг 8: Импорт данных на лист main ---
    ImportDataToMain importedSheet, wsMain
    
    ' --- Шаг 9: Импорт входящего номера документа (заглушка) ---
    Dim docNumber As String
    docNumber = ImportIncomingDocNumber(importedSheet)
    
    ' --- Шаг 10: Закрытие report.xlsx ---
    wbReport.Close SaveChanges:=False
    Set wbReport = Nothing
    
    ' Восстанавливаем настройки
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Импорт данных успешно завершён.", vbInformation, PROC_NAME
    Exit Sub
    
CloseReport:
    ' Закрываем report.xlsx без сохранения, если открыта
    On Error Resume Next
    If Not wbReport Is Nothing Then
        wbReport.Close SaveChanges:=False
        Set wbReport = Nothing
    End If
    On Error GoTo 0
    GoTo CleanUp
    
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Ошибка в " & PROC_NAME & ":" & vbCrLf & vbCrLf & _
           "Номер: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, PROC_NAME
    
    ' Восстанавливаем настройки
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' Закрываем report.xlsx, если открыта
    On Error Resume Next
    If Not wbReport Is Nothing Then
        wbReport.Close SaveChanges:=False
        Set wbReport = Nothing
    End If
    On Error GoTo 0
End Sub