Attribute VB_Name = "Mod_Import"
Option Explicit

' ============================================================
' Модуль: Mod_Import
' Назначение: Импорт данных из Excel в SQLite и обратно
' ============================================================

' ============================================================
' ОСНОВНЫЕ ФУНКЦИИ ИМПОРТА
' ============================================================

' --------------------------------------------------------------------------
' ImportSheet
' Импортирует лист из report.xlsx по ГРЗ в текущую книгу
' --------------------------------------------------------------------------
Public Sub ImportSheet(grz As String)
    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Dim wsMain As Worksheet
    Dim newName As String

    Set wsMain = ThisWorkbook.Sheets("main")

    Set wsSource = Mod_SheetOps.SearchSheetByGRZ(grz)
    If wsSource Is Nothing Then
        MsgBox "Лист с ГРЗ " & grz & " не найден!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    newName = Trim(wsMain.Range("B2").Value) & "M"
    On Error Resume Next
    ActiveSheet.Name = newName
    On Error GoTo 0

    Call ImportDataToMain(ActiveSheet)
    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при импорте данных: " & Err.Description & ". Импорт прерван.", vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_Import", "ImportSheet: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportDataToMain
' Переносит данные из листа-источника в лист main по столбцам
' Ищет таблицы "Выполненные работы" и "Расходная накладная" на листе
' --------------------------------------------------------------------------
Public Sub ImportDataToMain(wsSource As Worksheet)
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim srcLastRow As Long
    Dim foundWorks As Boolean
    Dim foundMaterials As Boolean
    Dim cell As Range
    Dim dataStartRow As Long   ' строка, с которой начинаются данные (после 2 строк заголовка)
    Dim targetRow As Long      ' целевая строка на листе main
    Dim endRow As Range        ' граница таблицы (строка "Итого")
    Dim wsRow As Range         ' для поиска границы между таблицами

    Set wsMain = ThisWorkbook.Sheets("main")

    ' Очистка диапазонов L:N и X:AA
    Dim lastRowL As Long, lastRowX As Long
    lastRowL = wsMain.Cells(wsMain.Rows.count, 12).End(xlUp).Row
    lastRowX = wsMain.Cells(wsMain.Rows.count, 24).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowL, lastRowX)
    If lastRow < 2 Then lastRow = 2

    wsMain.Range("L2:N" & lastRow).ClearContents
    wsMain.Range("X2:AA" & lastRow).ClearContents

    foundWorks = False
    foundMaterials = False

    ' ============================================================
    ' Поиск таблицы "Выполненные работы"
    ' Структура листа-источника (реальная):
    '   Строка 1: заголовок "№ | № кат. | Наименование | Кол. оп. | Цена | Норма | н/ч | Всего | в т.ч. НДС"
    '   Строка 2: подзаголовок "1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9"
    '   Строка 3+: данные
    ' Колонки: B(2)=№, C(3)=№ кат., D(4)=Наименование, E(5)=Кол.оп., F(6)=Цена,
    '          G(7)=Норма, H(8)=н/ч, I(9)=Всего, J(10)=в т.ч. НДС
    ' Маппинг на main: D(Наименование)→L(12), E(Кол.оп.)→M(13), I(Всего)→N(14)
    ' ============================================================
    Set cell = wsSource.Cells.Find(What:="Выполненные работы", LookAt:=xlPart, SearchOrder:=xlByRows)

    If Not cell Is Nothing Then
        foundWorks = True
        ' Определяем последнюю строку таблицы работ:
        ' сначала ищем "Итого работ" в столбце D на ограниченном диапазоне
        ' (от начала данных до строки с "Расходная накладная" или до конца листа)
        Set endRow = Nothing
        On Error Resume Next
        Set endRow = wsSource.Range(wsSource.Cells(cell.Row + 1, 4), wsSource.Cells(wsSource.Rows.count, 4)) _
                     .Find(What:="Итого", LookAt:=xlPart)
        On Error GoTo 0
        If Not endRow Is Nothing Then
            srcLastRow = endRow.Row - 1
        Else
            ' Если "Итого" не найдено — ищем последнюю непустую строку
            ' в столбце D, но не ниже строки с "Расходная накладная"
            Set wsRow = wsSource.Cells.Find(What:="Расходная накладная", LookAt:=xlPart, SearchOrder:=xlByRows)
            If Not wsRow Is Nothing Then
                srcLastRow = wsSource.Range(wsSource.Cells(cell.Row + 1, 4), wsSource.Cells(wsRow.Row - 1, 4)) _
                             .Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious).Row
            Else
                srcLastRow = wsSource.Cells(wsSource.Rows.count, 4).End(xlUp).Row
            End If
        End If
        ' Пропускаем пустые строки после названия таблицы
        dataStartRow = cell.Row + 1
        Do While dataStartRow <= srcLastRow
            If Trim(wsSource.Cells(dataStartRow, 4).Value) <> "" Then Exit Do
            dataStartRow = dataStartRow + 1
        Loop
        ' Пропускаем две строки заголовка (заголовок колонок + подзаголовок с номерами)
        dataStartRow = dataStartRow + 2

        targetRow = 2 ' данные на main начинаем писать со строки 2
        For i = dataStartRow To srcLastRow
            If wsSource.Cells(i, 4).Value <> "" Then
                wsMain.Cells(targetRow, 12).Value = wsSource.Cells(i, 4).Value ' D(Наименование) -> L
                wsMain.Cells(targetRow, 13).Value = wsSource.Cells(i, 5).Value ' E(Кол.оп.) -> M
                wsMain.Cells(targetRow, 14).Value = wsSource.Cells(i, 9).Value ' I(Всего) -> N
                ' Форматируем числовые колонки: убираем десятичные знаки для целых чисел
                If IsNumeric(wsMain.Cells(targetRow, 13).Value) Then
                    If wsMain.Cells(targetRow, 13).Value = Int(wsMain.Cells(targetRow, 13).Value) Then
                        wsMain.Cells(targetRow, 13).NumberFormat = "0"
                    End If
                End If
                If IsNumeric(wsMain.Cells(targetRow, 14).Value) Then
                    wsMain.Cells(targetRow, 14).NumberFormat = "# ##0,00"
                End If
                targetRow = targetRow + 1
            End If
        Next i
    End If

    ' ============================================================
    ' Поиск таблицы "Расходная накладная"
    ' Структура листа-источника (реальная):
    '   Строка 1: заголовок "№ | № кат. | Наименование | Кол-во | Ед.изм. | Цена | Всего | в т.ч. НДС"
    '   Строка 2: подзаголовок "1 | 2 | 3 | 4 | 5 | 6 | 7 | 8"
    '   Строка 3+: данные
    ' Колонки: A(1)=№, B(2)=№ кат., C(3)=Наименование, D(4)=Кол-во,
    '          E(5)=Ед.изм., F(6)=Цена, G(7)=Всего, H(8)=в т.ч. НДС
    ' Маппинг на main: B(№ кат.)→X(24), C(Наименование)→Y(25), D(Кол-во)→Z(26), G(Всего)→AA(27)
    ' ============================================================
    Set cell = wsSource.Cells.Find(What:="Расходная накладная", LookAt:=xlPart, SearchOrder:=xlByRows)

    If Not cell Is Nothing Then
        foundMaterials = True
        ' Определяем последнюю строку таблицы материалов:
        ' ищем "Итого" в столбце B начиная от строки данных до конца листа
        Set endRow = Nothing
        On Error Resume Next
        Set endRow = wsSource.Range(wsSource.Cells(cell.Row + 1, 2), wsSource.Cells(wsSource.Rows.count, 2)) _
                     .Find(What:="Итого", LookAt:=xlPart)
        On Error GoTo 0
        If Not endRow Is Nothing Then
            srcLastRow = endRow.Row - 1
        Else
            ' Если "Итого" не найдено — ищем последнюю непустую строку в столбце B
            srcLastRow = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).Row
        End If
        ' Пропускаем пустые строки после названия таблицы
        dataStartRow = cell.Row + 1
        Do While dataStartRow <= srcLastRow
            If Trim(wsSource.Cells(dataStartRow, 2).Value) <> "" Then Exit Do
            dataStartRow = dataStartRow + 1
        Loop
        ' Пропускаем две строки заголовка (заголовок колонок + подзаголовок с номерами)
        dataStartRow = dataStartRow + 2

        targetRow = 2 ' данные на main начинаем писать со строки 2
        For i = dataStartRow To srcLastRow
            If wsSource.Cells(i, 2).Value <> "" Then
                wsMain.Cells(targetRow, 24).Value = wsSource.Cells(i, 2).Value ' B(№ кат.) -> X
                wsMain.Cells(targetRow, 25).Value = wsSource.Cells(i, 3).Value ' C(Наименование) -> Y
                wsMain.Cells(targetRow, 26).Value = wsSource.Cells(i, 4).Value ' D(Кол-во) -> Z
                wsMain.Cells(targetRow, 27).Value = wsSource.Cells(i, 7).Value ' G(Всего) -> AA
                ' Форматируем числовые колонки
                If IsNumeric(wsMain.Cells(targetRow, 26).Value) Then
                    If wsMain.Cells(targetRow, 26).Value = Int(wsMain.Cells(targetRow, 26).Value) Then
                        wsMain.Cells(targetRow, 26).NumberFormat = "0"
                    End If
                End If
                If IsNumeric(wsMain.Cells(targetRow, 27).Value) Then
                    wsMain.Cells(targetRow, 27).NumberFormat = "# ##0,00"
                End If
                targetRow = targetRow + 1
            End If
        Next i
    End If

    If Not foundWorks Then
        MsgBox "Таблица 'Выполненные работы' не найдена!", vbExclamation, "Предупреждение"
    End If

    If Not foundMaterials Then
        MsgBox "Таблица 'Расходная накладная' не найдена!", vbExclamation, "Предупреждение"
    End If
    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при импорте данных: " & Err.Description & ". Импорт прерван.", vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_Import", "ImportDataToMain: " & Err.Description)
End Sub

' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' ImportSheet_UI
' Импортирует лист из report.xlsx по ГРЗ из ячейки B4 листа main
' --------------------------------------------------------------------------
Public Sub ImportSheet_UI()
    On Error GoTo ErrHandler

    Call ImportSheet(ThisWorkbook.Sheets("main").Range("B4").Value)

    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportSheet_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportSheet_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportByInput_UI
' Запрашивает ГРЗ через InputBox, вызывает ImportSheet
' --------------------------------------------------------------------------
Public Sub ImportByInput_UI()
    On Error GoTo ErrHandler

    Dim grz As String
    grz = InputBox("Введите ГРЗ для импорта:", "Импорт по ГРЗ")

    If grz = "" Then
        Exit Sub
    End If

    Call ImportSheet(grz)

    MsgBox "Импорт по ГРЗ '" & grz & "' выполнен.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportByInput_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportByInput_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' RenameSheets_UI
' Переименовывает листы в report.xlsx по ГРЗ
' --------------------------------------------------------------------------
Public Sub RenameSheets_UI()
    On Error GoTo ErrHandler

    Call Mod_SheetOps.RenameSheetsByGRZ

    MsgBox "Переименование листов выполнено.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в RenameSheets_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("RenameSheets_UI: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' ImportDataToMain_UI
' Переносит данные с активного листа в лист main
' --------------------------------------------------------------------------
Public Sub ImportDataToMain_UI()
    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Set wsSource = ActiveSheet

    If wsSource Is Nothing Then
        MsgBox "Нет активного листа!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    If wsSource.Name = "main" Then
        MsgBox "Активный лист не может быть main. Выберите другой лист.", vbExclamation, "Предупреждение"
        Exit Sub
    End If

    Call ImportDataToMain(wsSource)

    MsgBox "Данные с листа '" & wsSource.Name & "' перенесены в main.", vbInformation, "SysW"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка в ImportDataToMain_UI: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("ImportDataToMain_UI: " & Err.Description)
End Sub

' ============================================================
' ImportFromB2_UI
' Импорт данных на лист "мэйн" из листа {B2}M
' Если листа нет — копирует из report.xlsx
' Номер для поиска берётся из ячейки B2 листа "мэйн"
' ============================================================
Public Sub ImportFromB2_UI()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim grz As String
    Dim sheetName As String
    Dim wsSource As Worksheet
    Dim wbReport As Workbook
    Dim reportPath As String
    Dim ws As Worksheet
    Dim grzNumber As String

    ' Отключаем обновление экрана и события для производительности
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 1. Получаем лист "мэйн" и читаем B2
    Set wsMain = ThisWorkbook.Sheets("main")
    grz = Trim(CStr(wsMain.Range("B2").Value))

    ' 2. Проверяем, что B2 не пуст
    If grz = "" Or grz = "0" Then
        MsgBox "Ячейка B2 на листе 'main' пуста. Укажите номер заказа.", _
               vbExclamation, "Импорт ВХ"
        GoTo CleanUp
    End If

    ' 3. Формируем имя листа-источника
    sheetName = grz & "M"

    ' 4. Проверяем, существует ли лист {B2}M в текущей книге
    Set wsSource = Mod_Utils.GetSheetByName(ThisWorkbook, sheetName)

    If wsSource Is Nothing Then
        ' ---- Лист не существует — копируем из report.xlsx ----
        reportPath = ThisWorkbook.path & "\report.xlsx"

        ' Проверяем, существует ли файл report.xlsx
        If Not Mod_Utils.FileExists(reportPath) Then
            MsgBox "Файл report.xlsx не найден по пути:" & vbCrLf & reportPath, _
                   vbExclamation, "Импорт ВХ"
            GoTo CleanUp
        End If

        ' Открываем report.xlsx (ReadOnly) — ЕДИНСТВЕННЫЙ вызов Open
        Set wbReport = Workbooks.Open(reportPath, ReadOnly:=True)

        ' Извлекаем цифровой номер из ГРЗ для поиска листа
        grzNumber = Mod_SheetOps.ExtractNumberFromGRZ(grz)

        ' Ищем лист по номеру в уже открытой книге wbReport (без повторного Open)
        Set wsSource = Nothing
        If grzNumber <> "" Then
            For Each ws In wbReport.Sheets
                If InStr(1, ws.Name, grzNumber, vbTextCompare) > 0 Then
                    Set wsSource = ws
                    Exit For
                End If
            Next ws
        End If

        If wsSource Is Nothing Then
            MsgBox "Лист с номером '" & grz & "' не найден в файле report.xlsx.", _
                   vbExclamation, "Импорт ВХ"
            If Not wbReport Is Nothing Then
                wbReport.Close SaveChanges:=False
            End If
            GoTo CleanUp
        End If

        ' Копируем найденный лист в текущую книгу после листа "мэйн"
        wsSource.Copy After:=ThisWorkbook.Sheets("main")

        ' Закрываем report.xlsx
        wbReport.Close SaveChanges:=False
        Set wbReport = Nothing

        ' Переименовываем скопированный лист
        On Error Resume Next
        ActiveSheet.Name = sheetName
        If Err.Number <> 0 Then
            MsgBox "Не удалось переименовать лист в '" & sheetName & "': " & Err.Description, _
                   vbExclamation, "Импорт ВХ"
            On Error GoTo ErrHandler
            GoTo CleanUp
        End If
        On Error GoTo ErrHandler

        ' Получаем ссылку на новый лист
        Set wsSource = ActiveSheet
    End If

    ' 5. Вызываем ImportDataToMain для переноса данных
    Call ImportDataToMain(wsSource)

    ' 6. Заполняем шапку заказа из spisok и models
    If grz <> "" Then
        Call Mod_OrderHeader.FillHeaderFromOrder(grz)
    End If

    MsgBox "Импорт по номеру '" & grz & "' выполнен успешно.", _
           vbInformation, "Импорт ВХ"

CleanUp:
    ' Восстановление состояния приложения
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения при ошибке
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    ' Закрываем report.xlsx, если он ещё открыт
    If Not wbReport Is Nothing Then
        wbReport.Close SaveChanges:=False
    End If

    MsgBox "Ошибка при импорте ВХ: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_Import", "ImportFromB2_UI: " & Err.Description)
End Sub
