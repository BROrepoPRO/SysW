Attribute VB_Name = "Mod_Import"
Option Explicit
Option Private Module

' Флаг подавления MsgBox (используется при тестировании)
Public SilenceMsgBox As Boolean

' Предохранитель минимальной строки реальных данных (v1.0.16).
' Объявлен ЗДЕСЬ (до первого использования в ImportDataToMain): VBA не допускает
' обратную ссылку на модульную Const из процедур, расположенных в файле выше
' места объявления (иначе "Compile error: Variable not defined").
Private Const MIN_SPREAD_DATA_ROW As Long = 10

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

    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    Set wsSource = Mod_SheetOps.SearchSheetByGRZ(grz)
    If wsSource Is Nothing Then
        If Not SilenceMsgBox Then
            MsgBox "Лист с ГРЗ " & grz & " не найден!", vbExclamation, "Ошибка"
        End If
        Exit Sub
    End If

    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

    newName = Trim(wsMain.Range("B4").Value) & "M"
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
    If Not SilenceMsgBox Then
        MsgBox "Ошибка при импорте данных: " & Err.Description & ". Импорт прерван.", vbCritical, "Ошибка"
    End If
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
    ' Адаптивный маппинг (v1.0.16): динамические source-колонки и строка заголовков
    Dim lastCol As Long
    Dim hRow As Long
    Dim cName As Long, cQty As Long, cTotal As Long, cNum As Long
    Dim useAdaptive As Boolean

    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    ' Очистка диапазонов работ L:O и запчастей X:AB.
    ' O(15)/AB(28) — колонки подставленных модельных артикулов (проблема 2).
    Dim lastRowL As Long, lastRowO As Long, lastRowX As Long, lastRowAB As Long
    lastRowL = wsMain.Cells(wsMain.Rows.count, 12).End(xlUp).Row
    lastRowO = wsMain.Cells(wsMain.Rows.count, 15).End(xlUp).Row
    lastRowX = wsMain.Cells(wsMain.Rows.count, 24).End(xlUp).Row
    lastRowAB = wsMain.Cells(wsMain.Rows.count, 28).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowL, lastRowO, lastRowX, lastRowAB)
    If lastRow < 4 Then lastRow = 4

    wsMain.Range("L4:O" & lastRow).ClearContents
    wsMain.Range("X4:AB" & lastRow).ClearContents

    ' Провайдер и группа для глубокой подстановки модельных кодов.
    ' Получаются ОДИН раз на весь импорт; при Nothing либо пустой группе
    ' (или выключенном флаге) подстановка не выполняется.
    Dim matProv As IModelDataProvider
    Dim matGroup As String
    ' Активация глубокой подстановки модельных кодов при каждом импорте (задача 2 v1.0.7).
    Mod_Constants.ApplyMatLibSubstitution = True
    If Mod_Constants.ApplyMatLibSubstitution Then
        Call Mod_ModelDB.GetModelDataProvider(matProv)
        matGroup = Trim(CStr(wsMain.Range("B14").Value))
    End If

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
    ' Маппинг на main: D(4)→L(12), I(9)→M(13), M(13)→N(14)
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
        ' === Адаптивный маппинг работ (v1.0.16, Вариант A) ===
        hRow = FindHeaderRow(wsSource, cell, srcLastRow, "works")
        useAdaptive = False
        If hRow > 0 Then
            lastCol = wsSource.Cells(hRow, wsSource.Columns.count).End(xlToLeft).Column
            If lastCol < 1 Then lastCol = 1
            ' Предохранитель: строка заголовков расположена слишком высоко —
            ' признак искусственной минимальной фикстуры → fallback.
            If hRow + 2 >= MIN_SPREAD_DATA_ROW Then
                useAdaptive = BuildWorksSourceMap(wsSource, hRow, lastCol, cName, cQty, cTotal)
            End If
        End If

        If useAdaptive Then
            dataStartRow = hRow + 2
        Else
            ' Fallback: прежний позиционный маппинг работ 4/9/13
            cName = 4: cQty = 9: cTotal = 13
            ' Пропускаем пустые строки после названия таблицы
            dataStartRow = cell.Row + 1
            Do While dataStartRow <= srcLastRow
                If Trim(wsSource.Cells(dataStartRow, 4).Value) <> "" Then Exit Do
                dataStartRow = dataStartRow + 1
            Loop
            ' Пропускаем две строки заголовка (заголовок + подзаголовок с номерами)
            dataStartRow = dataStartRow + 2
        End If

        targetRow = 4 ' данные на main начинаем писать со строки 4
        For i = dataStartRow To srcLastRow
            If wsSource.Cells(i, cName).Value <> "" Then
                wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_NAME).Value = wsSource.Cells(i, cName).Value
                wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_QTY).Value = wsSource.Cells(i, cQty).Value
                ' Колонка «Всего» (N) необязательна: пишем/форматируем только при наличии роли
                If cTotal > 0 Then
                    wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_TOTAL).Value = wsSource.Cells(i, cTotal).Value
                End If
                ' Форматируем числовые колонки: убираем десятичные знаки для целых чисел
                If IsNumeric(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_QTY).Value) Then
                    If wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_QTY).Value = Int(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_QTY).Value) Then
                        wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_QTY).NumberFormat = "0"
                    End If
                End If
                If cTotal > 0 Then
                    If IsNumeric(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_TOTAL).Value) Then
                        wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_W_TOTAL).NumberFormat = "# ##0,00"
                    End If
                End If
                ' Глубокая подстановка модельного артикула работы (O), если флаг включён
                If Not matProv Is Nothing And matGroup <> "" Then
                    Call SubstituteWorkArticle(matProv, matGroup, wsMain, targetRow)
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
    ' Маппинг на main (v1.0.6, согласован с docs/table.md разд. 3.2):
    ' B(2)→X(24) № кат., C(3)→Y(25) Наименование, D(4)→Z(26) Кол-во,
    ' G(7)→AA(27) Всего. Подставленный модельный артикул — AB(28).
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
        ' === Адаптивный маппинг материалов (v1.0.16, Вариант A) ===
        hRow = FindHeaderRow(wsSource, cell, srcLastRow, "parts")
        useAdaptive = False
        If hRow > 0 Then
            lastCol = wsSource.Cells(hRow, wsSource.Columns.count).End(xlToLeft).Column
            If lastCol < 1 Then lastCol = 1
            ' Предохранитель: строка заголовков расположена слишком высоко —
            ' признак искусственной минимальной фикстуры → fallback.
            If hRow + 2 >= MIN_SPREAD_DATA_ROW Then
                useAdaptive = BuildPartsSourceMap(wsSource, hRow, lastCol, cNum, cName, cQty, cTotal)
            End If
        End If

        If useAdaptive Then
            dataStartRow = hRow + 2
        Else
            ' Fallback: прежний позиционный маппинг материалов 2/3/4/7
            cNum = 2: cName = 3: cQty = 4: cTotal = 7
            ' Пропускаем пустые строки после названия таблицы
            dataStartRow = cell.Row + 1
            Do While dataStartRow <= srcLastRow
                If Trim(wsSource.Cells(dataStartRow, 2).Value) <> "" Then Exit Do
                dataStartRow = dataStartRow + 1
            Loop
            ' Пропускаем две строки заголовка (заголовок + подзаголовок с номерами)
            dataStartRow = dataStartRow + 2
        End If

        targetRow = 4 ' данные на main начинаем писать со строки 4
        For i = dataStartRow To srcLastRow
            If wsSource.Cells(i, cNum).Value <> "" Then
                wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_NUM).Value = wsSource.Cells(i, cNum).Value
                wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_NAME).Value = wsSource.Cells(i, cName).Value
                wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_QTY).Value = wsSource.Cells(i, cQty).Value
                ' Колонка «Всего» (AA) необязательна: пишем/форматируем только при наличии роли
                If cTotal > 0 Then
                    wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_TOTAL).Value = wsSource.Cells(i, cTotal).Value
                End If
                ' Форматируем числовые колонки
                If IsNumeric(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_QTY).Value) Then
                    If wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_QTY).Value = Int(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_QTY).Value) Then
                        wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_QTY).NumberFormat = "0"
                    End If
                End If
                If cTotal > 0 Then
                    If IsNumeric(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_TOTAL).Value) Then
                        If wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_TOTAL).Value = Int(wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_TOTAL).Value) Then
                            wsMain.Cells(targetRow, Mod_Constants.MAIN_COL_P_TOTAL).NumberFormat = "# ##0,00"
                        End If
                    End If
                End If
                ' Глубокая подстановка модельного артикула запчасти (AB), если флаг включён
                If Not matProv Is Nothing And matGroup <> "" Then
                    Call SubstitutePartArticle(matProv, matGroup, wsMain, targetRow)
                End If
                targetRow = targetRow + 1
            End If
        Next i
    End If

    If Not foundWorks Then
        If Not SilenceMsgBox Then
            MsgBox "Таблица 'Выполненные работы' не найдена!", vbExclamation, "Предупреждение"
        End If
    End If

    If Not foundMaterials Then
        If Not SilenceMsgBox Then
            MsgBox "Таблица 'Расходная накладная' не найдена!", vbExclamation, "Предупреждение"
        End If
    End If
    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not SilenceMsgBox Then
        MsgBox "Ошибка при импорте данных: " & Err.Description & ". Импорт прерван.", vbCritical, "Ошибка"
    End If
    Call Mod_Logger.WriteLog("Mod_Import", "ImportDataToMain: " & Err.Description)
End Sub

' ============================================================
' ГЛУБОКАЯ ПОДСТАНОВКА МОДЕЛЬНЫХ КОДОВ (проблема 2)
' Вызывается из ImportDataToMain при включённом флаге
' Mod_Constants.ApplyMatLibSubstitution.
' Ключ поиска — matlib_entries.group_name + entry_code; подставляется
' target_code ТОЛЬКО при наличии точного совпадения и подходящем target_type.
' При отсутствии совпадения целевая ячейка остаётся пустой/без изменений.
' ============================================================

' --------------------------------------------------------------------------
' SubstituteWorkArticle
' Подставляет модельный артикул работы в колонку O(15).
' Ключ поиска — наименование работы L(12) (target_type = 'mod_work').
' Наименование в L сохраняется без изменений.
' --------------------------------------------------------------------------
Private Sub SubstituteWorkArticle(ByVal prov As IModelDataProvider, _
                                  ByVal groupName As String, _
                                  ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    ' Работы подбираются ТОЛЬКО по наименованию L(12); приоритет не меняется.
    Dim key As String
    key = Trim(CStr(ws.Cells(rowNum, 12).Value))   ' L — наименование работы
    If key = "" Then Exit Sub   ' пустой ключ — пропуск, O(15) не заполняется

    Dim entries As Collection
    Set entries = prov.GetMatLibEntries(groupName, key)

    ' Берём первую запись нужного типа MOD_WORK, пропуская записи другого типа.
    Dim idx As Long
    idx = FindFirstMatLibIndex(entries, "MOD_WORK")
    If idx > 0 Then
        Dim arr As Variant
        arr = entries(idx)   ' [target_type, target_code, target_name, coefficient]
        ws.Cells(rowNum, 15).Value = CStr(arr(1))   ' O — модельный артикул
    End If
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_Import", "SubstituteWorkArticle: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' SubstitutePartArticle
' Подставляет модельный артикул запчасти в колонку AB(28).
' v1.0.7 (согласованное бизнес-правило): поиск сначала по № кат. X(24);
' если № кат. пуст ИЛИ по нему нет совпадения типа MOD_PART — fallback
' по наименованию Y(25). При пустых обоих ключах ячейка AB(28) не заполняется.
' Наименование в Y сохраняется без изменений.
' --------------------------------------------------------------------------
Private Sub SubstitutePartArticle(ByVal prov As IModelDataProvider, _
                                  ByVal groupName As String, _
                                  ByVal ws As Worksheet, ByVal rowNum As Long)
    On Error GoTo ErrHandler

    Dim key As String
    Dim entries As Collection
    Dim idx As Long
    Dim arr As Variant

    ' 1. Поиск по № кат. X(24) — приоритетный ключ
    key = Trim(CStr(ws.Cells(rowNum, 24).Value))
    If key <> "" Then
        Set entries = prov.GetMatLibEntries(groupName, key)
        idx = FindFirstMatLibIndex(entries, "MOD_PART")
        If idx > 0 Then
            arr = entries(idx)
            ws.Cells(rowNum, 28).Value = CStr(arr(1))   ' AB — модельный артикул
            Exit Sub
        End If
    End If

    ' 2. Fallback: поиск по наименованию Y(25)
    key = Trim(CStr(ws.Cells(rowNum, 25).Value))
    If key = "" Then Exit Sub   ' пустой ключ — пропуск, AB(28) не заполняется

    Set entries = prov.GetMatLibEntries(groupName, key)
    idx = FindFirstMatLibIndex(entries, "MOD_PART")
    If idx > 0 Then
        arr = entries(idx)
        ws.Cells(rowNum, 28).Value = CStr(arr(1))   ' AB — модельный артикул
    End If
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_Import", "SubstitutePartArticle: " & Err.Description)
End Sub

' --------------------------------------------------------------------------
' FindFirstMatLibIndex
' Возвращает индекс (1-based) первой записи коллекции GetMatLibEntries,
' чей target_type равен wantedType (без учёта регистра); пропускает записи
' других типов. Возвращает 0, если запись нужного типа не найдена.
' Запись — массив [target_type, target_code, target_name, coefficient].
' --------------------------------------------------------------------------
Private Function FindFirstMatLibIndex(ByVal entries As Collection, _
                                      ByVal wantedType As String) As Long
    Dim i As Long
    Dim arr As Variant
    FindFirstMatLibIndex = 0
    For i = 1 To entries.Count
        arr = entries(i)
        If UCase$(CStr(arr(0))) = UCase$(wantedType) Then
            FindFirstMatLibIndex = i
            Exit Function
        End If
    Next i
End Function

' ============================================================
' АДАПТИВНЫЙ ПАРСИНГ ЗАГОЛОВКОВ (v1.0.16, Вариант A)
' Позволяет корректно импортировать данные из заказ-нарядов, где таблицы
' "Выполненные работы" / "Расходная накладная" содержат заголовки, чьи
' колонки могут отличаться по позициям от исторических фиксированных
' B..M / A..H. Роли колонок определяются по нормализованным заголовкам
' (без учёта регистра, пробелов и знаков препинания).
' ============================================================

' Примечание: предохранитель MIN_SPREAD_DATA_ROW объявлен в начале модуля (v1.0.16).
' Если найденная строка заголовков + 2 < MIN_SPREAD_DATA_ROW, импорт
' маршрутизируется в fallback (позиционный маппинг), чтобы искусственные
' минимальные фикстуры продолжали работать на старых колонках.

' --------------------------------------------------------------------------
' NormHeader
' Нормализует текст заголовка: UPPER, удаляет пробелы и знаки препинания.
' Возвращает строку из букв/цифр (кириллица и латиница сохраняются).
' --------------------------------------------------------------------------
Private Function NormHeader(ByVal txt As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim res As String
    s = UCase$(Trim$(CStr(txt)))
    res = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "A" To "Z", "0" To "9", "А" To "Я"
                res = res & ch
            Case Else
                ' пробел, пунктуация, «№» и пр. — пропускаем
        End Select
    Next i
    NormHeader = res
End Function

' --------------------------------------------------------------------------
' ResolveRole
' Определяет роль колонки по нормализованному заголовку и типу таблицы.
' tableKind: "works" (работы) или "parts" (материалы).
' Возвращает роль-строку либо пустую строку, если роль не распознана.
' Дизaмбигуация: точное совпадение приоритетнее частичного; разделение
' "Кол.оп."(работы) vs "Кол-во"(материалы) — по типу таблицы и маске.
' --------------------------------------------------------------------------
Private Function ResolveRole(ByVal norm As String, ByVal tableKind As String) As String
    Dim isWorks As Boolean
    isWorks = (tableKind = "works")
    ResolveRole = ""

    If norm = "" Then Exit Function

    ' Номер строки ("№"/"N") игнорируется для обоих типов таблиц
    If norm = "N" Then Exit Function

    If isWorks Then
        ' --- Работы ---
        If norm = "НАИМЕНОВАНИЕ" Or norm = "НАИМЕНОВАНИЕРАБОТ" Or norm = "НАИМ" Then
            ResolveRole = "wName"
        ElseIf norm = "КОЛОП" Or norm = "КОЛВООПЕР" Or norm = "КОЛИЧЕСТВООПЕРАЦИЙ" Or norm = "КОЛ" Then
            ResolveRole = "wQty"
        ElseIf norm = "ВСЕГО" Or norm = "СУММА" Or norm = "ИТОГО" Then
            ResolveRole = "wTotal"
        ElseIf norm = "ВТЧНДС" Or norm = "НДС" Then
            ResolveRole = "wNds"
        ElseIf norm = "ЦЕНА" Then
            ResolveRole = "wPrice"
        ElseIf norm = "НОРМА" Then
            ResolveRole = "wNorm"
        ElseIf norm = "НЧ" Or norm = "НОРМОЧАСЫ" Then
            ResolveRole = "wHours"
        End If
    Else
        ' --- Материалы ---
        ' № кат. — роль pNum требует наличия «КАТ» в нормализованном заголовке
        If norm = "КАТ" Or norm = "КАТАЛОЖНЫЙНОМЕР" Or norm = "НОМЕРКАТАЛОГА" Then
            ResolveRole = "pNum"
        ElseIf norm = "НАИМЕНОВАНИЕ" Or norm = "НАИМЕНОВАНИЕЗАПЧАСТИ" Or norm = "НАИМ" Then
            ResolveRole = "pName"
        ElseIf norm = "КОЛВО" Or norm = "КОЛИЧЕСТВО" Or norm = "КОЛ" Then
            ResolveRole = "pQty"
        ElseIf norm = "ЕДИЗМ" Or norm = "ЕД" Then
            ResolveRole = "pUnit"
        ElseIf norm = "ЦЕНА" Then
            ResolveRole = "pPrice"
        ElseIf norm = "ВСЕГО" Or norm = "СУММА" Then
            ResolveRole = "pTotal"
        ElseIf norm = "ВТЧНДС" Or norm = "НДС" Then
            ResolveRole = "pNds"
        End If
    End If
End Function

' --------------------------------------------------------------------------
' FindHeaderRow
' Ищет первую "богатую" строку заголовков ниже titleCell: строка считается
' заголовком, если в ней распознано >=3 ролей. Возвращает номер строки или 0.
'
' v1.0.16 fix: ширина сканирования lastCol берётся как МАКСИМУМ использованных
' колонок по области [titleCell.Row .. maxRow]. Нельзя ориентироваться только
' на строку titleCell: титул таблицы («Выполненные работы…»/«Расходная
' накладная…») занимает лишь пару колонок (обычно до C), тогда как реальные
' заголовки столбцов (№ кат. / Кол-во / Всего) уходят далеко вправо. Раньше
' скан ограничивался ~3 колонками, не набирал >=3 ролей и сваливался в fallback
' (что вызывало сдвиг запчастей).
' --------------------------------------------------------------------------
Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal titleCell As Range, _
                               ByVal maxRow As Long, ByVal tableKind As String) As Long
    Dim r As Long
    Dim c As Long
    Dim roleCount As Long
    Dim norm As String
    Dim lastCol As Long
    Dim lc As Long
    FindHeaderRow = 0

    ' Проход 1: максимальная правая использованная колонка в области таблицы.
    lastCol = 1
    For r = titleCell.Row To maxRow
        lc = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        If lc > lastCol Then lastCol = lc
    Next r
    If lastCol < 1 Then lastCol = 1

    ' Проход 2: поиск строки заголовков.
    For r = titleCell.Row + 1 To maxRow
        roleCount = 0
        For c = 1 To lastCol
            norm = NormHeader(ws.Cells(r, c).Value)
            If norm <> "" Then
                If ResolveRole(norm, tableKind) <> "" Then roleCount = roleCount + 1
            End If
        Next c
        If roleCount >= 3 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

' --------------------------------------------------------------------------
' BuildWorksSourceMap
' Строит маппинг ролей работ на source-колонки в строке заголовков headerRow.
' Возвращает True, если определены обязательные роли (wName, wQty);
' колонка «Всего» (wTotal) необязательна.
' --------------------------------------------------------------------------
Private Function BuildWorksSourceMap(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                     ByVal lastCol As Long, ByRef cName As Long, _
                                     ByRef cQty As Long, ByRef cTotal As Long) As Boolean
    Dim c As Long
    Dim role As String
    cName = 0: cQty = 0: cTotal = 0
    For c = 1 To lastCol
        role = ResolveRole(NormHeader(ws.Cells(headerRow, c).Value), "works")
        Select Case role
            Case "wName": cName = c
            Case "wQty": cQty = c
            Case "wTotal": cTotal = c
        End Select
    Next c
    ' Колонка «Всего» (wTotal) необязательна: адаптивный маппинг активен,
    ' если найдены имя и количество. При отсутствии «Всего» cTotal=0 и
    ' запись N пропускается (см. цикл импорта работ).
    BuildWorksSourceMap = (cName > 0 And cQty > 0)
End Function

' --------------------------------------------------------------------------
' BuildPartsSourceMap
' Строит маппинг ролей материалов на source-колонки.
' Возвращает True, если определены обязательные роли (pNum, pName, pQty);
' колонка «Всего» (pTotal) необязательна.
' --------------------------------------------------------------------------
Private Function BuildPartsSourceMap(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                     ByVal lastCol As Long, ByRef cNum As Long, _
                                     ByRef cName As Long, ByRef cQty As Long, _
                                     ByRef cTotal As Long) As Boolean
    Dim c As Long
    Dim role As String
    cNum = 0: cName = 0: cQty = 0: cTotal = 0
    For c = 1 To lastCol
        role = ResolveRole(NormHeader(ws.Cells(headerRow, c).Value), "parts")
        Select Case role
            Case "pNum": cNum = c
            Case "pName": cName = c
            Case "pQty": cQty = c
            Case "pTotal": cTotal = c
        End Select
    Next c
    ' Колонка «Всего» (pTotal) необязательна: адаптивный маппинг активен,
    ' если найдены № кат., имя и количество. При отсутствии «Всего» cTotal=0
    ' и запись AA пропускается (см. цикл импорта материалов). Это устраняет
    ' сдвиг ЗЧ, когда в таблице «Расходная накладная» нет колонки «Всего».
    BuildPartsSourceMap = (cNum > 0 And cName > 0 And cQty > 0)
End Function

' ============================================================
' _UI-ПРОЦЕДУРЫ (обёртки с пользовательским вводом/выводом)
' ============================================================

' --------------------------------------------------------------------------
' ImportDataToMain_UI
' Переносит данные с активного листа в лист main
' --------------------------------------------------------------------------
Public Sub ImportDataToMain_UI()
    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Set wsSource = ActiveSheet

    If wsSource Is Nothing Then
        If Not SilenceMsgBox Then
            MsgBox "Нет активного листа!", vbExclamation, "Ошибка"
        End If
        Exit Sub
    End If

    If wsSource.Name = Mod_Constants.SHEET_MAIN Then
        If Not SilenceMsgBox Then
            MsgBox "Активный лист не может быть main. Выберите другой лист.", vbExclamation, "Предупреждение"
        End If
        Exit Sub
    End If

    Call ImportDataToMain(wsSource)

    If Not SilenceMsgBox Then
        MsgBox "Данные с листа '" & wsSource.Name & "' перенесены в main.", vbInformation, "SysW"
    End If
    Exit Sub

ErrHandler:
    If Not SilenceMsgBox Then
        MsgBox "Ошибка в ImportDataToMain_UI: " & Err.Description, vbCritical, "Ошибка"
    End If
    Call Mod_Utils.WriteLog("ImportDataToMain_UI: " & Err.Description)
End Sub

' ============================================================
' ImportFromB2_UI
' Импорт данных на лист "мэйн" из листа {B4}M
' Если листа нет — копирует из report.xlsx
' Номер для поиска берётся из ячейки B4 листа "мэйн"
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

    ' 1. Получаем лист "мэйн" и читаем B4
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)
    grz = Trim(CStr(wsMain.Range("B4").Value))

    ' 2. Проверяем, что B4 не пуст
    If grz = "" Or grz = "0" Then
        If Not SilenceMsgBox Then
            MsgBox "Ячейка B4 на листе 'main' пуста. Укажите номер заказа.", _
                   vbExclamation, "Импорт ВХ"
        End If
        GoTo CleanUp
    End If

    ' 3. Формируем имя листа-источника
    sheetName = grz & "M"

    ' 4. Проверяем, существует ли лист {B4}M в текущей книге
    Set wsSource = Mod_Utils.GetSheetByName(ThisWorkbook, sheetName)

    If wsSource Is Nothing Then
        ' ---- Лист не существует — копируем из report.xlsx ----
        reportPath = ThisWorkbook.path & "\report.xlsx"

        ' Проверяем, существует ли файл report.xlsx
        If Not Mod_Utils.FileExists(reportPath) Then
            If Not SilenceMsgBox Then
                MsgBox "Файл report.xlsx не найден по пути:" & vbCrLf & reportPath, _
                       vbExclamation, "Импорт ВХ"
            End If
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
            If Not SilenceMsgBox Then
                MsgBox "Лист с номером '" & grz & "' не найден в файле report.xlsx.", _
                       vbExclamation, "Импорт ВХ"
            End If
            If Not wbReport Is Nothing Then
                wbReport.Close SaveChanges:=False
            End If
            GoTo CleanUp
        End If

        ' Копируем найденный лист в текущую книгу после листа "мэйн"
        wsSource.Copy After:=ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

        ' Закрываем report.xlsx
        wbReport.Close SaveChanges:=False
        Set wbReport = Nothing

        ' Переименовываем скопированный лист
        On Error Resume Next
        ActiveSheet.Name = sheetName
        If Err.Number <> 0 Then
            If Not SilenceMsgBox Then
                MsgBox "Не удалось переименовать лист в '" & sheetName & "': " & Err.Description, _
                       vbExclamation, "Импорт ВХ"
            End If
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

    If Not SilenceMsgBox Then
        MsgBox "Импорт по номеру '" & grz & "' выполнен успешно.", _
               vbInformation, "Импорт ВХ"
    End If

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

    If Not SilenceMsgBox Then
        MsgBox "Ошибка при импорте ВХ: " & Err.Description, vbCritical, "Ошибка"
    End If
    Call Mod_Logger.WriteLog("Mod_Import", "ImportFromB2_UI: " & Err.Description)
End Sub
