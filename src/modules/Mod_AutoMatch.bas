Attribute VB_Name = "Mod_AutoMatch"
Option Explicit
Option Private Module

' ============================================================
' Модуль: Mod_AutoMatch
' Назначение: Автоподбор работ и запчастей из тождеств работ и запчастей
' Вызывается из кнопок АВТО РАБ и АВТО ЗЧ на листе main
' ============================================================

' ============================================================
' Константы колонок листа main
' ============================================================

' --- Работы (D:K) ---
Private Const MAIN_W_ARTICLE As Long = 5     ' E — Артикул
Private Const MAIN_W_NAME As Long = 6        ' F — Наименование
Private Const MAIN_W_NORMHOURS As Long = 7   ' G — Кол-во н/ч
Private Const MAIN_W_QTY As Long = 8         ' H — Кол-во оп
Private Const MAIN_W_PRICE As Long = 9       ' I — Цена н/ч
Private Const MAIN_W_SUM As Long = 10        ' J — Сумма (формула)

' --- Входящие работы (L:N) ---
Private Const MAIN_W_IN_NAME As Long = 12    ' L — Наименование (входящее)
Private Const MAIN_W_IN_QTY As Long = 13     ' M — Кол. оп.
Private Const MAIN_W_IN_TOTAL As Long = 14   ' N — Всего

' --- Запчасти (P:W) ---
Private Const MAIN_P_ARTICLE As Long = 17    ' Q — Артикул
Private Const MAIN_P_NAME As Long = 18       ' R — Наименование
Private Const MAIN_P_QTY As Long = 20        ' T — Кол-во
Private Const MAIN_P_PRICE As Long = 21      ' U — Цена
Private Const MAIN_P_SUM As Long = 22        ' V — Сумма (формула)

' --- Входящие запчасти (X:AA) ---
Private Const MAIN_P_IN_CATNUM As Long = 24  ' X — № кат. (входящий)
Private Const MAIN_P_IN_NAME As Long = 25    ' Y — Наименование
Private Const MAIN_P_IN_QTY As Long = 26     ' Z — Кол-во
Private Const MAIN_P_IN_TOTAL As Long = 27   ' AA — Всего

' ============================================================
' Вспомогательные функции
' ============================================================

' --------------------------------------------------------------------------
' GetGroupName
' Возвращает название группы из ячейки B14 листа main
' --------------------------------------------------------------------------
Private Function GetGroupName() As String
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    GetGroupName = Trim(CStr(wsMain.Range("B14").Value))
    Exit Function

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "GetGroupName: Ошибка — " & Err.Description)
    GetGroupName = ""
End Function

' --------------------------------------------------------------------------
' HighlightNotFound
' Подсвечивает ячейку жёлтым и помечает "НЕ НАЙДЕНО"
' --------------------------------------------------------------------------
Private Sub HighlightNotFound(ByVal cell As Range)
    cell.Interior.Color = RGB(255, 255, 0) ' Жёлтый
    cell.Font.Color = RGB(255, 0, 0)       ' Красный текст
End Sub

' --------------------------------------------------------------------------
' ClearHighlight
' Очищает подсветку и форматирование в диапазоне
' --------------------------------------------------------------------------
Private Sub ClearHighlight(ByVal rng As Range)
    rng.Interior.ColorIndex = xlNone
    rng.Font.ColorIndex = xlAutomatic
End Sub

' --------------------------------------------------------------------------
' IsAllFound
' Проверяет, все ли строки в диапазоне найдены (нет подсветки "НЕ НАЙДЕНО")
' --------------------------------------------------------------------------
Private Function IsAllFound(ByVal ws As Worksheet, ByVal checkCol As Long, _
                            ByVal startRow As Long, ByVal endRow As Long) As Boolean
    Dim i As Long
    Dim cellVal As String

    For i = startRow To endRow
        cellVal = Trim(CStr(ws.Cells(i, checkCol).Value))
        If cellVal = "НЕ НАЙДЕНО" Then
            IsAllFound = False
            Exit Function
        End If
    Next i

    IsAllFound = True
End Function

' ============================================================
' AutoMatchWorks
' АВТО РАБ: автоподбор работ из тождеств работ
' ============================================================
Public Sub AutoMatchWorks()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim groupName As String
    Dim identities As Collection
    Dim identity As WorkIdentity
    Dim lastRow As Long
    Dim i As Long
    Dim inName As String
    Dim found As Boolean
    Dim matchCount As Long
    Dim notFoundCount As Long
    Dim priceNH As Double

    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: START")

    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    ' 1. Получаем группу из B14
    groupName = GetGroupName()
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: groupName=" & groupName)
    If groupName = "" Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Не указана группа в ячейке B14.", vbExclamation, "АВТО РАБ"
        End If
        Exit Sub
    End If

    ' 2. Читаем тождества работ
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: BEFORE GetWorkIdentities")
    On Error Resume Next
    Set identities = Mod_ModelDB.GetWorkIdentities(groupName)
    If Err.Number <> 0 Then
        Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: GetWorkIdentities error: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo ErrHandler
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: identities.Count=" & CStr(identities.Count))
    If identities Is Nothing Or identities.Count = 0 Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Не найдены тождества работ для группы " & groupName & ".", _
                   vbExclamation, "АВТО РАБ"
        End If
        Exit Sub
    End If

    ' 3. Получаем цену н/ч из B13
    priceNH = Val(wsMain.Range("B13").Value)
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: priceNH=" & CStr(priceNH))

    ' 4. Определяем последнюю строку входящих работ (столбец L)
    lastRow = wsMain.Cells(wsMain.Rows.Count, MAIN_W_IN_NAME).End(xlUp).Row
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: lastRow=" & CStr(lastRow))
    If lastRow < 4 Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Нет входящих работ для автоподбора.", vbInformation, "АВТО РАБ"
        End If
        Exit Sub
    End If

    ' 5. Очищаем предыдущие результаты в колонках E-K (начиная с 4-й строки)
    For i = 4 To lastRow
        ClearHighlight wsMain.Range(wsMain.Cells(i, MAIN_W_ARTICLE), _
                                    wsMain.Cells(i, MAIN_W_SUM))
        wsMain.Cells(i, MAIN_W_ARTICLE).Resize(1, 6).ClearContents
    Next i

    ' 6. Для каждой строки входящих работ ищем совпадение
    matchCount = 0
    notFoundCount = 0

    For i = 4 To lastRow
        inName = Trim(CStr(wsMain.Cells(i, MAIN_W_IN_NAME).Value))

        ' Пропускаем пустые строки
        If inName = "" Then GoTo ContinueWork

        found = False

        ' Ищем в коллекции тождеств по InName
        For Each identity In identities
            If UCase$(Trim$(identity.InName)) = UCase$(inName) Then
                ' Найдено — заполняем модельные колонки
                wsMain.Cells(i, MAIN_W_ARTICLE).Value = identity.OutArticle   ' E ← B
                wsMain.Cells(i, MAIN_W_NAME).Value = identity.OutName         ' F ← C
                wsMain.Cells(i, MAIN_W_NORMHOURS).Value = identity.NormHours  ' G ← D
                wsMain.Cells(i, MAIN_W_QTY).Value = identity.QtyZN            ' H ← G
                wsMain.Cells(i, MAIN_W_PRICE).Value = priceNH                 ' I ← B13
                Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: writing formula at row " & CStr(i))
                ' J — Сумма = ОКРУГЛ(G*H*I;2)
                wsMain.Cells(i, MAIN_W_SUM).FormulaLocal = _
                    "=ROUND(G" & i & "*H" & i & "*I" & i & ";2)"

                matchCount = matchCount + 1
                found = True
                Exit For
            End If
        Next identity

        If Not found Then
            ' Не найдено по тождествам — поиск по локальному листу работ группы
            If Mod_Constants.SilenceMsgBox Then
                ' Тестовый режим (без диалогов): пометить как не найденное
                HighlightNotFound wsMain.Cells(i, MAIN_W_IN_NAME)
                wsMain.Cells(i, MAIN_W_IN_NAME).Value = "НЕ НАЙДЕНО"
                notFoundCount = notFoundCount + 1
            Else
                Dim lw As WorkIdentity
                Set lw = Mod_ModelDB.ReadLocalWorkByName(groupName, inName)
                If lw Is Nothing Then
                    MsgBox "Работа не найдена по тождествам и в списке работ группы.", _
                           vbExclamation, "АВТО РАБ"
                    HighlightNotFound wsMain.Cells(i, MAIN_W_IN_NAME)
                    wsMain.Cells(i, MAIN_W_IN_NAME).Value = "НЕ НАЙДЕНО"
                    notFoundCount = notFoundCount + 1
                Else
                    wsMain.Cells(i, MAIN_W_ARTICLE).Value = lw.OutArticle
                    wsMain.Cells(i, MAIN_W_NAME).Value = lw.OutName
                    wsMain.Cells(i, MAIN_W_NORMHOURS).Value = lw.NormHours
                    wsMain.Cells(i, MAIN_W_QTY).Value = lw.QtyZN
                    wsMain.Cells(i, MAIN_W_PRICE).Value = priceNH
                    wsMain.Cells(i, MAIN_W_SUM).FormulaLocal = _
                        "=ROUND(G" & i & "*H" & i & "*I" & i & ";2)"
                    matchCount = matchCount + 1

                    Dim askWId As VbMsgBoxResult
                    askWId = MsgBox("Создать тождество для этой работы?", _
                                    vbYesNo, "АВТО РАБ")
                    If askWId = vbYes Then
                        Dim newW As WorkIdentity
                        Set newW = New WorkIdentity
                        newW.OutArticle = lw.OutArticle
                        newW.OutName = lw.OutName
                        newW.NormHours = lw.NormHours
                        newW.QtyZN = Val(wsMain.Cells(i, MAIN_W_IN_QTY).Value)
                        newW.InName = inName
                        If Not Mod_ModelDB.AppendWorkIdentity(groupName, newW) Then
                            MsgBox "Не удалось сохранить тождество.", _
                                   vbExclamation, "АВТО РАБ"
                        End If
                    End If
                End If
            End If
        End If

ContinueWork:
    Next i

    ' 7. Итоговое сообщение
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: Найдено " & matchCount & _
                             ", не найдено " & notFoundCount)

    If notFoundCount = 0 Then
        ' Проверяем, все ли запчасти тоже найдены (если есть)
        If IsAllFound(wsMain, MAIN_P_IN_CATNUM, 4, lastRow) Then
            If Not Mod_Constants.SilenceMsgBox Then
                MsgBox "ЗН заполнен!", vbInformation, "АВТО РАБ"
            End If
        Else
            If Not Mod_Constants.SilenceMsgBox Then
                MsgBox "Автоподбор работ завершён. Найдено: " & matchCount & _
                       ". Осталось подобрать запчасти.", vbInformation, "АВТО РАБ"
            End If
        End If
    Else
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Автоподбор работ завершён." & vbCrLf & _
                   "Найдено: " & matchCount & vbCrLf & _
                   "Не найдено: " & notFoundCount & vbCrLf & _
                   "Ненайденные строки подсвечены жёлтым.", _
                   vbExclamation, "АВТО РАБ"
        End If
    End If

    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchWorks: Ошибка — " & Err.Description)
    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox "Ошибка при автоподборе работ: " & Err.Description, vbCritical, "АВТО РАБ"
    End If
End Sub

' ============================================================
' AutoMatchParts
' АВТО ЗЧ: автоподбор запчастей из тождеств запчастей
' ============================================================
Public Sub AutoMatchParts()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim groupName As String
    Dim identities As Collection
    Dim identity As PartIdentity
    Dim lastRow As Long
    Dim i As Long
    Dim inCatNum As String
    Dim inName As String
    Dim found As Boolean
    Dim matchCount As Long
    Dim notFoundCount As Long

    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: START")

    Set wsMain = ThisWorkbook.Sheets(Mod_Constants.SHEET_MAIN)

    ' 1. Получаем группу из B14
    groupName = GetGroupName()
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: groupName=" & groupName)
    If groupName = "" Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Не указана группа в ячейке B14.", vbExclamation, "АВТО ЗЧ"
        End If
        Exit Sub
    End If

    ' 2. Читаем тождества запчастей
    Set identities = Mod_ModelDB.GetPartIdentities(groupName)
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: identities.Count=" & CStr(identities.Count))
    If identities Is Nothing Or identities.Count = 0 Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Не найдены тождества запчастей для группы " & groupName & ".", _
                   vbExclamation, "АВТО ЗЧ"
        End If
        Exit Sub
    End If

    ' 3. Определяем последнюю строку входящих запчастей (столбец X)
    lastRow = wsMain.Cells(wsMain.Rows.Count, MAIN_P_IN_CATNUM).End(xlUp).Row
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: lastRow=" & CStr(lastRow))
    If lastRow < 4 Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Нет входящих запчастей для автоподбора.", vbInformation, "АВТО ЗЧ"
        End If
        Exit Sub
    End If

    ' 4. Очищаем предыдущие результаты в колонках Q-V (начиная с 4-й строки)
    For i = 4 To lastRow
        ClearHighlight wsMain.Range(wsMain.Cells(i, MAIN_P_ARTICLE), _
                                    wsMain.Cells(i, MAIN_P_SUM))
        wsMain.Cells(i, MAIN_P_ARTICLE).Resize(1, 5).ClearContents
    Next i

    ' 5. Для каждой строки входящих запчастей ищем совпадение
    matchCount = 0
    notFoundCount = 0

    For i = 4 To lastRow
        inCatNum = Trim(CStr(wsMain.Cells(i, MAIN_P_IN_CATNUM).Value))
        inName = Trim(CStr(wsMain.Cells(i, MAIN_P_IN_NAME).Value))

        ' Пропускаем строки без ключа (пусты и № кат., и наименование)
        If inCatNum = "" And inName = "" Then GoTo ContinuePart

        found = False

        ' Ищем в коллекции тождеств по InCatNum
        For Each identity In identities
            If UCase$(Trim$(identity.InCatNum)) = UCase$(inCatNum) Then
                ' Найдено — заполняем модельные колонки
                wsMain.Cells(i, MAIN_P_ARTICLE).Value = identity.OutArticle   ' Q ← B
                wsMain.Cells(i, MAIN_P_NAME).Value = identity.OutName         ' R ← C
                wsMain.Cells(i, MAIN_P_QTY).Value = identity.QtyZN            ' T ← G
                wsMain.Cells(i, MAIN_P_PRICE).Value = identity.Price          ' U ← F
                Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: writing formula at row " & CStr(i))
                ' V — Сумма = ОКРУГЛ(T*U;2)
                wsMain.Cells(i, MAIN_P_SUM).FormulaLocal = _
                    "=ROUND(T" & i & "*U" & i & ";2)"

                matchCount = matchCount + 1
                found = True
                Exit For
            End If
        Next identity

        If Not found Then
            ' Не найдено по тождествам — предлагаем поиск по общей базе з/ч
            If Mod_Constants.SilenceMsgBox Then
                ' Тестовый режим (без диалогов): пометить как не найденное
                HighlightNotFound wsMain.Cells(i, MAIN_P_IN_CATNUM)
                wsMain.Cells(i, MAIN_P_IN_CATNUM).Value = "НЕ НАЙДЕНО"
                notFoundCount = notFoundCount + 1
            Else
                Dim askBase As VbMsgBoxResult
                askBase = MsgBox("По тождествам не найдено. Искать в общей базе з/ч?", _
                                 vbYesNo, "АВТО ЗЧ")
                If askBase = vbNo Then
                    HighlightNotFound wsMain.Cells(i, MAIN_P_IN_CATNUM)
                    wsMain.Cells(i, MAIN_P_IN_CATNUM).Value = "НЕ НАЙДЕНО"
                    notFoundCount = notFoundCount + 1
                Else
                    Dim gPart As PartIdentity
                    Set gPart = Mod_ModelDB.ReadGlobalPartByKey(inCatNum, inName)
                    If gPart Is Nothing Then
                        MsgBox "Запчасть не найдена в общей базе з/ч.", _
                               vbExclamation, "АВТО ЗЧ"
                        HighlightNotFound wsMain.Cells(i, MAIN_P_IN_CATNUM)
                        wsMain.Cells(i, MAIN_P_IN_CATNUM).Value = "НЕ НАЙДЕНО"
                        notFoundCount = notFoundCount + 1
                    Else
                        wsMain.Cells(i, MAIN_P_ARTICLE).Value = gPart.OutArticle
                        wsMain.Cells(i, MAIN_P_NAME).Value = gPart.OutName
                        wsMain.Cells(i, MAIN_P_QTY).Value = gPart.QtyZN
                        wsMain.Cells(i, MAIN_P_PRICE).Value = gPart.Price
                        wsMain.Cells(i, MAIN_P_SUM).FormulaLocal = _
                            "=ROUND(T" & i & "*U" & i & ";2)"
                        matchCount = matchCount + 1

                        Dim askId As VbMsgBoxResult
                        askId = MsgBox("Создать тождество для этой запчасти?", _
                                       vbYesNo, "АВТО ЗЧ")
                        If askId = vbYes Then
                            Dim newId As PartIdentity
                            Set newId = New PartIdentity
                            newId.OutArticle = gPart.OutArticle
                            newId.OutName = gPart.OutName
                            newId.Price = gPart.Price
                            newId.QtyZN = Val(wsMain.Cells(i, MAIN_P_IN_QTY).Value)
                            newId.InCatNum = inCatNum
                            newId.InName = inName
                            If Not Mod_ModelDB.AppendPartIdentity(groupName, newId) Then
                                MsgBox "Не удалось сохранить тождество.", _
                                       vbExclamation, "АВТО ЗЧ"
                            End If
                        End If
                    End If
                End If
            End If
        End If

ContinuePart:
    Next i

    ' 6. Итоговое сообщение
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: Найдено " & matchCount & _
                             ", не найдено " & notFoundCount)

    If notFoundCount = 0 Then
        ' Проверяем, все ли работы тоже найдены (если есть)
        If IsAllFound(wsMain, MAIN_W_IN_NAME, 4, lastRow) Then
            If Not Mod_Constants.SilenceMsgBox Then
                MsgBox "ЗН заполнен!", vbInformation, "АВТО ЗЧ"
            End If
        Else
            If Not Mod_Constants.SilenceMsgBox Then
                MsgBox "Автоподбор запчастей завершён. Найдено: " & matchCount & _
                       ". Осталось подобрать работы.", vbInformation, "АВТО ЗЧ"
            End If
        End If
    Else
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Автоподбор запчастей завершён." & vbCrLf & _
                   "Найдено: " & matchCount & vbCrLf & _
                   "Не найдено: " & notFoundCount & vbCrLf & _
                   "Ненайденные строки подсвечены жёлтым.", _
                   vbExclamation, "АВТО ЗЧ"
        End If
    End If

    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_AutoMatch", "AutoMatchParts: Ошибка — " & Err.Description)
    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox "Ошибка при автоподборе запчастей: " & Err.Description, vbCritical, "АВТО ЗЧ"
    End If
End Sub