Attribute VB_Name = "Mod_MainButtons"
Option Explicit

' ============================================================
' Модуль: Mod_MainButtons
' Назначение: Обработчики кнопок листа main
' Отдельный модуль, чтобы не смешивать с Mod_ButtonDispatcher
' ============================================================

' ============================================================
' ОЧИСТКА ДАННЫХ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_Clear
' Безопасная очистка всех данных на листе main:
'   - очищает диапазон B2:ZZ<lastRow>
'   - НЕ трогает строку 1 (заголовки)
'   - НЕ трогает столбец A (метки/номера)
'   - запрашивает подтверждение перед очисткой
' --------------------------------------------------------------------------
Public Sub Btn_main_Clear()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim response As VbMsgBoxResult

    ' 1. Запрос подтверждения
    response = MsgBox( _
        "Очистить все данные на листе main?" & vbCrLf & _
        "Строка 1 (заголовки) и столбец A (метки) будут сохранены.", _
        vbYesNo + vbQuestion, _
        "Подтверждение очистки")

    If response <> vbYes Then Exit Sub

    ' 2. Определение последней используемой строки
    Set wsMain = ThisWorkbook.Sheets("main")
    lastRow = wsMain.UsedRange.Rows.Count

    ' Если данных для очистки нет — выходим
    If lastRow < 2 Then
        MsgBox "Нет данных для очистки.", vbInformation, "SysW"
        Exit Sub
    End If

    ' 3. Очистка диапазона B2:ZZ<lastRow>
    '    Столбец A НЕ очищаем, строку 1 НЕ очищаем
    wsMain.Range("B2:ZZ" & lastRow).ClearContents

    ' 4. Уведомление об успехе
    MsgBox "Очистка завершена.", vbInformation, "SysW"

    Exit Sub

ErrHandler:
    MsgBox "Ошибка при очистке: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Btn_main_Clear: " & Err.Description)
End Sub

' ============================================================
' ИМПОРТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_Import
' Полный цикл импорта из report.xlsx по ГРЗ из ячейки B4 листа main.
' Алгоритм (8 этапов с проверками):
'   1. Проверить, открыта ли уже report.xlsx. Если нет — открыть ReadOnly.
'   2. Прочитать ГРЗ из main!B4.
'   3. Вызвать ExtractNumberFromGRZ(grz) из Mod_Import — получить первые 3 цифры.
'   4. Искать среди листов report.xlsx лист с именем = этой группе цифр.
'   5. Если лист найден — переименовать его в main!$B$2 & "M" (например, "101M").
'   6. Скопировать лист в work.xlsm.
'   7. Импортировать данные на лист main:
'      - Таблица «Выполненные работы» → столбцы L:N
'      - Таблица «Расходная накладная» → столбцы X:AA
'   8. Закрыть report.xlsx (без сохранения).
' --------------------------------------------------------------------------
Public Sub Btn_main_Import()
    On Error GoTo ErrHandler

    Dim wsMain As Worksheet
    Dim grzRaw As String
    Dim grzNumber As String
    Dim wbReport As Workbook
    Dim wsSource As Worksheet
    Dim ws As Worksheet
    Dim newSheetName As String
    Dim wsCopied As Worksheet
    Dim reportPath As String

    Set wsMain = ThisWorkbook.Sheets("main")

    ' ==========================================
    ' ЭТАП 1: Извлечение ГРЗ из ячейки B4
    ' ==========================================
    grzRaw = Trim(CStr(wsMain.Range("B4").Value))

    If grzRaw = "" Then
        MsgBox "Ячейка B4 (ГРЗ) пуста. Заполните ГРЗ и повторите.", _
               vbExclamation, "Импорт невозможен"
        Exit Sub
    End If

    ' ==========================================
    ' ЭТАП 2: Извлечение первой группы из 3 цифр
    ' ==========================================
    grzNumber = Mod_Import.ExtractNumberFromGRZ(grzRaw)

    If grzNumber = "" Or Len(grzNumber) < 3 Then
        MsgBox "Не удалось извлечь 3-значный код из ГРЗ '" & grzRaw & "'.", _
               vbExclamation, "Ошибка ГРЗ"
        Exit Sub
    End If

    ' ==========================================
    ' ЭТАП 3: Открытие report.xlsx
    ' ==========================================
    reportPath = ThisWorkbook.Path & "\report.xlsx"

    ' Проверка существования файла
    If Not Mod_Utils.FileExists(reportPath) Then
        MsgBox "Файл report.xlsx не найден в папке " & ThisWorkbook.Path & ".", _
               vbExclamation, "Файл не найден"
        Exit Sub
    End If

    ' Проверяем, не открыта ли уже книга
    On Error Resume Next
    Set wbReport = Workbooks("report.xlsx")
    If wbReport Is Nothing Then
        ' Если не открыта — открываем ReadOnly
        Set wbReport = Workbooks.Open(reportPath, ReadOnly:=True)
    End If
    On Error GoTo ErrHandler

    If wbReport Is Nothing Then
        MsgBox "Не удалось открыть report.xlsx.", _
               vbExclamation, "Ошибка открытия"
        Exit Sub
    End If

    ' ==========================================
    ' ЭТАП 4: Поиск листа по 3-значному коду
    ' ==========================================
    Set wsSource = Nothing

    For Each ws In wbReport.Sheets
        ' Ищем точное совпадение имени листа с grzNumber
        ' или содержащего grzNumber
        If ws.Name = grzNumber Or InStr(1, ws.Name, grzNumber, vbTextCompare) > 0 Then
            Set wsSource = ws
            Exit For
        End If
    Next ws

    If wsSource Is Nothing Then
        MsgBox "В report.xlsx не найден лист с кодом '" & grzNumber & "'." & vbCrLf & _
               "Проверьте, что лист существует и содержит нужный ГРЗ.", _
               vbExclamation, "Лист не найден"
        wbReport.Close False
        Exit Sub
    End If

    ' ==========================================
    ' ЭТАП 5: Проверка B2 и переименование листа
    ' ==========================================
    If wsMain.Range("B2").Value = "" Then
        MsgBox "Ячейка B2 (номер заказа) пуста. Заполните номер заказа.", _
               vbExclamation, "Ошибка"
        wbReport.Close False
        Exit Sub
    End If

    newSheetName = CStr(wsMain.Range("B2").Value) & "M"

    ' Переименовываем лист в report.xlsx
    On Error Resume Next
    wsSource.Name = newSheetName
    If Err.Number <> 0 Then
        Err.Clear
        ' Если имя занято — добавляем суффикс с временем
        wsSource.Name = newSheetName & "_" & Format(Now, "hhmmss")
    End If
    On Error GoTo ErrHandler

    ' ==========================================
    ' ЭТАП 6: Копирование листа в work.xlsm
    ' ==========================================
    wsSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' ==========================================
    ' ЭТАП 7: Импорт данных на лист main
    ' ==========================================
    ' Только что скопированный лист становится активным
    Set wsCopied = ActiveSheet

    Call Mod_Import.ImportDataToMain(wsCopied)

    ' ==========================================
    ' ЭТАП 8: Завершение — закрытие report.xlsx
    ' ==========================================
    wbReport.Close False

    MsgBox "Импорт из листа '" & newSheetName & "' выполнен успешно.", _
           vbInformation, "Импорт завершён"

    Exit Sub

ErrHandler:
    ' Закрываем report.xlsx, если он был открыт
    On Error Resume Next
    If Not wbReport Is Nothing Then
        wbReport.Close False
    End If
    On Error GoTo 0

    MsgBox "Ошибка при импорте: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Btn_main_Import: " & Err.Description)
End Sub

' ============================================================
' АВТОПОДБОР ЗАПЧАСТЕЙ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_AUTOz4
' Заглушка: автоподбор запчастей — в разработке
' --------------------------------------------------------------------------
Public Sub Btn_main_AUTOz4()
    MsgBox "Автоподбор запчастей — в разработке.", _
           vbInformation, "SysW"
End Sub

' ============================================================
' АВТОПОДБОР РАБОТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_AUTOw
' Заглушка: автоподбор работ — в разработке
' --------------------------------------------------------------------------
Public Sub Btn_main_AUTOw()
    MsgBox "Автоподбор работ — в разработке.", _
           vbInformation, "SysW"
End Sub

' ============================================================
' РУЧНОЙ ПОДБОР ЗАПЧАСТЕЙ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_MANz4
' Заглушка: ручной подбор запчастей — в разработке
' --------------------------------------------------------------------------
Public Sub Btn_main_MANz4()
    MsgBox "Ручной подбор запчастей — в разработке.", _
           vbInformation, "SysW"
End Sub

' ============================================================
' РУЧНОЙ ПОДБОР РАБОТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_MANw
' Заглушка: ручной подбор работ — в разработке
' --------------------------------------------------------------------------
Public Sub Btn_main_MANw()
    MsgBox "Ручной подбор работ — в разработке.", _
           vbInformation, "SysW"
End Sub