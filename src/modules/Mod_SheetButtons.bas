Attribute VB_Name = "Mod_SheetButtons"
Option Explicit

' ============================================================
' Модуль: Mod_SheetButtons
' Назначение: Обработчики кнопок листов z4 и work
' ============================================================

' ============================================================
' КНОПКИ ПОИСКА ДЛЯ ЛИСТОВ UAZ (UAZ, UAZw, z4, UAZz4)
' ============================================================
' Адаптировано из Mod_Search (work26.xlsm)
' Поле ввода: C1
' Данные: с 4-й строки
' ============================================================

' --------------------------------------------------------------------------
' IsSearchableSheet
' Проверяет, подходит ли лист для поиска (минимум 4 строки данных)
' --------------------------------------------------------------------------
Private Function IsSearchableSheet(ws As Worksheet) As Boolean
    IsSearchableSheet = (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row >= 4)
End Function

' --------------------------------------------------------------------------
' ExecuteUAZSearch
' Общая функция поиска для листов UAZ.
' searchColumn: номер столбца для фильтрации (2=B Артикул, 3=C Наименование)
' --------------------------------------------------------------------------
Private Function ExecuteUAZSearch(ByVal searchColumn As Integer) As Boolean
    Dim ws As Worksheet
    Dim searchValue As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim countVisible As Long
    Dim dataRange As Range

    Set ws = ActiveSheet

    ' Проверяем, подходит ли лист для поиска
    If Not IsSearchableSheet(ws) Then
        MsgBox "Лист не содержит данных или не соответствует структуре." & vbCrLf & _
               "Данные должны начинаться с 4-й строки.", vbExclamation
        ExecuteUAZSearch = False
        Exit Function
    End If

    ' Берём значение из ячейки C1 (поле ввода)
    searchValue = Trim(ws.Range("C1").Value)

    ' Проверяем, не пустое ли поле
    If searchValue = "" Then
        MsgBox "Введите текст для поиска в ячейке C1.", vbInformation
        ExecuteUAZSearch = False
        Exit Function
    End If

    ' Находим последнюю строку данных (начиная с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Таблица не содержит данных (данные начинаются с 4-й строки).", vbExclamation
        ExecuteUAZSearch = False
        Exit Function
    End If

    ' Находим последний столбец с данными в 4-й строке (строка заголовков)
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    ' Сбрасываем предыдущий фильтр
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    ' Определяем диапазон для фильтрации (все данные с 4 строки)
    Set dataRange = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol))

    ' Применяем фильтр "содержит" к указанному столбцу
    dataRange.AutoFilter _
        Field:=searchColumn, _
        Criteria1:="*" & searchValue & "*", _
        Operator:=xlFilterValues

    ' Считаем количество видимых строк (начиная с 5-й строки)
    On Error Resume Next
    countVisible = ws.Range(ws.Cells(5, searchColumn), ws.Cells(lastRow, searchColumn)) _
        .SpecialCells(xlCellTypeVisible).Count
    On Error GoTo 0

    ' Проверяем результат
    If countVisible > 0 Then
        MsgBox "Найдено строк: " & countVisible, vbInformation
        ExecuteUAZSearch = True
    Else
        MsgBox "Совпадений не найдено.", vbExclamation
        ws.ShowAllData
        ExecuteUAZSearch = False
    End If

    ' Убираем фокус с поля ввода
    ws.Range("A1").Select
End Function

' --------------------------------------------------------------------------
' Btn_UAZ_SearchByArticle
' Поиск по артикулу (столбец B) на активном листе UAZ
' --------------------------------------------------------------------------
Public Sub Btn_UAZ_SearchByArticle()
    ExecuteUAZSearch searchColumn:=2
End Sub

' --------------------------------------------------------------------------
' Btn_UAZ_SearchByName
' Поиск по наименованию (столбец C) на активном листе UAZ
' --------------------------------------------------------------------------
Public Sub Btn_UAZ_SearchByName()
    ExecuteUAZSearch searchColumn:=3
End Sub

' --------------------------------------------------------------------------
' Btn_UAZ_ClearFilter
' Сброс фильтра и очистка поля ввода C1 на активном листе UAZ
' --------------------------------------------------------------------------
Public Sub Btn_UAZ_ClearFilter()
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Сбрасываем все фильтры
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    ' Очищаем поле ввода (ячейка C1)
    ws.Range("C1").Value = ""

    ' Убираем фокус с поля ввода
    ws.Range("A1").Select

    MsgBox "Все фильтры сброшены. Таблица показывает все данные.", vbInformation
End Sub