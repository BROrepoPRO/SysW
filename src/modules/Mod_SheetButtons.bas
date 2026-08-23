Attribute VB_Name = "Mod_SheetButtons"
Option Explicit
Option Private Module

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

    ' Находим последний столбец с данными в строке заголовков (MODELS_HEADER_ROW = 3)
    lastCol = ws.Cells(Mod_Constants.MODELS_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    ' Сбрасываем предыдущий фильтр
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    ' Определяем диапазон для фильтрации (заголовки + данные со строки MODELS_DATA_START_ROW)
    Set dataRange = ws.Range(ws.Cells(Mod_Constants.MODELS_HEADER_ROW, 1), _
                             ws.Cells(lastRow, lastCol))

    ' Применяем фильтр "содержит" к указанному столбцу
    dataRange.AutoFilter _
        Field:=searchColumn, _
        Criteria1:="*" & searchValue & "*", _
        Operator:=xlFilterValues

    ' Считаем количество видимых строк (начиная с первой строки данных — MODELS_DATA_START_ROW)
    On Error Resume Next
    countVisible = ws.Range(ws.Cells(Mod_Constants.MODELS_DATA_START_ROW, searchColumn), _
                            ws.Cells(lastRow, searchColumn)) _
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

' ============================================================
' КНОПКИ ПОИСКА ДЛЯ ЛИСТОВ ЗАПЧАСТЕЙ (z4, {GroupName}z4)
' ============================================================
' Поле ввода: C1
' Данные: с 4-й строки
' Столбцы: B(2) — артикул, C(3) — наименование
' ============================================================

' --------------------------------------------------------------------------
' IsPartsSheet
' Проверяет, что активный лист является листом запчастей (z4 или {GroupName}z4).
' Лист spisok и листы работ ({GroupName}) в поиске НЕ участвуют.
' --------------------------------------------------------------------------
Private Function IsPartsSheet(ws As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = ws.Name

    ' Основной лист всех запчастей группы
    If LCase$(sheetName) = "z4" Then
        IsPartsSheet = True
        Exit Function
    End If

    ' Модельные листы запчастей с суффиксом z4 (например UAZz4, GAZz4)
    If Len(sheetName) > 2 Then
        If Right$(LCase$(sheetName), 2) = "z4" Then
            IsPartsSheet = True
            Exit Function
        End If
    End If

    IsPartsSheet = False
End Function

' --------------------------------------------------------------------------
' ExecutePartsSearch
' Общая функция поиска «содержит» по столбцу на активном листе запчастей.
' searchColumn: номер столбца для фильтрации (2=B Артикул, 3=C Наименование)
' --------------------------------------------------------------------------
Private Function ExecutePartsSearch(ByVal searchColumn As Integer) As Boolean
    Dim ws As Worksheet
    Dim searchValue As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim countVisible As Long
    Dim dataRange As Range

    Set ws = ActiveSheet

    ' Проверяем, что активный лист является листом запчастей
    If Not IsPartsSheet(ws) Then
        MsgBox "Поиск доступен только на листах запчастей (z4, {Группа}z4).", vbExclamation
        ExecutePartsSearch = False
        Exit Function
    End If

    ' Проверяем, подходит ли лист для поиска (данные с 4-й строки)
    If Not IsSearchableSheet(ws) Then
        MsgBox "Лист не содержит данных или не соответствует структуре." & vbCrLf & _
               "Данные должны начинаться с 4-й строки.", vbExclamation
        ExecutePartsSearch = False
        Exit Function
    End If

    ' Берём значение из ячейки C1 (поле ввода)
    searchValue = Trim(ws.Range("C1").Value)

    ' Проверяем, не пустое ли поле
    If searchValue = "" Then
        MsgBox "Введите текст для поиска в ячейке C1.", vbInformation
        ExecutePartsSearch = False
        Exit Function
    End If

    ' Находим последнюю строку данных (начиная с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Таблица не содержит данных (данные начинаются с 4-й строки).", vbExclamation
        ExecutePartsSearch = False
        Exit Function
    End If

    ' Находим последний столбец с данными в строке заголовков (MODELS_HEADER_ROW = 3)
    lastCol = ws.Cells(Mod_Constants.MODELS_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    ' Сбрасываем предыдущий фильтр
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    ' Определяем диапазон для фильтрации (заголовки + данные со строки MODELS_DATA_START_ROW)
    Set dataRange = ws.Range(ws.Cells(Mod_Constants.MODELS_HEADER_ROW, 1), _
                             ws.Cells(lastRow, lastCol))

    ' Применяем фильтр «содержит» к указанному столбцу
    dataRange.AutoFilter _
        Field:=searchColumn, _
        Criteria1:="*" & searchValue & "*", _
        Operator:=xlFilterValues

    ' Считаем количество видимых строк (начиная с первой строки данных — MODELS_DATA_START_ROW)
    On Error Resume Next
    countVisible = ws.Range(ws.Cells(Mod_Constants.MODELS_DATA_START_ROW, searchColumn), _
                            ws.Cells(lastRow, searchColumn)) _
        .SpecialCells(xlCellTypeVisible).Count
    On Error GoTo 0

    ' Проверяем результат
    If countVisible > 0 Then
        MsgBox "Найдено строк: " & countVisible, vbInformation
        ExecutePartsSearch = True
    Else
        MsgBox "Совпадений не найдено.", vbExclamation
        ws.ShowAllData
        ExecutePartsSearch = False
    End If

    ' Убираем фокус с поля ввода
    ws.Range("A1").Select
End Function

' --------------------------------------------------------------------------
' Btn_Parts_SearchByArticle
' Поиск по артикулу (столбец B) на активном листе запчастей
' --------------------------------------------------------------------------
Public Sub Btn_Parts_SearchByArticle()
    ExecutePartsSearch searchColumn:=2
End Sub

' --------------------------------------------------------------------------
' Btn_Parts_SearchByName
' Поиск по наименованию (столбец C) на активном листе запчастей
' --------------------------------------------------------------------------
Public Sub Btn_Parts_SearchByName()
    ExecutePartsSearch searchColumn:=3
End Sub

' --------------------------------------------------------------------------
' Btn_Parts_ClearFilter
' Сброс фильтра и очистка поля ввода C1 на активном листе запчастей
' --------------------------------------------------------------------------
Public Sub Btn_Parts_ClearFilter()
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Проверяем, что активный лист является листом запчастей
    If Not IsPartsSheet(ws) Then
        MsgBox "Сброс фильтра доступен только на листах запчастей (z4, {Группа}z4).", vbExclamation
        Exit Sub
    End If

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