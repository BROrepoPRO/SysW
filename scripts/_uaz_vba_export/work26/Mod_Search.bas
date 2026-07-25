Attribute VB_Name = "Mod_Search"
'===========================================================
' МОДУЛЬ: Mod_Search
' НАЗНАЧЕНИЕ: Ручной поиск в листах по артикулу и наименованию
' УСЛОВИЯ: Строки 1-3 технические, данные с 4 строки
'          Поиск по столбцам B (Артикул) и C (Наименование)
'          Ячейка C1 - поле ввода поиска
'===========================================================

Option Explicit

' Вспомогательная функция: проверяет, подходит ли лист для поиска
Private Function IsSearchableSheet(ws As Worksheet) As Boolean
    ' Лист должен иметь минимум 4 строки данных
    If ws.Cells(ws.Rows.count, "A").End(xlUp).Row >= 4 Then
        IsSearchableSheet = True
    Else
        IsSearchableSheet = False
    End If
End Function

' ==============================================
' ОБЩАЯ ФУНКЦИЯ ДЛЯ ВСЕХ МАКРОСОВ
' ==============================================
Public Function ExecuteSearch(ByVal searchColumn As Integer) As Boolean
    Dim ws As Worksheet
    Dim searchValue As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim countVisible As Long
    Dim dataRange As Range
    
    ' Работаем с активным листом
    Set ws = ActiveSheet
    
    ' Проверяем, подходит ли лист для поиска
    If Not IsSearchableSheet(ws) Then
        MsgBox "Лист не содержит данных или не соответствует структуре." & vbCrLf & _
               "Данные должны начинаться с 4-й строки.", vbExclamation
        ExecuteSearch = False
        Exit Function
    End If
    
    ' Берём значение из ячейки C1 (поле ввода)
    searchValue = Trim(ws.Range("C1").Value)
    
    ' Проверяем, не пустое ли поле
    If searchValue = "" Then
        MsgBox "Введите текст для поиска в ячейке C1.", vbInformation
        ExecuteSearch = False
        Exit Function
    End If
    
    ' Находим последнюю строку данных (начиная с 4-й строки)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Таблица не содержит данных (данные начинаются с 4-й строки).", vbExclamation
        ExecuteSearch = False
        Exit Function
    End If
    
    ' Находим последний столбец с данными в 4-й строке (строка заголовков)
    lastCol = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column
    
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
        .SpecialCells(xlCellTypeVisible).count
    On Error GoTo 0
    
    ' Проверяем результат
    If countVisible > 0 Then
        MsgBox "Найдено строк: " & countVisible, vbInformation
        ExecuteSearch = True
    Else
        MsgBox "Совпадений не найдено.", vbExclamation
        ws.ShowAllData
        ExecuteSearch = False
    End If
    
    ' Убираем фокус с поля ввода
    ws.Range("A1").Select
End Function

' ==============================================
' КНОПКА 1: ПОИСК ПО АРТИКУЛУ
' ==============================================
Public Sub SearchByArticle()
    ' Вызываем общую функцию для столбца B (Артикул)
    ExecuteSearch searchColumn:=2
End Sub

' ==============================================
' КНОПКА 2: ПОИСК ПО НАИМЕНОВАНИЮ
' ==============================================
Public Sub SearchByName()
    ' Вызываем общую функцию для столбца C (Наименование)
    ExecuteSearch searchColumn:=3
End Sub

' ==============================================
' КНОПКА 3: ОЧИСТКА ФИЛЬТРА
' ==============================================
Public Sub ClearSearchFilter()
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
    
    ' Сообщение пользователю
    MsgBox "Все фильтры сброшены. Таблица показывает все данные.", vbInformation
End Sub

' ==============================================
' ФУНКЦИЯ ДЛЯ ИНИЦИАЛИЗАЦИИ ЛИСТА
' ==============================================
Public Sub InitializeSheetForSearch(ws As Worksheet)
    ' Убеждаемся, что ячейка C1 свободна для ввода поиска
    ' (Можно добавить подсказку, но не обязательно)
    If ws.Range("C1").Value = "" Then
        ' Опционально: можно добавить подсказку
        ' ws.Range("C1").Value = "Введите текст для поиска"
    End If
    
    
End Sub

