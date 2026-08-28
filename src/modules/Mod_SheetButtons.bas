Attribute VB_Name = "Mod_SheetButtons"
Option Explicit

' ============================================================
' Модуль: Mod_SheetButtons
' Назначение: Обработчики кнопок поиска на листах работ и запчастей
' ============================================================

' ============================================================
' КНОПКИ ПОИСКА ПО АРТИКУЛУ/НАИМЕНОВАНИЮ (работы и запчасти)
' ============================================================
' Адаптировано из Mod_Search (work26.xlsm)
' Имя группы динамически читается из main!$B$14.
' Поле ввода: C1
' Данные: с 4-й строки
' ============================================================

' ============================================================
' Тип листа относительно имени группы {Group}
' ============================================================
Public Enum SheetKind
    skUnknown = 0      ' прочее (не участвует в поиске)
    skWorks = 1        ' {Group}
    skWorksModel = 2   ' {Group}w
    skParts = 3        ' z4
    skPartsModel = 4   ' {Group}z4
End Enum

' --------------------------------------------------------------------------
' ResolveGroupName
' Определяет имя группы для активного листа. Сначала читает main!$B$14
' через единый хелпер Mod_Utils.GetGroupName; если он пуст/недоступен
' (например, в модельной книге нет листа main), выводит группу из имени
' листа ({Group}, {Group}w, {Group}z4) либо "" для z4.
' Это позволяет поиску работать в любом листе книги без жёсткой привязки к main.
' --------------------------------------------------------------------------
Private Function ResolveGroupName(ByVal ws As Worksheet) As String
    Dim g As String
    g = Mod_Utils.GetGroupName()
    If g <> "" Then
        ResolveGroupName = g
        Exit Function
    End If

    Dim n As String
    n = LCase$(Trim(ws.Name))
    If n = "z4" Then
        ResolveGroupName = ""
    ElseIf Right$(n, 1) = "w" Then
        ResolveGroupName = Left$(n, Len(n) - 1)
    ElseIf Right$(n, 2) = "z4" Then
        ResolveGroupName = Left$(n, Len(n) - 2)
    Else
        ResolveGroupName = n
    End If
End Function

' --------------------------------------------------------------------------
' ClassifySheet
' Универсальная классификация листа относительно имени группы {Group}:
'   {Group}   -> skWorks
'   {Group}w  -> skWorksModel
'   z4        -> skParts
'   {Group}z4 -> skPartsModel
'   прочее    -> skUnknown
' Имя листа никогда не зашито в код: все правила строятся от groupName + суффиксов.
' --------------------------------------------------------------------------
Public Function ClassifySheet(ByVal ws As Worksheet, ByVal groupName As String) As SheetKind
    Dim name As String
    Dim g As String

    name = LCase$(Trim(ws.Name))
    g = LCase$(Trim(groupName))

    If name = g Then
        ClassifySheet = skWorks
    ElseIf g <> "" And name = g & "w" Then
        ClassifySheet = skWorksModel
    ElseIf name = "z4" Then
        ClassifySheet = skParts
    ElseIf g <> "" And name = g & "z4" Then
        ClassifySheet = skPartsModel
    Else
        ClassifySheet = skUnknown
    End If
End Function

' --------------------------------------------------------------------------
' IsSearchableSheet
' Проверяет, подходит ли лист для поиска (минимум 4 строки данных)
' --------------------------------------------------------------------------
Private Function IsSearchableSheet(ws As Worksheet) As Boolean
    IsSearchableSheet = (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row >= 4)
End Function

' --------------------------------------------------------------------------
' ExecuteSearch
' Единая функция поиска «содержит» по активному листу.
' Работает на любом листе работ/запчастей книги (группа читается из B14).
' searchColumn: номер столбца для фильтрации (2=B Артикул, 3=C Наименование)
' --------------------------------------------------------------------------
Private Function ExecuteSearch(ByVal searchColumn As Integer) As Boolean
    Dim ws As Worksheet
    Dim groupName As String
    Dim searchValue As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim countVisible As Long
    Dim dataRange As Range

    Set ws = ActiveSheet

    ' Определяем группу (main!$B$14 или из имени листа) и классифицируем лист
    groupName = ResolveGroupName(ws)

    If ClassifySheet(ws, groupName) = skUnknown Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Поиск доступен только на листах работ или запчастей." & vbCrLf & _
                   "(листы {Группа}, {Группа}w, z4, {Группа}z4).", vbExclamation
        End If
        ExecuteSearch = False
        Exit Function
    End If

    ' Проверяем, подходит ли лист для поиска
    If Not IsSearchableSheet(ws) Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Лист не содержит данных или не соответствует структуре." & vbCrLf & _
                   "Данные должны начинаться с 4-й строки.", vbExclamation
        End If
        ExecuteSearch = False
        Exit Function
    End If

    ' Берём значение из ячейки C1 (поле ввода)
    searchValue = Trim(ws.Range("C1").Value)

    ' Проверяем, не пустое ли поле
    If searchValue = "" Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Введите текст для поиска в ячейке C1.", vbInformation
        End If
        ExecuteSearch = False
        Exit Function
    End If

    ' Находим последнюю строку данных (начиная с 4-й строки)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Таблица не содержит данных (данные начинаются с 4-й строки).", vbExclamation
        End If
        ExecuteSearch = False
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
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Найдено строк: " & countVisible, vbInformation
        End If
        ExecuteSearch = True
    Else
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Совпадений не найдено.", vbExclamation
        End If
        ws.ShowAllData
        ExecuteSearch = False
    End If

    ' Убираем фокус с поля ввода
    ws.Range("A1").Select
End Function

' --------------------------------------------------------------------------
' Btn_Search_ByArticle
' Поиск по артикулу (столбец B) на активном листе работ/запчастей
' --------------------------------------------------------------------------
Public Sub Btn_Search_ByArticle()
    ExecuteSearch searchColumn:=2
End Sub

' --------------------------------------------------------------------------
' Btn_Search_ByName
' Поиск по наименованию (столбец C) на активном листе работ/запчастей
' --------------------------------------------------------------------------
Public Sub Btn_Search_ByName()
    ExecuteSearch searchColumn:=3
End Sub

' --------------------------------------------------------------------------
' Btn_ClearFilter
' Сброс фильтра и очистка поля ввода C1 на активном листе работ/запчастей
' --------------------------------------------------------------------------
Public Sub Btn_ClearFilter()
    Dim ws As Worksheet
    Dim groupName As String

    Set ws = ActiveSheet

    ' Классифицируем лист: сброс доступен только на листах работ/запчастей
    groupName = ResolveGroupName(ws)

    If ClassifySheet(ws, groupName) = skUnknown Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Сброс фильтра доступен только на листах работ или запчастей " & vbCrLf & _
                   "(листы {Группа}, {Группа}w, z4, {Группа}z4).", vbExclamation
        End If
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

    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox "Все фильтры сброшены. Таблица показывает все данные.", vbInformation
    End If
End Sub