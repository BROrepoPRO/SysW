VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==============================================
' ОБЩАЯ ФУНКЦИЯ ДЛЯ ВСЕХ МАКРОСОВ
' ==============================================
Private Function ВыполнитьПоиск(столбец As Integer) As Boolean
    Dim searchValue As String
    Dim lastRow As Long
    Dim countVisible As Long
    
    ' Берем значение из текстового поля
    searchValue = Trim(Me.TextBoxSearch.text)
    
    ' Проверяем, не пустое ли поле
    If searchValue = "" Then
        MsgBox "Введите текст для поиска в поле выше.", vbInformation
        ВыполнитьПоиск = False
        Exit Function
    End If
    
    ' Находим последнюю строку данных
    lastRow = Me.Cells(Me.Rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Таблица не содержит данных.", vbExclamation
        ВыполнитьПоиск = False
        Exit Function
    End If
    
    ' Сбрасываем предыдущий фильтр
    On Error Resume Next
    Me.ShowAllData
    On Error GoTo 0
    
    ' Применяем фильтр "содержит" к указанному столбцу
    Me.Range("A4:F" & lastRow).AutoFilter _
        Field:=столбец, _
        Criteria1:="*" & searchValue & "*"
    
    ' Считаем количество видимых строк
    On Error Resume Next
    countVisible = Me.Range("B4:B" & lastRow).SpecialCells(xlCellTypeVisible).count
    On Error GoTo 0
    
    ' Проверяем результат
    If countVisible > 0 Then
        MsgBox "Найдено строк: " & countVisible, vbInformation
        ВыполнитьПоиск = True
    Else
        MsgBox "Совпадений не найдено.", vbExclamation
        Me.ShowAllData
        ВыполнитьПоиск = False
    End If
End Function

' ==============================================
' КНОПКА 1: ПОИСК В АРТИКУЛЕ (СТОЛБЕЦ B)
' ==============================================
Private Sub BtnArticle_Click()
    ' Вызываем общую функцию для столбца B (это 2-й столбец в диапазоне A4:F...)
    ВыполнитьПоиск столбец:=2
End Sub

' ==============================================
' КНОПКА 2: ПОИСК В НАИМЕНОВАНИИ (СТОЛБЕЦ C)
' ==============================================
Private Sub BtnName_Click()
    ' Вызываем общую функцию для столбца C (это 3-й столбец в диапазоне A4:F...)
    ВыполнитьПоиск столбец:=3
End Sub

' ==============================================
' КНОПКА 3: ОЧИСТКА ФИЛЬТРА
' ==============================================
Private Sub BtnClear_Click()
    ' Сбрасываем все фильтры
    On Error Resume Next
    Me.ShowAllData
    On Error GoTo 0
    
    ' Очищаем текстовое поле
    Me.TextBoxSearch.text = ""
    
    ' Убираем фокус с поля ввода
    Me.Range("A1").Select
    
    ' Сообщение пользователю
    MsgBox "Все фильтры сброшены. Таблица показывает все данные.", vbInformation
End Sub

' ==============================================
' ДОПОЛНИТЕЛЬНАЯ ФУНКЦИОНАЛЬНОСТЬ
' ==============================================

' Автозапуск поиска при нажатии Enter в текстовом поле
Private Sub TextBoxSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' 13 = клавиша Enter
        ' Автоматически запускаем поиск по артикулу при нажатии Enter
        BtnArticle_Click
    End If
End Sub







