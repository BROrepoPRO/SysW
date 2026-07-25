VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===========================================================
' ОБЪЕКТ: ThisWorkbook
' ОБРАБОТЧИКИ СОБЫТИЙ КНИГИ
'===========================================================

Option Explicit

' Обработчик открытия книги
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Инициализируем лист main
    Call Mod_Main.InitializeMainSheet
    
    ' Проверяем существование листа ЗН и инициализируем при необходимости
    InitializeZNSheet
    
    ' Инициализируем поиск на всех листах
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "_SETTINGS" And ws.Name <> "spisok" And ws.Name <> "model" Then
            Call Mod_Search.InitializeSheetForSearch(ws)
        End If
    Next ws
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Система автоматизации готова. Введите номер в B2"
End Sub

' Инициализация листа ЗН при необходимости
Private Sub InitializeZNSheet()
    Dim wsZN As Worksheet
    On Error Resume Next
    Set wsZN = ThisWorkbook.Sheets("ЗН")
    On Error GoTo 0
    
    If wsZN Is Nothing Then
        ' Лист ЗН не существует - создаем его
        Set wsZN = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsZN.Name = "ЗН"
        
        ' Настраиваем базовую структуру листа ЗН
        InitializeZNStructure wsZN
    End If
End Sub

' Настройка структуры листа ЗН
Private Sub InitializeZNStructure(ws As Worksheet)
    With ws
        ' Заголовки для столбцов
        .Range("B16").Value = "Сумма"
        .Range("K311").Value = 0 ' Инициализируем нулем
        
        ' Устанавливаем формат для ячеек с суммами
        .Range("K311").numberFormat = "#,##0.00"
        .Range("G313").numberFormat = "@" ' Текстовый формат
        
        ' Автоподбор ширины столбцов
        .Columns("B:B").AutoFit
        .Columns("G:G").AutoFit
        .Columns("K:K").AutoFit
    End With
End Sub

' Обработчик сохранения книги
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Сохраняем настройки макета
    Call Mod_Settings.SaveMainLayout
    Application.StatusBar = "Сохранение данных..."
End Sub

' Обработчик закрытия книги
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Сохраняем настройки перед закрытием
    Call Mod_Settings.SaveMainLayout
    Application.StatusBar = ""
End Sub
