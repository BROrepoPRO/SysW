VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===========================================================
' ОБЪЕКТ: Worksheet - main
' ОБРАБОТЧИКИ СОБЫТИЙ
'===========================================================

Option Explicit

' Событие изменения ячейки
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Реагируем только на изменение ячейки B2
    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        ' Даем Excel время обновить значение
        DoEvents
        
        ' Если введено значение - автоматически заполняем шапку
        If Not IsEmpty(Me.Range("B2").Value) And IsNumeric(Me.Range("B2").Value) Then
            Application.EnableEvents = False
            On Error Resume Next
            Call Mod_Main.FillHeaderTransposed
            Application.EnableEvents = True
        ElseIf IsEmpty(Me.Range("B2").Value) Then
            ' Если очистили поле - очищаем шапку
            Application.EnableEvents = False
            On Error Resume Next
            Call Mod_Main.ClearHeader
            Application.EnableEvents = True
        End If
    End If
End Sub

' Событие активации листа
Private Sub Worksheet_Activate()
    Application.StatusBar = "Введите номер строки из листа 'spisok' в ячейку B2"
End Sub
