Attribute VB_Name = "Mod_Models"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Модуль: Mod_Models
' Назначение: Инициализация и управление листом models
' ============================================================

' ============================================================
' InitModelsSheet
' Создаёт или переименовывает лист model -> models
' и заполняет его структурой:
'   Строка 1 — заголовки (Модель, Группа, Цена н/ч)
'   Строка 2 — латинские имена (model, group, hrpr)
'   Строка 3+ — данные из старого листа model (если был)
' ============================================================
Public Sub InitModelsSheet()
    On Error GoTo ErrHandler

    Dim wsModels As Worksheet
    Dim wsOldModel As Worksheet
    Dim lastRowOld As Long
    Dim lastColOld As Long
    Dim dataRange As Range
    Dim targetRow As Long

    ' 1. Если лист models уже существует — ничего не делаем
    Set wsModels = Mod_Utils.GetSheetByName(ThisWorkbook, "models")
    If Not wsModels Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Лист models уже существует, пропуск")
        Exit Sub
    End If

    ' 2. Проверяем, существует ли старый лист model
    Set wsOldModel = Mod_Utils.GetSheetByName(ThisWorkbook, "model")

    If Not wsOldModel Is Nothing Then
        ' 2a. Переименовываем model -> models
        wsOldModel.Name = "models"
        Set wsModels = wsOldModel
        Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Лист model переименован в models")
    Else
        ' 2b. Создаём новый лист models
        Set wsModels = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsModels.Name = "models"
        Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Создан новый лист models")
    End If

    ' 3. Заполняем заголовки (строка 1)
    wsModels.Cells(1, MODELS_COL_MODEL).Value = "Модель"
    wsModels.Cells(1, MODELS_COL_GROUP).Value = "Группа"
    wsModels.Cells(1, MODELS_COL_PRICE).Value = "Цена н/ч"

    ' 4. Заполняем латинские имена (строка 2)
    wsModels.Cells(2, MODELS_COL_MODEL).Value = "model"
    wsModels.Cells(2, MODELS_COL_GROUP).Value = "group"
    wsModels.Cells(2, MODELS_COL_PRICE).Value = "hrpr"

    ' 5. Если был старый лист model с данными — переносим их
    '    Данные начинаются со строки 3 (после двух строк заголовков)
    If Not wsOldModel Is Nothing Then
        lastRowOld = wsOldModel.Cells(wsOldModel.Rows.Count, 1).End(xlUp).Row
        If lastRowOld >= 3 Then
            ' Определяем последний используемый столбец
            lastColOld = wsOldModel.Cells(2, wsOldModel.Columns.Count).End(xlToLeft).Column
            If lastColOld < 1 Then lastColOld = 3

            ' Копируем данные со строки 3 до последней строки
            Set dataRange = wsOldModel.Range( _
                wsOldModel.Cells(3, 1), _
                wsOldModel.Cells(lastRowOld, lastColOld) _
            )
            targetRow = 3
            dataRange.Copy Destination:=wsModels.Cells(targetRow, 1)

            Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Перенесено " & _
                (lastRowOld - 2) & " строк данных из model")
        End If
    End If

    ' 6. Форматирование: заголовки жирным, автоширина
    With wsModels
        .Rows(1).Font.Bold = True
        .Columns("A:C").AutoFit
    End With

    Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Лист models успешно инициализирован")
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_Models", "InitModelsSheet: Ошибка — " & Err.Description)
    MsgBox "Ошибка при инициализации листа models: " & Err.Description, vbCritical, "Ошибка"
End Sub