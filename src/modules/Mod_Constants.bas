Attribute VB_Name = "Mod_Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Модуль: Mod_Constants
' Назначение: Централизованное хранение констант проекта
'             и управление листом libname (реестр имён)
' ============================================================

' ============================================================
' Константы столбцов листа spisok
' ============================================================
Public Const SPISOK_COL_NUM As Long = 1
Public Const SPISOK_COL_MODEL As Long = 2
Public Const SPISOK_COL_GRZ As Long = 3
Public Const SPISOK_COL_VIN As Long = 4
Public Const SPISOK_COL_GARAGE As Long = 5
Public Const SPISOK_COL_YEAR As Long = 6
Public Const SPISOK_COL_MILEAGE As Long = 7
Public Const SPISOK_COL_DATE As Long = 8
Public Const SPISOK_COL_GROUP As Long = 9
Public Const SPISOK_COL_NOTE As Long = 10

' ============================================================
' Константы столбцов листа models
' ============================================================
Public Const MODELS_COL_MODEL As Long = 1
Public Const MODELS_COL_GROUP As Long = 2
Public Const MODELS_COL_PRICE As Long = 3

' ============================================================
' Строковые константы для листа libname (реестр имён)
' Соглашение: {лист}_COL_{сущность}_NAME — England-имя
' ============================================================

' --- Лист spisok ---
Public Const SPISOK_COL_NUM_NAME As String = "spisok"
Public Const SPISOK_COL_MODEL_NAME As String = "model"
Public Const SPISOK_COL_GRZ_NAME As String = "grz"
Public Const SPISOK_COL_VIN_NAME As String = "vin"
Public Const SPISOK_COL_GARAGE_NAME As String = "garnum"
Public Const SPISOK_COL_YEAR_NAME As String = "year"
Public Const SPISOK_COL_MILEAGE_NAME As String = "mileage"
Public Const SPISOK_COL_DATE_NAME As String = "date"
Public Const SPISOK_COL_GROUP_NAME As String = "group"
Public Const SPISOK_COL_NOTE_NAME As String = "reserve"

' --- Лист models ---
Public Const MODELS_COL_MODEL_NAME As String = "model_name"
Public Const MODELS_COL_GROUP_NAME As String = "group"
Public Const MODELS_COL_PRICE_NAME As String = "hrpr"

' --- Глобальные ---
Public Const WORK_NAME As String = "work"
Public Const Z4_NAME As String = "z4"

' ============================================================
' Приватный тип: запись реестра libname
' ============================================================
Private Type LibNameEntry
    NameKey As String    ' {_name} — объявленное имя
    England As String    ' England — значение/пояснение на английском
    Russian As String    ' Русский — описание на русском
End Type

' ============================================================
' InitLibName
' Заполняет лист libname начальными данными реестра имён.
' Если лист уже содержит данные (непустая строка 2) — пропускает.
' ============================================================
Public Sub InitLibName()
    On Error GoTo ErrHandler

    Dim wsLib As Worksheet
    Dim entries As Variant
    Dim i As Long

    ' 1. Проверка существования листа libname
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, "libname")
    If wsLib Is Nothing Then
        Call Mod_Logger.WriteLog("Mod_Constants", "InitLibName: Лист libname не найден")
        MsgBox "Лист libname не найден в книге. Заполнение прервано.", vbCritical, "Ошибка"
        Exit Sub
    End If

    ' 2. Проверка, не заполнен ли уже лист (строка 2 непуста)
    If Not IsEmpty(wsLib.Cells(2, 1).Value) Then
        Call Mod_Logger.WriteLog("Mod_Constants", "InitLibName: Лист libname уже содержит данные, пропуск")
        Exit Sub
    End If

    ' 3. Получение массива записей
    entries = BuildEntryArray()

    ' 4. Запись данных построчно
    For i = LBound(entries, 1) To UBound(entries, 1)
        wsLib.Cells(i + 2, 1).Value = entries(i, 0)
        wsLib.Cells(i + 2, 2).Value = entries(i, 1)
        wsLib.Cells(i + 2, 3).Value = entries(i, 2)
    Next i

    ' 5. Автоширина столбцов
    wsLib.Columns("A:C").AutoFit

    Call Mod_Logger.WriteLog("Mod_Constants", "InitLibName: Заполнено " & _
        (UBound(entries, 1) - LBound(entries, 1) + 1) & " записей")
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_Constants", "InitLibName: Ошибка — " & Err.Description)
    MsgBox "Ошибка при заполнении libname: " & Err.Description, vbCritical, "Ошибка"
End Sub

' ============================================================
' BuildEntryArray
' Возвращает двумерный массив записей для заполнения листа libname.
' Столбцы: 0 — NameKey, 1 — England, 2 — Russian.
' Использует строковые константы модуля.
' ============================================================
Private Function BuildEntryArray() As Variant
    Dim arr(0 To 14, 0 To 2) As Variant

    ' --- Записи для листа spisok ---

    ' spisok_col_num — лист spisok с входящим списком авто
    arr(0, 0) = "spisok_col_num"
    arr(0, 1) = SPISOK_COL_NUM_NAME
    arr(0, 2) = "лист spisok с входящим списком авто"

    ' spisok_col_model — Модель (столбец B листа spisok)
    arr(1, 0) = "spisok_col_model"
    arr(1, 1) = SPISOK_COL_MODEL_NAME
    arr(1, 2) = "Модель (столбец B листа spisok)"

    ' spisok_col_grz — ГРЗ (столбец C листа spisok)
    arr(2, 0) = "spisok_col_grz"
    arr(2, 1) = SPISOK_COL_GRZ_NAME
    arr(2, 2) = "ГРЗ (столбец C листа spisok)"

    ' spisok_col_vin — VIN (столбец D листа spisok)
    arr(3, 0) = "spisok_col_vin"
    arr(3, 1) = SPISOK_COL_VIN_NAME
    arr(3, 2) = "VIN (столбец D листа spisok)"

    ' spisok_col_garnum — Гараж. № (столбец E листа spisok)
    arr(4, 0) = "spisok_col_garnum"
    arr(4, 1) = SPISOK_COL_GARAGE_NAME
    arr(4, 2) = "Гараж. № (столбец E листа spisok)"

    ' spisok_col_year — Год выпуска (столбец F листа spisok)
    arr(5, 0) = "spisok_col_year"
    arr(5, 1) = SPISOK_COL_YEAR_NAME
    arr(5, 2) = "Год выпуска (столбец F листа spisok)"

    ' spisok_col_mileage — Пробег (столбец G листа spisok)
    arr(6, 0) = "spisok_col_mileage"
    arr(6, 1) = SPISOK_COL_MILEAGE_NAME
    arr(6, 2) = "Пробег (столбец G листа spisok)"

    ' spisok_col_date — Дата (столбец H листа spisok)
    arr(7, 0) = "spisok_col_date"
    arr(7, 1) = SPISOK_COL_DATE_NAME
    arr(7, 2) = "Дата (столбец H листа spisok)"

    ' spisok_col_group — Группа (столбец I листа spisok)
    arr(8, 0) = "spisok_col_group"
    arr(8, 1) = SPISOK_COL_GROUP_NAME
    arr(8, 2) = "Группа (столбец I листа spisok)"

    ' spisok_col_note — РЕЗЕРВ (столбец J листа spisok)
    arr(9, 0) = "spisok_col_note"
    arr(9, 1) = SPISOK_COL_NOTE_NAME
    arr(9, 2) = "РЕЗЕРВ (столбец J листа spisok)"

    ' --- Записи для листа models ---

    ' models_col_model_name — Модель (столбец A листа models)
    arr(10, 0) = "models_col_model_name"
    arr(10, 1) = MODELS_COL_MODEL_NAME
    arr(10, 2) = "Модель (столбец A листа models)"

    ' models_col_group — Группа (столбец B листа models)
    arr(11, 0) = "models_col_group"
    arr(11, 1) = MODELS_COL_GROUP_NAME
    arr(11, 2) = "Группа (столбец B листа models)"

    ' models_col_hrpr — Цена н/ч (столбец C листа models)
    arr(12, 0) = "models_col_hrpr"
    arr(12, 1) = MODELS_COL_PRICE_NAME
    arr(12, 2) = "Цена н/ч (столбец C листа models)"

    ' --- Запись для файла work.xlsm (глобальная, без привязки к листу) ---

    ' work.xlsm — книга Excel с макросами
    arr(13, 0) = "work.xlsm"
    arr(13, 1) = WORK_NAME
    arr(13, 2) = "книга Excel с макросами (возм. разговорн - ""ворк"")"

    ' --- Запись для листа z4 (запчасти) ---

    ' z4 — лист запчастей
    arr(14, 0) = "z4"
    arr(14, 1) = Z4_NAME
    arr(14, 2) = "лист z4 для работы с запчастями"

    BuildEntryArray = arr
End Function

' ============================================================
' AddWorkEntry
' Добавляет запись для work.xlsm в конец списка на листе libname.
' ============================================================
Public Sub AddWorkEntry()
    On Error GoTo ErrHandler

    Dim wsLib As Worksheet
    Dim lastRow As Long

    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, "libname")
    If wsLib Is Nothing Then
        Exit Sub
    End If

    ' Поиск последней заполненной строки
    lastRow = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row

    ' Проверка, что запись work.xlsm ещё не добавлена
    If lastRow >= 2 Then
        Dim checkVal As String
        checkVal = Trim(CStr(wsLib.Cells(lastRow, 1).Value))
        If checkVal = "work.xlsm" Then
            Exit Sub
        End If
    End If

    ' Добавление в следующую строку
    wsLib.Cells(lastRow + 1, 1).Value = "work.xlsm"
    wsLib.Cells(lastRow + 1, 2).Value = WORK_NAME
    wsLib.Cells(lastRow + 1, 3).Value = "книга Excel с макросами"

    wsLib.Columns("A:C").AutoFit

    Call Mod_Logger.WriteLog("Mod_Constants", "AddWorkEntry: Добавлена запись work.xlsm")
    Exit Sub

ErrHandler:
    Call Mod_Logger.WriteLog("Mod_Constants", "AddWorkEntry: Ошибка — " & Err.Description)
End Sub
