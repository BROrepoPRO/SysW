Attribute VB_Name = "Mod_PickWork"
Option Explicit
Option Private Module

' ============================================================
' Модуль: Mod_PickWork
' Назначение: Ручной подбор работ из справочника группы
'             Открывает файл группы, активирует лист работ,
'             пользователь ищет и копирует данные вручную
' ============================================================

' ============================================================
' Вспомогательные функции
' ============================================================

' --------------------------------------------------------------------------
' GetGroupNameFromMain
' Читает название группы из ячейки B14 листа main.
' Сохранена как Public-обёртка для обратной совместимости (вызывается TC-36);
' фактическое чтение делегируется единому хелперу Mod_Utils.GetGroupName.
' --------------------------------------------------------------------------
Public Function GetGroupNameFromMain() As String
    GetGroupNameFromMain = Mod_Utils.GetGroupName()
End Function

' --------------------------------------------------------------------------
' GetWorkSheetName
' Возвращает имя листа работ в файле группы.
' Имя листа совпадает с именем группы.
' --------------------------------------------------------------------------
Public Function GetWorkSheetName(ByVal groupName As String) As String
    GetWorkSheetName = groupName
End Function

' --------------------------------------------------------------------------
' GetPartsSheetName
' Возвращает имя модельного листа запчастей группы: "{Group}z4".
' Если такой лист отсутствует в книге, используется общий лист "z4".
' --------------------------------------------------------------------------
Public Function GetPartsSheetName(ByVal wb As Workbook, ByVal groupName As String) As String
    Dim modelName As String
    modelName = Trim(groupName) & "z4"

    If Not Mod_Utils.GetSheetByName(wb, modelName) Is Nothing Then
        GetPartsSheetName = modelName
    ElseIf Not Mod_Utils.GetSheetByName(wb, "z4") Is Nothing Then
        GetPartsSheetName = "z4"
    Else
        GetPartsSheetName = ""
    End If
End Function

' ============================================================
' Основные процедуры
' ============================================================

' --------------------------------------------------------------------------
' PickWork_UI
' Главная точка входа для ручного подбора работ.
' 1. Читает группу из B14
' 2. Открывает файл группы через Mod_ModelDB
' 3. Активирует лист работ
' 4. Показывает инструкцию пользователю
' --------------------------------------------------------------------------
Public Sub PickWork_UI()
    On Error GoTo ErrHandler

    Dim groupName As String
    Dim wbGroup As Workbook
    Dim wsWork As Worksheet
    Dim sheetName As String
    Dim msg As String

    ' 1. Отключаем обновление экрана
    Application.ScreenUpdating = False

    ' 2. Получаем группу из B14
    groupName = GetGroupNameFromMain()
    If groupName = "" Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Группа не указана. Заполните ячейку B14 на листе main.", _
                   vbExclamation, "Ручной подбор работ"
        End If
        GoTo CleanUp
    End If

    ' 3. Открываем файл группы
    Set wbGroup = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wbGroup Is Nothing Then
        msg = "Файл группы '" & groupName & "' не найден." & vbCrLf & vbCrLf & _
              "Ожидаемый путь: " & Mod_ModelDB.GetModelGroupFilePath(groupName) & vbCrLf & vbCrLf & _
              "Убедитесь, что файл существует в каталоге base\models\"
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox msg, vbExclamation, "Ручной подбор работ"
        End If
        GoTo CleanUp
    End If

    ' 4. Получаем имя листа работ
    sheetName = GetWorkSheetName(groupName)

    ' 5. Проверяем существование листа
    On Error Resume Next
    Set wsWork = wbGroup.Sheets(sheetName)
    On Error GoTo ErrHandler

    If wsWork Is Nothing Then
        msg = "Лист '" & sheetName & "' не найден в файле группы '" & groupName & ".xlsx'." & vbCrLf & _
              "Убедитесь, что лист с именем группы существует."
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox msg, vbExclamation, "Ручной подбор работ"
        End If
        GoTo CleanUp
    End If

    ' 6. Активируем лист работ (только в видимом режиме)
    If Not Mod_Constants.SilenceMsgBox Then
        wsWork.Activate
        wsWork.Select
    End If

    ' 7. Показываем инструкцию
    msg = "Файл группы '" & groupName & "' открыт." & vbCrLf & vbCrLf & _
          "Инструкция:" & vbCrLf & _
          "1. В ячейку C1 введите текст для поиска" & vbCrLf & _
          "2. Нажмите кнопку поиска по артикулу или наименованию" & vbCrLf & _
          "3. В столбце G (Кол-во ЗН) проставьте количество" & vbCrLf & _
          "4. Отфильтруйте по столбцу G — всё кроме 0" & vbCrLf & _
          "5. Скопируйте отфильтрованные строки" & vbCrLf & _
          "6. Вставьте на лист main в диапазон E4:H" & vbCrLf & vbCrLf & _
          "Внимание: строки 1-2 на main не заполнять!"

    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox msg, vbInformation, "Ручной подбор работ"
    End If

    GoTo CleanUp

ErrHandler:
    ' Восстановление состояния приложения
    Application.ScreenUpdating = True
    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox "Ошибка при ручном подборе работ: " & Err.Description, vbCritical, "Ошибка"
    End If
    Call Mod_Logger.WriteLog("Mod_PickWork", "PickWork_UI: " & Err.Description)
    GoTo CleanUp

CleanUp:
    ' Закрываем файл группы, если он был открыт
    If Not wbGroup Is Nothing Then
        On Error Resume Next
        wbGroup.Close SaveChanges:=False
        Set wbGroup = Nothing
        On Error GoTo 0
    End If
    Application.ScreenUpdating = True
End Sub

' ============================================================
' Ручной подбор запчастей
' ============================================================

' --------------------------------------------------------------------------
' PickParts_UI
' Ручной подбор запчастей (зеркало PickWork_UI для листа запчастей).
' 1. Читает группу из B14
' 2. Открывает файл группы через Mod_ModelDB
' 3. Активирует лист запчастей ({Group}z4 или z4)
' 4. Показывает инструкцию пользователю
' --------------------------------------------------------------------------
Public Sub PickParts_UI()
    On Error GoTo ErrHandler

    Dim groupName As String
    Dim wbGroup As Workbook
    Dim wsParts As Worksheet
    Dim sheetName As String
    Dim msg As String

    ' 1. Отключаем обновление экрана
    Application.ScreenUpdating = False

    ' 2. Получаем группу из B14
    groupName = GetGroupNameFromMain()
    If groupName = "" Then
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox "Группа не указана. Заполните ячейку B14 на листе main.", _
                   vbExclamation, "Ручной подбор запчастей"
        End If
        GoTo CleanUp
    End If

    ' 3. Открываем файл группы
    Set wbGroup = Mod_ModelDB.OpenModelGroupFile(groupName)
    If wbGroup Is Nothing Then
        msg = "Файл группы '" & groupName & "' не найден." & vbCrLf & vbCrLf & _
              "Ожидаемый путь: " & Mod_ModelDB.GetModelGroupFilePath(groupName) & vbCrLf & vbCrLf & _
              "Убедитесь, что файл существует в каталоге base\models\"
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox msg, vbExclamation, "Ручной подбор запчастей"
        End If
        GoTo CleanUp
    End If

    ' 4. Определяем лист запчастей ({Group}z4, при отсутствии — общий z4)
    sheetName = GetPartsSheetName(wbGroup, groupName)

    If sheetName = "" Then
        msg = "Лист запчастей (z4 / " & groupName & "z4) не найден в файле группы '" & groupName & "'."
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox msg, vbExclamation, "Ручной подбор запчастей"
        End If
        GoTo CleanUp
    End If

    ' 5. Проверяем существование листа
    On Error Resume Next
    Set wsParts = wbGroup.Sheets(sheetName)
    On Error GoTo ErrHandler

    If wsParts Is Nothing Then
        msg = "Лист '" & sheetName & "' не найден в файле группы '" & groupName & "'."
        If Not Mod_Constants.SilenceMsgBox Then
            MsgBox msg, vbExclamation, "Ручной подбор запчастей"
        End If
        GoTo CleanUp
    End If

    ' 6. Активируем лист запчастей (только в видимом режиме)
    If Not Mod_Constants.SilenceMsgBox Then
        wsParts.Activate
        wsParts.Select
    End If

    ' 7. Показываем инструкцию
    msg = "Файл группы '" & groupName & "' открыт. Лист запчастей: '" & sheetName & "'." & vbCrLf & vbCrLf & _
          "Инструкция:" & vbCrLf & _
          "1. В ячейку C1 введите текст для поиска" & vbCrLf & _
          "2. Нажмите кнопку поиска по артикулу или наименованию" & vbCrLf & _
          "3. Скопируйте нужные строки запчастей" & vbCrLf & _
          "4. Вставьте на лист main в область запчастей (P:W / X:AB)" & vbCrLf & vbCrLf & _
          "Внимание: строки 1-2 на main не заполнять!"

    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox msg, vbInformation, "Ручной подбор запчастей"
    End If

    GoTo CleanUp

ErrHandler:
    ' Восстановление состояния приложения
    Application.ScreenUpdating = True
    If Not Mod_Constants.SilenceMsgBox Then
        MsgBox "Ошибка при ручном подборе запчастей: " & Err.Description, vbCritical, "Ошибка"
    End If
    Call Mod_Logger.WriteLog("Mod_PickWork", "PickParts_UI: " & Err.Description)
    GoTo CleanUp

CleanUp:
    ' Закрываем файл группы, если он был открыт
    If Not wbGroup Is Nothing Then
        On Error Resume Next
        wbGroup.Close SaveChanges:=False
        Set wbGroup = Nothing
        On Error GoTo 0
    End If
    Application.ScreenUpdating = True
End Sub