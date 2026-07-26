Attribute VB_Name = "Mod_MainButtons"
Option Explicit
Option Private Module

' ============================================================
' Модуль: Mod_MainButtons
' Назначение: Обработчики кнопок листа main
' Отдельный модуль, чтобы не смешивать с Mod_ButtonDispatcher
' ============================================================


' ============================================================
' ИМПОРТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_Import
' Импорт из report.xlsx по ГРЗ из ячейки B6 листа main.
' Делегирует выполнение Mod_Import.ImportSheet_UI.
' --------------------------------------------------------------------------
Public Sub Btn_main_Import()
    On Error GoTo ErrHandler

    Call Mod_Import.ImportSheet_UI

    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при импорте: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Utils.WriteLog("Btn_main_Import: " & Err.Description)
End Sub

' ============================================================
' АВТОПОДБОР ЗАПЧАСТЕЙ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_AUTOz4
' Автоподбор запчастей из тождеств UAZz4
' --------------------------------------------------------------------------
Public Sub Btn_main_AUTOz4()
    On Error GoTo ErrHandler

    Call Mod_AutoMatch.AutoMatchParts

    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при автоподборе запчастей: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_MainButtons", "Btn_main_AUTOz4: " & Err.Description)
End Sub

' ============================================================
' АВТОПОДБОР РАБОТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_AUTOw
' Автоподбор работ из тождеств UAZw
' --------------------------------------------------------------------------
Public Sub Btn_main_AUTOw()
    On Error GoTo ErrHandler

    Call Mod_AutoMatch.AutoMatchWorks

    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при автоподборе работ: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_MainButtons", "Btn_main_AUTOw: " & Err.Description)
End Sub

' ============================================================
' ИМПОРТ ПО VIN/ГРЗ (ДЕЛЕГИРОВАНО)
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_ImportVH_Click
' Обработчик кнопки импорта по VIN/ГРЗ.
' Делегировано Mod_ButtonDispatcher в v0.13.0.
' --------------------------------------------------------------------------
Public Sub Btn_main_ImportVH_Click()
    ' Делегировано Mod_ButtonDispatcher в v0.13.0
    Mod_ButtonDispatcher.Btn_main_ImportVH_Click
End Sub

' ============================================================
' РУЧНОЙ ПОДБОР РАБОТ
' ============================================================

' --------------------------------------------------------------------------
' Btn_main_MANWRK
' Ручной подбор работ из справочника группы.
' Открывает файл группы, пользователь ищет и копирует данные вручную.
' --------------------------------------------------------------------------
Public Sub Btn_main_MANWRK()
    On Error GoTo ErrHandler

    Call Mod_PickWork.PickWork_UI

    Exit Sub

ErrHandler:
    ' Восстановление состояния приложения
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка при ручном подборе работ: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_MainButtons", "Btn_main_MANWRK: " & Err.Description)
End Sub
