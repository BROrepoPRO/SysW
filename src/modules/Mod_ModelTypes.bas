Attribute VB_Name = "Mod_ModelTypes"
Option Explicit

' =============================================
' Модуль: Mod_ModelTypes
' Назначение: Определения пользовательских типов (UDT)
'
' С v1.0.5 объектные типы вынесены в классы (для единообразного
' возврата объектов из методов IModelDataProvider):
'   - src/classes/WorkEntry.cls     (запись работы; заменяет UDT WorkEntry)
'   - src/classes/WorkIdentity.cls  (тождество работ)
'   - src/classes/PartIdentity.cls  (тождество запчастей)
'
' UDT WorkEntry удалён — используется класс WorkEntry.
' Модуль сохранён для обратной совместимости ссылок.
' =============================================