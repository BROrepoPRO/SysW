# Отчёт о функциях проекта SysW

## Mod_OrderHeader.bas
- `Public Sub FillHeaderFromOrder(ByVal OrderNum As String, ByVal wsMain As Worksheet, ByVal wsSpisok As Worksheet, ByVal wsModel As Worksheet)` — заполняет ячейки B3:B15 на листе main данными из листов spisok и model по номеру заказа. Принимает 4 параметра (номер заказа + 3 листа)
- `Public Function FindOrder(ByVal OrderNum As String, ByRef Header As OrderHeader) As Boolean` — ищет заказ по номеру на листе spisok, заполняет Header (UDT). Возвращает Boolean (True/False)

## Mod_Import.bas
- `Public Function ExtractNumberFromGRZ(ByVal GRZ As String) As String` — извлекает цифры из ГРЗ (госномера)
- `Public Sub RenameSheetsByGRZ()` — переименовывает листы (кроме main, spisok, model) по госномеру из ячейки B2 в формат "GRZ_<цифры>"
- `Public Function SearchSheetByGRZ(ByVal GRZ As String) As Worksheet` — ищет лист по госномеру в имени листа (по цифрам)
- `Public Sub ImportSheet(ByVal GRZ As String)` — импорт данных с листа, найденного по госномеру, через ImportDataToMain
- `Public Sub ImportIncomingDocNumber(ByVal GRZ As String)` — импорт входящего номера документа (столбец A) в столбец W листа main
- `Public Sub ImportDataToMain(ByVal wsSource As Worksheet)` — импорт с маппингом столбцов: C→L, D→M, H→N, B→X, C→Y, D→Z, G→AA
- `Public Sub ImportFromReport()` — импорт из листа report (столбцы A, B, C) в лист main (столбцы A, B, C)
- `Public Sub ClearAll()` — **НЕ СУЩЕСТВУЕТ** (отсутствует в модуле)

## Mod_Utils.bas
- `Public Type OrderHeader` — UDT: OrderNumber, ClientName, CarModel, CarPlate, Mileage (Long), DateIn (Date), DateOut (Date), Status (String) — 8 полей
- `Public Function GetSheetByName(ByVal wb As Workbook, ByVal SheetName As String) As Worksheet` — получить лист по имени (без ошибки если нет)
- `Public Function GetWorkbookPath() As String` — путь к текущей книге
- `Public Function FileExists(ByVal FilePath As String) As Boolean` — проверка существования файла через Dir()
- `Public Function GetCurrentUser() As String` — имя текущего пользователя (Environ("USERNAME"))
- `Public Function FormatDateSQL(ByVal d As Date) As String` — форматирование даты в формат SQLite (ГГГГ-ММ-ДД)
- `Public Sub WriteLog(ByVal Message As String)` — запись в лог-файл log.txt

## Mod_ButtonDispatcher.bas
- `Public Sub Btn_main_Import_Click()` — вызывает Mod_Import.RunImport (**НО RunImport не существует!**)
- `Public Sub Btn_main_Test_Click()` — вызывает Mod_MinimalTestRunner.RunAllTests
- `Public Sub Btn_main_Clear_Click()` — очищает содержимое активного листа (с подтверждением)
- `Public Sub Btn_main_Refresh_Click()` — заглушка (MsgBox "Обновление данных...")
- `Public Sub Btn_z4_Action_Click()` — заглушка (MsgBox "Функция z4 в разработке.")
- `Public Sub Btn_work_Action_Click()` — заглушка (MsgBox "Функция work в разработке.")

## Sheet1_main.cls
- `Private Sub Worksheet_Change(ByVal Target As Range)` — при изменении ячейки B2 вызывает FillHeaderFromOrder. Содержит защиту от рекурсии (Static IsProcessing) и проверку Application.EnableEvents

## Mod_MinimalTestRunner.bas (текущий)
- `Public Sub RunAllTests()` — запускает тесты TC-01..TC-09
- TC-01: FileExists с существующим файлом (C:\Windows\notepad.exe)
- TC-02: FileExists с несуществующим файлом
- TC-03: FormatDateSQL корректной даты (2026-07-12 → "2026-07-12")
- TC-04: FormatDateSQL с пустой датой (0 → "1899-12-30")
- TC-05: ExtractNumberFromGRZ ("А123ВВ77" → "12377")
- TC-06: GetSheetByName существующий лист ("main")
- TC-07: GetSheetByName несуществующий лист
- TC-08: FindOrder с существующим заказом ("ЗЗ-001") — **ПРОВАЛЕН**
- TC-09: FindOrder с несуществующим заказом

## Проблемы
1. **Текущие тесты покрывают только 5 из 16 Public функций (31%)** — тестируются: FileExists, FormatDateSQL, ExtractNumberFromGRZ, GetSheetByName, FindOrder. Не тестируются: FillHeaderFromOrder, RenameSheetsByGRZ, SearchSheetByGRZ, ImportSheet, ImportIncomingDocNumber, ImportDataToMain, ImportFromReport, GetWorkbookPath, GetCurrentUser, WriteLog, RunAllTests
2. **Btn_main_Import_Click вызывает несуществующую Mod_Import.RunImport** — процедура RunImport отсутствует в Mod_Import.bas
3. **TC-08 провален** — FindOrder("ЗЗ-001", Header) не находит заказ, т.к. тест запускается в среде без реального листа spisok с данными
4. **Btn_main_ClearHeader_Click не существует** — в отчёте указана, но в коде отсутствует
5. **Mod_Import.ClearAll не существует** — в отчёте указана, но в коде отсутствует
6. **FillHeaderFromOrder принимает 4 параметра** (а не 1, как указано в отчёте)
7. **FindOrder возвращает Boolean** (а не OrderHeader, как указано в отчёте)
8. **OrderHeader UDT содержит 8 полей** (а не 6, как указано в отчёте)