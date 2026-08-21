Attribute VB_Name = "Mod_Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Private Module

' ============================================================
' Модуль: Mod_Constants
' Назначение: Централизованное хранение констант проекта
'             и управление листом libname (реестр имён)
' ============================================================

' ============================================================
' Версия приложения (единый источник для всей системы)
' ============================================================
Public Const APP_VERSION As String = "1.0.5"

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
' Константы путей
' ============================================================
' Директория для логов (относительно книги)
Public Const LOGS_DIR As String = "logs"

' ============================================================
' Константы агрегатов (категорий работ/запчастей)
' Соглашение: AGG_{КОД} — трёхбуквенный код агрегата
' ============================================================
Public Const AGG_DIAG As String = "DIAG"     ' Диагностика
Public Const AGG_TO As String = "TO"         ' ТО (Maintenance/Service)
Public Const AGG_ENG As String = "ENG"       ' Двигатель (Engine)
Public Const AGG_TRANS As String = "TRANS"   ' Трансмиссия (Transmission)
Public Const AGG_CLUTCH As String = "CLUTCH" ' Сцепление (Clutch)
Public Const AGG_SUSP As String = "SUSP"     ' Подвеска (Suspension)
Public Const AGG_CHASS As String = "CHASS"   ' Ходовая (Chassis)
Public Const AGG_BRAKE As String = "BRAKE"   ' Тормозная система (Brake system)
Public Const AGG_FUEL As String = "FUEL"     ' Топливная система (Fuel system)
Public Const AGG_COOL As String = "COOL"     ' Система охлаждения (Cooling system)
Public Const AGG_HVAC As String = "HVAC"     ' Система обогрева и кондиционирования (HVAC)
Public Const AGG_STEER As String = "STEER"   ' Система рулевого управления (Steering system)
Public Const AGG_ELEC As String = "ELEC"     ' Электрооборудование (Electrical system)
Public Const AGG_EXH As String = "EXH"       ' Система выхлопных газов (Exhaust system)
Public Const AGG_BODY As String = "BODY"     ' Кузов (Body)
Public Const AGG_OTHERS As String = "OTHERS" ' Прочие работы (Others)

' ============================================================
' Константы строк листа main
' ============================================================
Public Const MAIN_HEADER_START_ROW As Long = 4   ' B4 — номер заказа (ввод пользователя)
Public Const MAIN_DATA_START_ROW As Long = 4      ' Строка, с которой начинаются таблицы работ/материалов

' ============================================================
' Константы имён листов
' ============================================================
Public Const SHEET_MAIN As String = "main"
Public Const SHEET_SPISOK As String = "spisok"
Public Const SHEET_MODELS As String = "models"
Public Const SHEET_LIBNAME As String = "libname"
Public Const SHEET_REPORT As String = "report"

' ============================================================
' Глобальный флаг подавления MsgBox (для тестового режима)
' ============================================================
Public SilenceMsgBox As Boolean

' ============================================================
' Флаг глубокой подстановки модельных кодов при импорте (проблема 2)
' ============================================================
' True  — в ImportDataToMain выполняется подстановка артикулов из
'         matlib_entries: для работ — target_code в колонку O(15),
'         для запчастей — в колонку AB(28) (только при точном совпадении).
' False — поведение импорта идентично прежнему (по умолчанию; TC-29 зелёный).
' ============================================================
Public ApplyMatLibSubstitution As Boolean

' ============================================================
' Константы колонок ручного подбора работ (E4:H)
' ============================================================
Public Const MANWRK_COL_ARTICLE As Long = 5    ' E — Артикул
Public Const MANWRK_COL_NAME As Long = 6       ' F — Наименование
Public Const MANWRK_COL_NORMHOURS As Long = 7  ' G — Кол-во н/ч
Public Const MANWRK_COL_QTY As Long = 8        ' H — Кол-во оп
Public Const MANWRK_START_ROW As Long = 4      ' Строка начала данных

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
    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_LIBNAME)
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
    ' 15 существующих + 16 агрегатов = 31 запись
    Dim arr(0 To 30, 0 To 2) As Variant

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

    ' --- Записи для агрегатов (категорий работ/запчастей) ---
    ' Соглашение: agg_{код} — трёхбуквенный код агрегата

    ' agg_diag — Диагностика
    arr(15, 0) = "agg_diag"
    arr(15, 1) = AGG_DIAG
    arr(15, 2) = "Диагностика (Diagnostics)"

    ' agg_to — ТО
    arr(16, 0) = "agg_to"
    arr(16, 1) = AGG_TO
    arr(16, 2) = "ТО (Maintenance/Service)"

    ' agg_eng — Двигатель
    arr(17, 0) = "agg_eng"
    arr(17, 1) = AGG_ENG
    arr(17, 2) = "Двигатель (Engine)"

    ' agg_trans — Трансмиссия
    arr(18, 0) = "agg_trans"
    arr(18, 1) = AGG_TRANS
    arr(18, 2) = "Трансмиссия (Transmission)"

    ' agg_clutch — Сцепление
    arr(19, 0) = "agg_clutch"
    arr(19, 1) = AGG_CLUTCH
    arr(19, 2) = "Сцепление (Clutch)"

    ' agg_susp — Подвеска
    arr(20, 0) = "agg_susp"
    arr(20, 1) = AGG_SUSP
    arr(20, 2) = "Подвеска (Suspension)"

    ' agg_chass — Ходовая
    arr(21, 0) = "agg_chass"
    arr(21, 1) = AGG_CHASS
    arr(21, 2) = "Ходовая (Chassis)"

    ' agg_brake — Тормозная система
    arr(22, 0) = "agg_brake"
    arr(22, 1) = AGG_BRAKE
    arr(22, 2) = "Тормозная система (Brake system)"

    ' agg_fuel — Топливная система
    arr(23, 0) = "agg_fuel"
    arr(23, 1) = AGG_FUEL
    arr(23, 2) = "Топливная система (Fuel system)"

    ' agg_cool — Система охлаждения
    arr(24, 0) = "agg_cool"
    arr(24, 1) = AGG_COOL
    arr(24, 2) = "Система охлаждения (Cooling system)"

    ' agg_hvac — Система обогрева и кондиционирования
    arr(25, 0) = "agg_hvac"
    arr(25, 1) = AGG_HVAC
    arr(25, 2) = "Система обогрева и кондиционирования (HVAC)"

    ' agg_steer — Система рулевого управления
    arr(26, 0) = "agg_steer"
    arr(26, 1) = AGG_STEER
    arr(26, 2) = "Система рулевого управления (Steering system)"

    ' agg_elec — Электрооборудование
    arr(27, 0) = "agg_elec"
    arr(27, 1) = AGG_ELEC
    arr(27, 2) = "Электрооборудование (Electrical system)"

    ' agg_exh — Система выхлопных газов
    arr(28, 0) = "agg_exh"
    arr(28, 1) = AGG_EXH
    arr(28, 2) = "Система выхлопных газов (Exhaust system)"

    ' agg_body — Кузов
    arr(29, 0) = "agg_body"
    arr(29, 1) = AGG_BODY
    arr(29, 2) = "Кузов (Body)"

    ' agg_others — Прочие работы
    arr(30, 0) = "agg_others"
    arr(30, 1) = AGG_OTHERS
    arr(30, 2) = "Прочие работы (Others)"

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

    Set wsLib = Mod_Utils.GetSheetByName(ThisWorkbook, Mod_Constants.SHEET_LIBNAME)
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

' ============================================================
' GetAggregateName
' Возвращает русское название агрегата по его коду.
' Если код не найден — возвращает пустую строку.
' ============================================================
Public Function GetAggregateName(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case AGG_DIAG:   GetAggregateName = "Диагностика"
        Case AGG_TO:     GetAggregateName = "ТО"
        Case AGG_ENG:    GetAggregateName = "Двигатель"
        Case AGG_TRANS:  GetAggregateName = "Трансмиссия"
        Case AGG_CLUTCH: GetAggregateName = "Сцепление"
        Case AGG_SUSP:   GetAggregateName = "Подвеска"
        Case AGG_CHASS:  GetAggregateName = "Ходовая"
        Case AGG_BRAKE:  GetAggregateName = "Тормозная система"
        Case AGG_FUEL:   GetAggregateName = "Топливная система"
        Case AGG_COOL:   GetAggregateName = "Система охлаждения"
        Case AGG_HVAC:   GetAggregateName = "Система обогрева и кондиционирования"
        Case AGG_STEER:  GetAggregateName = "Система рулевого управления"
        Case AGG_ELEC:   GetAggregateName = "Электрооборудование"
        Case AGG_EXH:    GetAggregateName = "Система выхлопных газов"
        Case AGG_BODY:   GetAggregateName = "Кузов"
        Case AGG_OTHERS: GetAggregateName = "Прочие работы"
        Case Else:       GetAggregateName = ""
    End Select
End Function

' ============================================================
' Константы выбора провайдера данных моделей
' ============================================================
' True = SQLite (основной провайдер, SysW.db через ADO/ODBC),
' False = Excel (резервный провайдер, легаси-файлы base/models/*.xlsm).
' Фабрика Mod_ModelDB.GetModelDataProvider() учитывает этот флаг;
' при True, но недоступных БД/драйвере — автоматический fallback на Excel.
'
' ВАЖНО: флаг ранее объявлялся как Public Const / Public-переменная, однако
' модульные объявления идентификаторов в этом модуле не резолвятся (ошибка
' компиляции 461 / Variable not defined), тогда как вызовы функций работают.
' Поэтому флаг возвращается функцией без обращения к модульным переменным.
' ============================================================

' --------------------------------------------------------------------------
' SqliteProviderEnabled — возвращает значение флага выбора провайдера данных.
' True — SQLite (основной провайдер); для принудительного использования
' Excel-провайдера измените возвращаемое значение на False.
' --------------------------------------------------------------------------
Public Function SqliteProviderEnabled() As Boolean
    SqliteProviderEnabled = True
End Function
