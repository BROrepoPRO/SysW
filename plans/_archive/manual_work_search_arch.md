# Архитектура ручного поиска работ

## 1. Итоги опроса (Ask)

| Параметр | Решение |
|----------|---------|
| Диапазон на main | **E4:H** (данные с 4-й строки) |
| Колонки E4:H | E=Артикул, F=Наименование, G=Кол-во н/ч, H=Кол-во оп |
| Интерфейс | Открытие файла группы + фильтр (скрипт поиска) |
| Процесс | Поиск → проставить G(Кол-во ЗН) → фильтр по G → вручную скопировать в E4:H |
| Структура файла группы | A=№п/п, B=Артикул, C=Наименование, D=н/ч, E=кол-во оп(константа), F=цена н/ч, G=Кол-во ЗН, H=Сумма ЗН |
| Mod_ModelDB | Минимальный: `OpenModelGroupFile` + `GetWorks` |
| Модуль подбора | `Mod_PickWork.bas` |
| Перенос на main | Вручную (автоматизация в будущем) |
| Имя макроса кнопки | `Btn_main_MANWRK` |

## 2. Общая схема работы

```mermaid
flowchart TD
    A[Пользователь нажимает РУЧ РАБ] --> B[Btn_main_MANWRK в Mod_MainButtons]
    B --> C[Mod_PickWork.PickWork_UI]
    C --> D[Читает B14 = группа]
    D --> E{Файл группы существует?}
    E -->|Да| F[Mod_ModelDB.OpenModelGroupFile]
    E -->|Нет| G[MsgBox: Файл группы не найден]
    F --> H[Активирует лист {GroupName}]
    H --> I[Пользователь ищет работы]
    I --> J[Вводит поиск в C1]
    J --> K[Нажимает Поиск по артикулу/B или наименованию/C]
    K --> L[AutoFilter по столбцу]
    L --> M[Проставляет G Кол-во ЗН]
    M --> N[Фильтрует по G все кроме 0]
    N --> O[Вручную копирует в E4:H на main]
```

## 3. Модули и файлы

### 3.1. Новый файл: `src/modules/Mod_ModelDB.bas`

Минимальная версия модуля доступа к файлам модельных групп.

**Константы:**
```vba
Public Const MODELDB_BASE_PATH As String = "L:\PROject\SysW\base\models\"
```

**Функции:**

1. **`OpenModelGroupFile(groupName As String) As Workbook`**
   - Проверяет, не открыт ли уже файл `{groupName}.xlsx`
   - Если открыт — возвращает ссылку
   - Если нет — проверяет существование через `Dir()`
   - Открывает `ReadOnly:=False`, возвращает ссылку
   - Если файл не найден — возвращает `Nothing`

2. **`GetWorks(groupName As String, filters As Variant) As Collection`**
   - Открывает файл группы через `OpenModelGroupFile`
   - Читает данные листа `{groupName}` (данные с 4-й строки)
   - Применяет фильтры (поддержка wildcard `*`)
   - Возвращает коллекцию `WorkEntry`
   - **На данном этапе — заглушка для будущего использования**

**Структура WorkEntry:**
```vba
Public Type WorkEntry
    Code As String
    Name As String
    Unit As String
    NormHours As Double
    Price As Currency
    Note As String
End Type
```

**Вспомогательные:**
- `GetModelGroupFilePath(groupName As String) As String` — полный путь к файлу
- `ModelGroupFileExists(groupName As String) As Boolean` — проверка существования

### 3.2. Новый файл: `src/modules/Mod_PickWork.bas`

Модуль логики ручного подбора работ.

**Функции:**

1. **`PickWork_UI()`** — главная точка входа
   - Читает группу из `main.B14`
   - Вызывает `Mod_ModelDB.OpenModelGroupFile`
   - Если файла нет — сообщение пользователю
   - Если файл открыт — активирует лист `{GroupName}`
   - Показывает инструкцию пользователю (MsgBox)

2. **`GetGroupNameFromMain() As String`** — читает группу из B14
   - `Trim(wsMain.Range("B14").Value)`

3. **`GetWorkSheetName(groupName As String) As String`** — имя листа работ
   - Возвращает `groupName` (имя листа совпадает с именем группы)

### 3.3. Изменения в `src/modules/Mod_MainButtons.bas`

**Замена заглушки `Btn_main_MANw()` на `Btn_main_MANWRK()`:**
```vba
Public Sub Btn_main_MANWRK()
    On Error GoTo ErrHandler
    Call Mod_PickWork.PickWork_UI
    Exit Sub
ErrHandler:
    MsgBox "Ошибка при ручном подборе работ: " & Err.Description, vbCritical, "Ошибка"
    Call Mod_Logger.WriteLog("Mod_MainButtons", "Btn_main_MANWRK: " & Err.Description)
End Sub
```

### 3.4. Изменения в `src/modules/Mod_Constants.bas`

Добавить константы для новых колонок E4:H:
```vba
' Колонки ручного подбора работ (E4:H)
Public Const MANWRK_COL_ARTICLE As Long = 5    ' E — Артикул
Public Const MANWRK_COL_NAME As Long = 6       ' F — Наименование
Public Const MANWRK_COL_NORMHOURS As Long = 7  ' G — Кол-во н/ч
Public Const MANWRK_COL_QTY As Long = 8        ' H — Кол-во оп
Public Const MANWRK_START_ROW As Long = 4      ' Строка начала данных
```

## 4. Скрипт поиска (встраивается в файлы групп)

Скрипт, который пользователь предоставил, будет встроен в каждый файл группы (`base/models/{Group}.xlsx`) как стандартный модуль поиска. Он не является частью `work.xlsm`, а живёт внутри файлов групп.

**Скрипт включает:**
- `SearchByArticle()` — поиск по столбцу B (Артикул)
- `SearchByName()` — поиск по столбцу C (Наименование)
- `ClearSearchFilter()` — сброс фильтра
- Поле ввода — ячейка C1

**Наша задача в Mod_PickWork:** только открыть файл группы и активировать нужный лист. Поиск — средствами скрипта внутри файла группы.

## 5. Порядок реализации

```mermaid
flowchart LR
    A[1. Mod_Constants] --> B[2. Mod_ModelDB]
    B --> C[3. Mod_PickWork]
    C --> D[4. Mod_MainButtons]
    D --> E[5. Синхронизация]
    E --> F[6. Проверка]
```

### Шаг 1: Mod_Constants
- Добавить константы `MANWRK_COL_ARTICLE`, `MANWRK_COL_NAME`, `MANWRK_COL_NORMHOURS`, `MANWRK_COL_QTY`, `MANWRK_START_ROW`

### Шаг 2: Mod_ModelDB (новый файл)
- Константа `MODELDB_BASE_PATH`
- Тип `WorkEntry`
- Функция `GetModelGroupFilePath`
- Функция `ModelGroupFileExists`
- Функция `OpenModelGroupFile`
- Функция `GetWorks` (заглушка, возвращает пустую коллекцию)

### Шаг 3: Mod_PickWork (новый файл)
- Функция `GetGroupNameFromMain`
- Функция `GetWorkSheetName`
- Процедура `PickWork_UI`

### Шаг 4: Mod_MainButtons
- Заменить заглушку `Btn_main_MANw` на `Btn_main_MANWRK` с вызовом `Mod_PickWork.PickWork_UI`

### Шаг 5: Синхронизация
- `python scripts/impVBA.py` для синхронизации модулей с книгой

### Шаг 6: Проверка
- Открыть `work.xlsm`
- Нажать кнопку РУЧ РАБ
- Убедиться, что открывается файл группы
- Убедиться, что строки 1-2 остаются пустыми
- Убедиться, что существующий импорт не сломан

## 6. Важные правила

- Все операции с ячейками используют константу `MAIN_DATA_START_ROW = 4`
- Строки 1-2 не заполняются
- Имена файлов и переменных — только латиница
- Код модульный, с чёткими функциями
- Обработка ошибок через `On Error GoTo ErrHandler`
- Логирование через `Mod_Logger.WriteLog`