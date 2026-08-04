# Анализ ошибки компиляции VBA (0x800A03EC)

## Симптом

При запуске тестов (`python scripts/run_tests.py`) возникает ошибка:

```
Ошибка: (-2147352567, 'Ошибка.', (0, None, None, None, 0, -2146788248), None)
```

Код `-2146788248` = `0x800A03EC` — типичная ошибка компиляции VBA:
> "Only public user defined types defined in public object modules can be used as parameters or return types for public procedures of class modules or as fields of public user defined types"

## Исходный код (на диске) — корректен

### 1. `src/modules/Mod_ModelTypes.bas` (НОВЫЙ)
- Строки 17–57: три `Public Type` — `WorkEntry`, `WorkIdentity`, `PartIdentity`
- Размещение в стандартном модуле (`.bas`) — **правильное**
- UDT объявлены как `Public` — **правильно**

### 2. `src/classes/ModelTypes.cls` (ОБНОВЛЁН)
- Строки 1–20: класс **пустой**, содержит только комментарии
- UDT удалены, есть только ссылки-комментарии на `Mod_ModelTypes.*`
- `Attribute VB_Exposed = True` — класс остаётся в проекте

### 3. `src/modules/Mod_ModelDB.bas` (ОБНОВЛЁН)
- Строка 133: `Dim tmpEntry As Mod_ModelTypes.WorkEntry` — **корректно**
- Строка 204: `Dim tmpIdentity As Mod_ModelTypes.WorkIdentity` — **корректно**
- Строка 279: `Dim tmpIdentity As Mod_ModelTypes.PartIdentity` — **корректно**
- Все три функции (`GetWorks`, `GetWorkIdentities`, `GetPartIdentities`) — `Public` в стандартном модуле — **разрешено**

### 4. `src/modules/Mod_AutoMatch.bas`
- Строки 130, 250: вызывает `Mod_ModelDB.GetWorkIdentities` / `GetPartIdentities`
- Использует `Variant` для итерации по `Collection` — **корректно**
- Ни один класс (`.cls`) не использует UDT напрямую

## Причина ошибки

**Проблема НЕ в исходном коде на диске, а в несоответствии между исходным кодом и тем, что реально находится внутри `work.xlsm`.**

### Первопричина: `ModelTypes.cls` не удаляется при импорте

Скрипт импорта [`scripts/impVBA.py`](scripts/impVBA.py) работает в два прохода:

1. **Первый проход** (строки 257–294): удаляет существующие компоненты по `VB_Name`
2. **Второй проход** (строки 299–387): импортирует компоненты заново

**Проблема:** `ModelTypes.cls` — это класс (`.cls`), который:
- Имеет `Attribute VB_Name = "ModelTypes"` (строка 1)
- Имеет `Attribute VB_Exposed = True` (строка 5)

При первом проходе скрипт пытается удалить `ModelTypes` из `VBComponents`. Если удаление прошло успешно, то при втором проходе импортируется **пустой** класс — это корректно.

**НО если удаление НЕ сработало** (например, из-за блокировки COM, или компонент был найден по другому имени), то:
- Старый `ModelTypes.cls` с UDT остаётся в книге
- Новый `Mod_ModelTypes.bas` импортируется
- Возникает **конфликт**: UDT определены и в классе (`ModelTypes`), и в модуле (`Mod_ModelTypes`)
- Компилятор VBA видит дублирование или находит UDT в классе и выдаёт ошибку `0x800A03EC`

### Альтернативная причина: порядок компиляции

VBA компилирует классы раньше стандартных модулей. Если:
- `ModelTypes.cls` удалён, но
- `Mod_ModelTypes.bas` ещё не импортирован (или импортирован, но не скомпилирован)

то компилятор не находит типы `WorkEntry`/`WorkIdentity`/`PartIdentity` и выдаёт ошибку.

Однако скрипт импорта делает всё в рамках одной сессии Excel, поэтому эта причина маловероятна.

## План исправления

### Вариант A (рекомендуемый): Удалить `ModelTypes.cls` из проекта полностью

Класс пустой и не нужен. UDT живут в `Mod_ModelTypes.bas`.

1. Удалить файл [`src/classes/ModelTypes.cls`](src/classes/ModelTypes.cls)
2. В [`scripts/impVBA.py`](scripts/impVBA.py) убедиться, что скрипт корректно удаляет `ModelTypes` при импорте (строка 281: `vb_project.VBComponents.Remove(existing)`)
3. Запустить импорт заново

### Вариант B: Убедиться, что `ModelTypes.cls` удаляется гарантированно

Если класс нужно сохранить (пусть даже пустой), добавить в скрипт импорта принудительное удаление `ModelTypes` перед импортом:

```python
# В первый проход (строки 257-294) добавить:
for name_to_remove in ['ModelTypes']:
    try:
        existing = vb_project.VBComponents.Item(name_to_remove)
        vb_project.VBComponents.Remove(existing)
        print(f"  Removed (forced): {name_to_remove}")
    except Exception:
        pass  # уже удалён или не существует
```

### Вариант C: Проверить лог импорта

Запустить импорт с выводом лога и проверить строки:
- `Removed: ModelTypes (from ModelTypes.cls)` — должно быть
- `Successfully imported: ModelTypes.cls` — должно быть (пустой класс)

Если строка удаления отсутствует — значит компонент не удалился, и старый код с UDT остался в книге.

## Вывод

**Наиболее вероятная причина:** `ModelTypes.cls` не был удалён при импорте, и старый код с UDT внутри класса конфликтует с новым `Mod_ModelTypes.bas`. Либо `ModelTypes.cls` был удалён, но импортирован обратно с UDT (если файл на диске не соответствовал ожидаемому).

**Рекомендация:** выполнить Вариант A — удалить `ModelTypes.cls` из проекта и из файловой системы, так как он пустой и не несёт функциональной нагрузки.