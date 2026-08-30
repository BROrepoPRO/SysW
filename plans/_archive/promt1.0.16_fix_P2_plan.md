# План фикса дефекта П2 «все листы main во всех книгах work сломаны» (v1.0.16)

Роль: 🏗️ Architect (проектирование). Исполнение — роль Code после согласования с юзером.
Файлы фикса: только markdown + правки в `scripts/template_protection.py` / порядок сборки
(исполняет Code). Рабочие `.xlsm`/`.db` не пишутся на этом этапе.

---

## Контекст и корневая причина

1. **Блоат.** В [`scripts/template_protection.py:35`](../scripts/template_protection.py:35)
   константа `WORK_MAIN_DATA = "D4:AB2000"`. В COM-ветке [`apply_protection()`](../scripts/template_protection.py:126)
   на шагах [`:147-151`](../scripts/template_protection.py:147) (`ws.Range(rng).Locked = False`)
   и [`_unlock_used()`](../scripts/template_protection.py:70) (`UsedRange.Locked=False`)
   диапазон `D4:AB2000` материализуется в sheetXML → ~50 тыс. ячеек → защищённый лист
   выглядит как «одна большая пустая область». FreezePanes A4 — следствие той же логики.
2. **Корневой `work.xlsm` сжат до 5 листов** (было 19). Файл **исключён из git**
   ([`docs/git-workflow.md:178`](../docs/git-workflow.md:178)), `_backup/` пуст → эталон
   полной книги из VCS/бэкапа напрямую недоступен.
3. **Загрязнение `work0.xlsm`:** унаследован лог тестов в `Z1` и `vbaProject.bin`
   (конвейер run_tests → build_templates).
4. **openpyxl(keep_vba=True)+save опасен** для `.xlsm` (валидировано v1.0.9). Не применять.

Целевая структура `main` (из кода и docs): `A3:AB133`, заголовки стр.3, данные с 4-й,
`B4` № ЗН, `B5:B17` значения (покрывает `B13` цена н/ч, `B14` группа), `C1`/`Z1` служебные,
работы `L:N` (+ артикул `O`=15), запчасти `X:AA` (+ артикул `AB`=28).

---

## Этапы фикса

### Этап 0. Резервное копирование (обязателен, перед любыми действиями)

Исполняет Code. Скопировать в `_backup/` с меткой времени:
- корневой `work.xlsm`;
- `base/templates/work.xlsm`, `work0.xlsm`;
- `SysW.db`.
Плюс сделать дополнительную копию текущего `base/templates/work.xlsm` отдельно как
`_backup/work_template_bloated_<ts>.xlsm` — единственный носитель эталонной структуры `main`
с VBA (см. Этап 3).

---

### Этап 1. Де-блоат и очистка `main` (root и base/templates/work.xlsm)

Исполняет Code через Excel COM (не openpyxl).

Для листа `main` в корневом `work.xlsm` и в `base/templates/work.xlsm`:
1. `ws.Unprotect()` (снять существующую защиту/AllowEditRanges).
2. Очистить служебную ячейку лога тестов: `ws.Range("Z1").ClearContents()`.
3. Вычислить фактическую последнюю строку данных:
   `last = max(L.End(xlUp).Row, X.End(xlUp).Row)`; если `last < 4` → `last = 4`;
   предел-предохранитель `last <= 500` (не даём расти выше разумного).
4. Удалить лишние пустые строки и столбцы, стянув UsedRange к `A3:AB{last}`:
   - `ws.Range(last+1 & ":" & ws.Rows.Count).Delete()` (строки ниже данных);
   - `ws.Range("AC:" & ws.Columns.Count).Delete()` (столбцы правее AB).
5. Сбросить зоны: удалить существующие AllowEditRanges и снять `Locked` в фактическом
   диапазоне `A3:AB{last}` (малый диапазон, не материализует).
6. `wb.Save()`.

Результат: `main` снова `A3:AB{last}` (~133), без `D4:AB2000`, `Z1` пуст.

---

### Этап 2. Правка `scripts/template_protection.py` (корень фикса)

Исполняет Code. Цель — **не раздувать лист и блокировать только фактически используемые
ячейки**. Ключевые правки:

1. **Заменить жёсткую константу динамической логикой.**
   [`:35`](../scripts/template_protection.py:35):
   ```
   WORK_MAIN_DATA_MAX_ROW_FALLBACK = 133   # эталонный низ main (A3:AB133)
   ```
   Убрать использование `"D4:AB2000"` из COM-веток.

2. **Добавить хелпер фактической последней строки из XML** (без openpyxl):
   ```python
   def _sheet_max_data_row(xml_text, default=4):
       rows = re.findall(r'<row r="(\d+)"', xml_text)
       nums = [int(r) for r in rows]
       return max(nums) if nums else default
   ```
   Источник истины — фактический `sheetData`, а не фиксированная граница.

3. **COM-ветка [`apply_protection()`](../scripts/template_protection.py:126) — НЕ трогать
   данные большим диапазоном.**
   - В [`_resolve_zones()`](../scripts/template_protection.py:80) для `is_main` вернуть
     только `[B4, B5:B17, C1, Z1]` (убрать `WORK_MAIN_DATA`) → строки [`:90-92`](../scripts/template_protection.py:90).
   - В блоке `if is_main:` ([`:154-160`](../scripts/template_protection.py:154)) оставить
     `ws.Range("A:B").Locked = True` (блокировка столбцов на уровне колонок — не
     материализует ячейки) и разблокировку только `["B4","B5:B17"]` (малые диапазоны).
   - Не вызывать `_apply_allow_edit_ranges` для данных из COM: `edit_ranges` для main теперь
     без `D4:AB2000` → COM не добавит раздувающий AllowEditRange.

4. **XML-ветка [`apply_protection_xml()`](../scripts/template_protection.py:444) — доставлять
   динамическую зону данных.**
   Для листа `main` при построении зон подставить фактическую последнюю строку из XML:
   ```python
   last = _sheet_max_data_row(text, 4)
   data_zone = f"D4:AB{last}"     # покрывает работы L:N и запчасти X:AA
   zones = [z for z in zones if z != "D4:AB2000"] + [data_zone]
   ```
   Это создаёт лишь запись `allowEditRanges` (не ячейки) → лист не раздувается, а данные
   остаются редактируемыми под защитой.
   - Обновить [`build_zone_map()`](../scripts/template_protection.py:469), ветку `'work'`
     ([`:472-478`](../scripts/template_protection.py:472)): `z['main']` держать без жёсткого
     `D4:AB2000`; фактический диапазон добавляется в `apply_protection_xml` (см. выше).

5. **Убрать опасный openpyxl-save для `.xlsm`.**
   - [`ensure_freeze_panes_after_save()`](../scripts/template_protection.py:209) (load+save
     через openpyxl) **заменить** вызовами `apply_freeze_panes_xml()` (чистый zip, правка
     только `<pane>`).
   - В legacy [`apply_protection_templates.py`](../scripts/apply_protection_templates.py:58,67,76,85,94)
     заменить все `ensure_freeze_panes_after_save(...)` → `apply_freeze_panes_xml(...)`.
   - [`_extract_sheet_xmls()`](../scripts/template_protection.py:189) использует openpyxl
     ТОЛЬКО на чтение (keep_vba=True, без save) для имён листов — допустимо; при желании
     заменить на чтение `xl/workbook.xml` (как `_sheet_name_map`). Пометить как не-save.

6. **Добавить функцию безопасного удаления VBA-проекта на zip-уровне** для чистоты `work0`:
   ```python
   def strip_vba_project(path):
       """Удаляет xl/vbaProject.bin, его Relationship в workbook.xml.rels и Override
       в [Content_Types].xml; остальные записи переносятся без изменений."""
   ```
   (по аналогии с `apply_protection_xml`, но с фильтром по этим трём записям).
   Используется на Этапе 4 для гарантии отсутствия `vbaProject.bin` в `work0.xlsm`.

---

### Этап 3. Восстановление корневого `work.xlsm`

**Оценка источников эталона:**
- git: `work.xlsm` НЕ отслеживается ([`docs/git-workflow.md:178`](../docs/git-workflow.md:178))
  → восстановить из VCS нельзя (Code всё же проверит `git log --all -- work.xlsm` и `git reflog`
  на случай случайного добавления в прошлом).
- `_backup/`: пуст → недоступен.
- Внешние копии: OneDrive/сеть/Recycle Bin/Excel-autorecover — **запросить у юзера**.

**Безопасная последовательность (fallback, когда внешнего бэкапа нет):**
1. Взять за основу **де-блоатированный** `base/templates/work.xlsm` (после Этапа 1) — он
   содержит эталонную структуру `main A3:AB133` и импортированные VBA-модули.
2. `scripts/impVBA.py` поверх него — гарантировать, что весь код соответствует `src/`
   (чистые модули, без остатков тестов).
3. `scripts/check_vba_syntax.py` — ранний контроль компиляции.
4. Привести ядро листов к рабочему набору: `main, spisok, models, libname`
   (как в [`build_templates.build_work_templates`](../scripts/build_templates.py:245)).
5. Снять защиту, очистить `Z1`, сохранить как корневой `work.xlsm`.

**Риск этапа:** исходные 19 листов (полная книга) невосстановимы без внешней копии;
производные листы (группы моделей, temp*, макеты) генерируются макросами при работе
(`ImportVH`, `CreateModelGroupFile`) — поэтому ядро из 4 листов с корректным `main` является
рабочей основой. Полное соответствие исходной 19-листовой раскладке подтверждается только
при наличии пользовательского бэкапа. Этот риск фиксируется в критериях приёмки как
допущение.

---

### Этап 4. Пересборка шаблонов в безопасном порядке (анти-загрязнение)

Требование: шаблоны собираются из **чистой корневой структуры ПОСЛЕ прогона тестов и
очистки Z1**. Контролируемая последовательность (для фикса, не полный `build_all`):

1. Чистый корневой `work.xlsm` из Этапа 3 (Z1 пуст, без защиты).
2. `scripts/impVBA.py` (код из `src/`).
3. `scripts/check_vba_syntax.py`.
4. `scripts/run_tests.py` — прогон на корневом `work.xlsm` (валидирует `ImportDataToMain`,
   должен выполняться, не SKIP). По завершении файл НЕ сохраняется
   (`Close(SaveChanges=False)` в [`run_tests.py:243`](../scripts/run_tests.py:243) — это уже так).
5. **Пост-тестовая очистка:** открыть корневой `work.xlsm`, `Z1.ClearContents()`, снять
   защиту, удалить служебные/временные листы, если появились, `Save()`.
6. `scripts/build_templates.py` — собрать шаблоны из чистого корневого `work.xlsm`.
   - Для `work0.xlsm` после `_remove_vba` вызвать `strip_vba_project()` (Этап 2.6) — гарантия
     отсутствия `vbaProject.bin`.
7. Применить защиту через исправленную логику (динамическая зона) — см. Этап 5.

**Рекомендация по `build_all.py`** ([`main()`](../scripts/build_all.py:398)): постоянный порядок
— добавить этап «пост-тестовой очистки корневого work.xlsm» и перенести построение шаблонов
`build_templates` ПОСЛЕ `run_tests` + очистки Z1. Это устраняет загрязнение конвейерно,
а не разово. (Выполняется Code как часть фикса либо отдельным follow-up.)

---

### Этап 5. Применение защиты/FreezePanes (безопасная доставка)

Исполняет Code. Все доставки для `.xlsm` — **только точечной XML-инъекцией (Вариант D)**,
без openpyxl-rewrite:
- `apply_protection_xml(path, build_zone_map('work'))` — sheetProtection + allowEditRanges
  с динамической зоной `D4:AB{last}` для main (после Этапа 2.4);
- `apply_freeze_panes_xml(path)` — FreezePanes A4 (правка только `<pane>` в zip);
- для `work0.xlsm` — `apply_freeze_panes_xml` + `strip_vba_project`;
- для моделей — существующий `apply_freeze_panes_to_models` (уже Вариант C, безопасен).

Логика в [`build_templates.apply_all_protection`](../scripts/build_templates.py:332) уже
использует Вариант D; после Этапа 2 она дополнительно получает динамическую зону.

---

### Этап 6. Очистка временных и диагностических файлов

Исполняет Code. Удалить (если присутствуют):
- `diag_main.py`, `diag_zip.py` в корне (на момент анализа уже отсутствуют — только вкладки;
  проверить и при наличии удалить);
- любые `diag_*.py` в корне;
- `_temp_export/`, `_temp_import/`, `__pycache__/`, `*.pyc`, `*.pyo`;
- остатки zip-правок `*.prot.tmp`, `*.fr.tmp`, `*.tmp`;
- `logs/*.log` при необходимости (кроме `.gitkeep`).
Добавить эти пути в этап очистки и в `.codeassistantignore`, чтобы не попадали в коммиты.

---

### Этап 7. Проверка (Debug) — критерии приёмки фикса П2

Запускает роль Debug/Code после фикса.

**Автоматические проверки:**
1. `python scripts/check_vba_syntax.py` — синтаксис `src/` без ошибок (exit 0).
2. Статическая проверка XML шаблонов через `verify_sheet_protection()`:
   - лист `main` `base/templates/work.xlsm`: protect=да, allowEdit=да (есть `B4`, `B5:B17`,
     `C1`, `Z1`, зона `D4:AB{last~133}` БЕЗ `AB2000`), pane=A4;
   - в `main` нет `D4:AB2000` и max `r` в sheetData ≈ 133 (не 2000).
3. `work0.xlsm`: отсутствует `xl/vbaProject.bin` (проверка через zip), pane=A4, `Z1` пуст.
4. Корневой `work.xlsm` открывается COM-ом; лист `main` имеет структуру `A3:AB133`
   (заголовки стр.3, данные с 4-й); UsedRange не простирается до `AB2000`.
5. Ячейки `B4`, `B13`, `B14`, зоны `L:N` и `X:AA` доступны для ввода при включённой защите
   (не заблокированы) — проверка AllowEditRanges + отсутствие блокировки.
6. Прогон `ImportDataToMain` в тестах выполняется (результат PASS, не SKIP).

**Критерии приёмки (итоговый чек-лист):**
- [ ] `main` открывается корректно, не выглядит «одной большой пустой ячейкой».
- [ ] Структура `main` = `A3:AB133` (заголовки стр.3, данные с 4-й), 28 столбцов.
- [ ] Лист не раздут до `AB2000` (UsedRange ≤ фактического).
- [ ] `B4/B13/B14` и зоны `L:N`/`X:AA` доступны в рабочей книге (не защищены) — для теста П1.
- [ ] Нет лога тестов в `Z1` у root и шаблонов.
- [ ] `work0.xlsm` без `vbaProject.bin`.
- [ ] Тест `ImportDataToMain` выполняется, а не SKIP.
- [ ] Фикс не вносит openpyxl-save в `.xlsm` (только zip/COM/XML-инъекция).

---

## Риски и митигации

| Риск | Митигация |
| --- | --- |
| Полная 19-листовая книга невосстановима (нет в git/бэкапе) | Запросить внешний бэкап; fallback — ядро из 4 листов + генерация производных макросами; зафиксировать как допущение. |
| Изменение порядка сборки `build_all` может затронуть другие прогоны | Сначала контролируемая последовательность для фикса; реордер — отдельным осторожным изменением с проверкой. |
| Динамическая зона из XML даст неверный `last`, если данные шире `AB`/ниже 133 | Хелпер берёт max `r` из `sheetData`; предохранитель `last<=500`; де-блоат стягивает UsedRange заранее. |
| COM-шаг снова материализует ячейки | Убраны большие `Locked=False`/`AllowEditRanges` из COM; большие операции только на XML-уровне. |
| `work0` снова получит `vbaProject.bin` | `strip_vba_project()` на zip-уровне + проверка в Этапе 7. |
| openpyxl-save в legacy-скрипте повредит `.xlsm` | Замена `ensure_freeze_panes_after_save` на `apply_freeze_panes_xml` во всех вызовах. |

---

## Итог

Фикс П2 = де-блоат `main` + замена жёсткой зоны `D4:AB2000` динамической фактической
последней строкой на XML-уровне + безопасная доставка защиты/FreezePanes без openpyxl-save +
анти-загрязняющий порядок пересборки (после тестов и очистки Z1) + strip `vbaProject.bin` из
`work0` + очистка диагностических файлов. Передаётся роли Code для исполнения после согласования.