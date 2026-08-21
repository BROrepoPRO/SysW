-- ============================================================
-- SysW.db  Схема v1 (PRAGMA user_version = 1)
-- ============================================================
-- Единая база данных модельных данных системы SysW.
-- Источник данных: base/models/*.xlsm (конвертер scripts/migrate_models_to_sqlite.py).
-- Чтение: VBA через Mod_SQLiteDB (ADO/ODBC); Python через sqlite3 stdlib.
--
-- Таблицы:
--   aggregates     — справочник агрегатов (16 кодов AGG_*)
--   model_groups   — группы моделей (UAZ, GAZ, 4x4, 2170, 2180, 2190, ...)
--   parts          — каталог запчастей группы (лист z4)
--   works          — каталог работ группы (лист {GroupName})
--   model_works    — модельные работы + тождества работ (лист {GroupName}w)
--   model_parts    — модельные запчасти + тождества запчастей (лист {GroupName}z4)
--   matlib_entries — нормализованная библиотека соответствий (тождеств)
-- ============================================================

-- Справочник агрегатов (16 кодов AGG_*)
CREATE TABLE IF NOT EXISTS aggregates (
    code TEXT PRIMARY KEY,               -- DIAG, TO, ENG, ...
    name_ru TEXT NOT NULL,               -- "Диагностика", ...
    sort_order INTEGER DEFAULT 0
);

-- Группы моделей (соответствует base/models/{GroupName}.xlsm)
CREATE TABLE IF NOT EXISTS model_groups (
    group_name TEXT PRIMARY KEY,         -- UAZ, GAZ, 4x4, 2170, ...
    created_at TEXT DEFAULT (datetime('now')),
    note TEXT
);

-- Каталог запчастей (лист z4 файла группы; код уникален в пределах группы)
CREATE TABLE IF NOT EXISTS parts (
    group_name TEXT NOT NULL,
    code TEXT NOT NULL,                  -- A: Code
    name TEXT NOT NULL,                  -- B: Name
    unit TEXT,                           -- C: Unit
    price REAL,                          -- D: Price
    note TEXT,                           -- E: Note
    PRIMARY KEY (group_name, code),
    FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
);

-- Каталог работ (лист {GroupName}; код уникален в пределах группы)
CREATE TABLE IF NOT EXISTS works (
    group_name TEXT NOT NULL,
    code TEXT NOT NULL,                  -- A: Code
    name TEXT NOT NULL,                  -- B: Name
    unit TEXT,                           -- C: Unit
    norm_hours REAL,                     -- D: NormHours
    price REAL,                          -- E: Price
    note TEXT,                           -- F: Note
    PRIMARY KEY (group_name, code),
    FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
);

-- Модельные запчасти (лист {GroupName}z4, данные с 4-й строки)
-- Содержит и тождества (OutArticle/InCatNum/InName/Aggregate), т.к. соответствия
-- хранятся ВНУТРИ листа {GroupName}z4.
CREATE TABLE IF NOT EXISTS model_parts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_name TEXT NOT NULL,
    out_article TEXT NOT NULL,           -- B: OutArticle (артикул модельный)
    out_name TEXT,                       -- C: OutName
    qty_zn REAL DEFAULT 0,               -- G: QtyZN (кол-во ЗН)
    price REAL DEFAULT 0,                -- F: Price (цена за ед.)
    aggregate TEXT,                      -- I: Агрегат (код)
    in_catnum TEXT,                      -- J: № кат. (входящий)
    in_name TEXT,                        -- K: Наим-ние (входящее)
    note TEXT,
    FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
);

-- Модельные работы (лист {GroupName}w, данные с 4-й строки)
-- Содержит тождества работ (OutArticle/InName/Aggregate).
CREATE TABLE IF NOT EXISTS model_works (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_name TEXT NOT NULL,
    out_article TEXT NOT NULL,           -- B: OutArticle (артикул модельный)
    out_name TEXT,                       -- C: OutName
    norm_hours REAL DEFAULT 0,           -- D: н/ч
    qty_zn REAL DEFAULT 0,               -- G: QtyZN (кол-во ЗН)
    aggregate TEXT,                      -- I: Агрегат (код)
    in_name TEXT,                        -- J: Наим-ние (входящее)
    note TEXT,
    FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
);

-- Глобальная таблица соответствий (библиотека соответствий).
-- В легаси соответствия лежат ВНУТРИ {GroupName}z4 / {GroupName}w,
-- поэтому здесь таблица НОРМАЛИЗУЕТ их в единый реестр (для GetMatLibEntries).
CREATE TABLE IF NOT EXISTS matlib_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    group_name TEXT NOT NULL,
    entry_type TEXT NOT NULL,            -- 'work' | 'part' (входящая позиция)
    entry_code TEXT,                     -- код входящей позиции (поисковый ключ)
    entry_name TEXT,
    target_type TEXT NOT NULL,           -- 'work' | 'mod_work' | 'part' | 'mod_part'
    target_code TEXT,
    target_name TEXT,
    coefficient REAL DEFAULT 1.0,
    target_sheet TEXT,
    note TEXT,
    FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
);

-- Индексы для производительности
CREATE INDEX IF NOT EXISTS idx_parts_group ON parts(group_name);
CREATE INDEX IF NOT EXISTS idx_works_group ON works(group_name);
CREATE INDEX IF NOT EXISTS idx_model_parts_group ON model_parts(group_name);
CREATE INDEX IF NOT EXISTS idx_model_parts_in_catnum ON model_parts(group_name, in_catnum);
CREATE INDEX IF NOT EXISTS idx_model_works_group ON model_works(group_name);
CREATE INDEX IF NOT EXISTS idx_model_works_in_name ON model_works(group_name, in_name);
CREATE INDEX IF NOT EXISTS idx_matlib_entry ON matlib_entries(group_name, entry_type, entry_code);

-- Версионирование БД
PRAGMA user_version = 1;