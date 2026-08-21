#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль схемы SysW.db (SQLite).

Единый источник DDL-схемы для программного создания БД.
Схема дублирует db/schema.sql и является Python-структурой,
используемой конвертером migrate_models_to_sqlite.py при init_db().

Совместимость:
- Версия схемы: 1 (PRAGMA user_version = 1).
- Таблицы: aggregates, model_groups, parts, works, model_works,
  model_parts, matlib_entries.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

# Версия схемы БД (должна совпадать с PRAGMA user_version в DDL)
DB_SCHEMA_VERSION = 1

# Путь к файлу DDL-схемы (относительно корня проекта)
SCHEMA_SQL_PATH = Path(__file__).resolve().parent.parent / "db" / "schema.sql"

# ---------------------------------------------------------------------------
# 16 агрегатов (коды AGG_* из src/modules/Mod_Constants.bas).
# Порядок соответствует исходному порядку констант (sort_order 1..16).
# ---------------------------------------------------------------------------
AGGREGATES = [
    ("DIAG", "Диагностика"),
    ("TO", "ТО"),
    ("ENG", "Двигатель"),
    ("TRANS", "Трансмиссия"),
    ("CLUTCH", "Сцепление"),
    ("SUSP", "Подвеска"),
    ("CHASS", "Ходовая"),
    ("BRAKE", "Тормозная система"),
    ("FUEL", "Топливная система"),
    ("COOL", "Система охлаждения"),
    ("HVAC", "Система обогрева и кондиционирования"),
    ("STEER", "Система рулевого управления"),
    ("ELEC", "Электрооборудование"),
    ("EXH", "Система выхлопных газов"),
    ("BODY", "Кузов"),
    ("OTHERS", "Прочие работы"),
]

# ---------------------------------------------------------------------------
# DDL-схема как список SQL-операторов (без PRAGMA user_version —
# он устанавливается отдельно функцией set_user_version).
# ---------------------------------------------------------------------------
SCHEMA_STATEMENTS = [
    """
    CREATE TABLE IF NOT EXISTS aggregates (
        code TEXT PRIMARY KEY,               -- DIAG, TO, ENG, ...
        name_ru TEXT NOT NULL,               -- "Диагностика", ...
        sort_order INTEGER DEFAULT 0
    );
    """,
    """
    CREATE TABLE IF NOT EXISTS model_groups (
        group_name TEXT PRIMARY KEY,         -- UAZ, GAZ, 4x4, 2170, ...
        created_at TEXT DEFAULT (datetime('now')),
        note TEXT
    );
    """,
    """
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
    """,
    """
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
    """,
    """
    CREATE TABLE IF NOT EXISTS model_parts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT NOT NULL,
        out_article TEXT NOT NULL,           -- B: OutArticle
        out_name TEXT,                       -- C: OutName
        qty_zn REAL DEFAULT 0,               -- G: QtyZN
        price REAL DEFAULT 0,                -- F: Price
        aggregate TEXT,                      -- I: Агрегат (код)
        in_catnum TEXT,                      -- J: № кат. (входящий)
        in_name TEXT,                        -- K: Наим-ние (входящее)
        note TEXT,
        FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
    );
    """,
    """
    CREATE TABLE IF NOT EXISTS model_works (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT NOT NULL,
        out_article TEXT NOT NULL,           -- B: OutArticle
        out_name TEXT,                       -- C: OutName
        norm_hours REAL DEFAULT 0,           -- D: н/ч
        qty_zn REAL DEFAULT 0,               -- G: QtyZN
        aggregate TEXT,                      -- I: Агрегат (код)
        in_name TEXT,                        -- J: Наим-ние (входящее)
        note TEXT,
        FOREIGN KEY (group_name) REFERENCES model_groups(group_name)
    );
    """,
    """
    CREATE TABLE IF NOT EXISTS matlib_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT NOT NULL,
        entry_type TEXT NOT NULL,            -- 'work' | 'part'
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
    """,
    "CREATE INDEX IF NOT EXISTS idx_parts_group ON parts(group_name);",
    "CREATE INDEX IF NOT EXISTS idx_works_group ON works(group_name);",
    "CREATE INDEX IF NOT EXISTS idx_model_parts_group ON model_parts(group_name);",
    "CREATE INDEX IF NOT EXISTS idx_model_parts_in_catnum ON model_parts(group_name, in_catnum);",
    "CREATE INDEX IF NOT EXISTS idx_model_works_group ON model_works(group_name);",
    "CREATE INDEX IF NOT EXISTS idx_model_works_in_name ON model_works(group_name, in_name);",
    "CREATE INDEX IF NOT EXISTS idx_matlib_entry ON matlib_entries(group_name, entry_type, entry_code);",
]


def schema_sql_text() -> str:
    """Возвращает полный текст DDL-схемы (с PRAGMA user_version) для документации/логов."""
    parts = [
        "-- ============================================================\n"
        "-- SysW.db  Схема v1 (PRAGMA user_version = 1)\n"
        "-- ============================================================\n"
    ]
    parts.extend(stmt.strip() for stmt in SCHEMA_STATEMENTS)
    parts.append("PRAGMA user_version = {};".format(DB_SCHEMA_VERSION))
    return "\n\n".join(parts)


def set_user_version(conn: sqlite3.Connection, version: int = DB_SCHEMA_VERSION) -> None:
    """Устанавливает PRAGMA user_version на БД."""
    conn.execute("PRAGMA user_version = {}".format(int(version)))


def get_user_version(conn: sqlite3.Connection) -> int:
    """Возвращает текущее значение PRAGMA user_version."""
    row = conn.execute("PRAGMA user_version").fetchone()
    return int(row[0]) if row else 0


def init_db(db_path, *, drop_tables: bool = False, conn: sqlite3.Connection | None = None) -> sqlite3.Connection:
    """Создаёт/обновляет схему SysW.db и возвращает открытое соединение.

    Параметры:
        db_path: путь к файлу БД (Path или str).
        drop_tables: если True — удаляет таблицы перед созданием (полный сброс).
        conn: существующее соединение (если передано, используется оно;
              иначе создаётся новое для db_path).

    Идемпотентность: CREATE TABLE IF NOT EXISTS / CREATE INDEX IF NOT EXISTS.
    Функция применяет DDL, заполняет агрегаты и выставляет user_version.
    """
    own_conn = conn is None
    if conn is None:
        conn = sqlite3.connect(str(db_path))

    conn.execute("PRAGMA foreign_keys = ON")

    if drop_tables:
        # Порядок важен из-за внешних ключей (дочерние удаляются первыми).
        tables = [
            "matlib_entries",
            "model_works",
            "model_parts",
            "works",
            "parts",
            "model_groups",
            "aggregates",
        ]
        for t in tables:
            conn.execute('DROP TABLE IF EXISTS "{}";'.format(t))

    for stmt in SCHEMA_STATEMENTS:
        conn.execute(stmt)

    # Заполнение справочника агрегатов (идемпотентно).
    for idx, (code, name) in enumerate(AGGREGATES, start=1):
        conn.execute(
            "INSERT OR REPLACE INTO aggregates(code, name_ru, sort_order) VALUES (?, ?, ?)",
            (code, name, idx),
        )

    set_user_version(conn)
    conn.commit()

    if own_conn:
        conn.close()
        conn = sqlite3.connect(str(db_path))

    return conn


if __name__ == "__main__":
    # Демонстрационный запуск: создать БД в корне проекта.
    from config import PROJECT_DIR  # type: ignore

    _db = PROJECT_DIR / "SysW.db"
    _c = init_db(_db, drop_tables=True)
    _ver = get_user_version(_c)
    print("SysW.db создан:", _db)
    print("user_version:", _ver)
    _c.close()