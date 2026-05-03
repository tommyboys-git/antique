import sqlite3
import json
from datetime import datetime
from pathlib import Path
from config import DB_PATH


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS items (
            item_id         TEXT PRIMARY KEY,
            photo_paths     TEXT,           -- JSON list of absolute paths
            status          TEXT DEFAULT 'pending',
            -- pending | identifying | identified | comps_pending | comps_done | listing_done | exported | reviewed

            -- Phase 1: Identification
            item_type               TEXT,
            probable_maker          TEXT,
            makers_marks_observed   TEXT,
            probable_era            TEXT,
            material                TEXT,
            style_keywords          TEXT,   -- JSON list
            condition_notes         TEXT,
            confidence_score        INTEGER,
            confidence_reasoning    TEXT,
            suggested_search_query  TEXT,
            needs_review            INTEGER DEFAULT 0,

            -- Phase 2: Comps
            comp_low            REAL,
            comp_median         REAL,
            comp_high           REAL,
            comp_sample_size    INTEGER,
            comp_confidence     TEXT,
            comp_listings       TEXT,       -- JSON list of {title, price, date, condition, url}

            -- Phase 3: Listing
            ebay_title              TEXT,
            etsy_title              TEXT,
            fb_marketplace_title    TEXT,
            description             TEXT,
            suggested_price         REAL,
            suggested_platform      TEXT,
            tags                    TEXT,   -- JSON list
            category_ebay           TEXT,

            -- Meta
            created_at  TEXT DEFAULT (datetime('now')),
            updated_at  TEXT DEFAULT (datetime('now')),
            error_log   TEXT
        );

        CREATE TABLE IF NOT EXISTS run_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id     TEXT,
            phase       TEXT,
            outcome     TEXT,
            message     TEXT,
            ts          TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS lots (
            lot_id              TEXT PRIMARY KEY,
            lot_name            TEXT,
            item_ids            TEXT,   -- JSON list of item_id strings
            lot_price           REAL,
            lot_description     TEXT,
            suggested_platform  TEXT,
            thumbnail_path      TEXT,
            status              TEXT DEFAULT 'draft',
            created_at          TEXT DEFAULT (datetime('now')),
            updated_at          TEXT DEFAULT (datetime('now'))
        );
    """)

    # Migrate existing items table — add columns introduced in Phase 3.5
    for col_def in ["lot_id TEXT", "yard_sale INTEGER DEFAULT 0"]:
        try:
            conn.execute(f"ALTER TABLE items ADD COLUMN {col_def}")
        except Exception:
            pass  # column already exists

    conn.commit()
    conn.close()


def upsert_item(item_id: str, **fields):
    fields["updated_at"] = datetime.utcnow().isoformat()
    conn = get_conn()
    existing = conn.execute("SELECT item_id FROM items WHERE item_id = ?", (item_id,)).fetchone()
    if existing:
        set_clause = ", ".join(f"{k} = ?" for k in fields)
        values = list(fields.values()) + [item_id]
        conn.execute(f"UPDATE items SET {set_clause} WHERE item_id = ?", values)
    else:
        fields["item_id"] = item_id
        cols = ", ".join(fields.keys())
        placeholders = ", ".join("?" for _ in fields)
        conn.execute(f"INSERT INTO items ({cols}) VALUES ({placeholders})", list(fields.values()))
    conn.commit()
    conn.close()


def get_item(item_id: str) -> dict | None:
    conn = get_conn()
    row = conn.execute("SELECT * FROM items WHERE item_id = ?", (item_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_items_by_status(status: str) -> list[dict]:
    conn = get_conn()
    rows = conn.execute("SELECT * FROM items WHERE status = ?", (status,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_items() -> list[dict]:
    conn = get_conn()
    rows = conn.execute("SELECT * FROM items ORDER BY created_at").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def log_run(item_id: str, phase: str, outcome: str, message: str = ""):
    conn = get_conn()
    conn.execute(
        "INSERT INTO run_log (item_id, phase, outcome, message) VALUES (?, ?, ?, ?)",
        (item_id, phase, outcome, message),
    )
    conn.commit()
    conn.close()


def upsert_lot(lot_id: str, **fields):
    fields["updated_at"] = datetime.utcnow().isoformat()
    conn = get_conn()
    existing = conn.execute("SELECT lot_id FROM lots WHERE lot_id = ?", (lot_id,)).fetchone()
    if existing:
        set_clause = ", ".join(f"{k} = ?" for k in fields)
        values = list(fields.values()) + [lot_id]
        conn.execute(f"UPDATE lots SET {set_clause} WHERE lot_id = ?", values)
    else:
        fields["lot_id"] = lot_id
        cols = ", ".join(fields.keys())
        placeholders = ", ".join("?" for _ in fields)
        conn.execute(f"INSERT INTO lots ({cols}) VALUES ({placeholders})", list(fields.values()))
    conn.commit()
    conn.close()


def get_all_lots() -> list[dict]:
    conn = get_conn()
    rows = conn.execute("SELECT * FROM lots ORDER BY created_at").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_lot(lot_id: str) -> dict | None:
    conn = get_conn()
    row = conn.execute("SELECT * FROM lots WHERE lot_id = ?", (lot_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_lottable_items(price_ceiling: float = 20.0) -> list[dict]:
    """Items with suggested_price under ceiling, not yet assigned to a lot, listing complete."""
    conn = get_conn()
    rows = conn.execute(
        """SELECT * FROM items
           WHERE suggested_price < ?
             AND (lot_id IS NULL OR lot_id = '')
             AND (yard_sale IS NULL OR yard_sale = 0)
             AND status IN ('listing_done', 'reviewed')
           ORDER BY suggested_price DESC""",
        (price_ceiling,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def append_error(item_id: str, msg: str):
    conn = get_conn()
    row = conn.execute("SELECT error_log FROM items WHERE item_id = ?", (item_id,)).fetchone()
    existing = (row["error_log"] or "") if row else ""
    new_log = f"{existing}\n[{datetime.utcnow().isoformat()}] {msg}".strip()
    conn.execute("UPDATE items SET error_log = ? WHERE item_id = ?", (new_log, item_id))
    conn.commit()
    conn.close()
