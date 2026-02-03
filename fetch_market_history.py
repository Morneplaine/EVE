"""
Fetch market history from EVE Tycoon API and store in market_history_daily.

API: https://evetycoon.com/api/v1/market/history/{regionId}/{typeId}

- Data structure: API returns an array of daily records. Each record has:
  date (Unix ms), regionId, typeId, average, highest, lowest, orderCount, volume.
- History: The API provides full history (many days) in one response. No need to
  accumulate over time; we overwrite/upsert by (region_id, type_id, date_utc).
- Batching: There is no batch endpoint. One request per (regionId, typeId).
  Use --all-items to fetch for every item in the items table (long run; use --delay).
  Use --start N to resume after interrupt (progress log shows the number to use).

Jita is in The Forge; for EVE Tycoon API use region_id = 44992.
"""

import sqlite3
import time
import logging
from pathlib import Path
from datetime import datetime

import requests

DB_FILE = "eve_manufacturing.db"
EVETYCOON_BASE = "https://evetycoon.com/api/v1/market/history"
THE_FORGE_REGION_ID = 44992  # Jita / The Forge (EVE Tycoon region id)

# Same category exclusions as analyze_all_modules
EXCLUDED_CATEGORY_IDS = (
    25, 91, 1, 2, 3, 4, 5, 17, 29, 14, 9, 10, 11, 16, 20,
    2100, 2118, 24, 26, 30, 350001
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


def get_all_type_ids(conn):
    """Return list of all type_id from items table, ordered by typeID (for deterministic restart)."""
    cur = conn.execute("SELECT typeID FROM items ORDER BY typeID")
    return [row[0] for row in cur.fetchall()]


def get_reprocessable_type_ids(conn, limit=None):
    """Return list of type_ids from reprocessing_outputs (reprocessable modules) in The Forge, excluding categories."""
    placeholders = ",".join("?" * len(EXCLUDED_CATEGORY_IDS))
    query = f"""
        SELECT DISTINCT ro.itemTypeID
        FROM reprocessing_outputs ro
        JOIN items i ON ro.itemTypeID = i.typeID
        WHERE i.categoryID NOT IN ({placeholders})
        ORDER BY ro.itemTypeID
    """
    if limit:
        query += f" LIMIT {int(limit)}"
    cur = conn.execute(query, EXCLUDED_CATEGORY_IDS)
    return [row[0] for row in cur.fetchall()]


def transaction_skew(average, highest, lowest):
    """transaction_skew = 1 - ((average - lowest) / (highest - lowest)). Returns None if highest == lowest."""
    if highest is None or lowest is None or average is None:
        return None
    if highest <= lowest:
        return None
    return 1.0 - ((average - lowest) / (highest - lowest))


def fetch_history_for_type(region_id, type_id):
    """GET one type's history. Returns list of dicts with date_utc, average, highest, lowest, order_count, volume, transaction_skew."""
    url = f"{EVETYCOON_BASE}/{region_id}/{type_id}"
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        logger.warning("Fetch %s: %s", url, e)
        return None
    if not isinstance(data, list):
        return None
    rows = []
    for rec in data:
        try:
            ts_ms = rec.get("date")
            average = rec.get("average")
            highest = rec.get("highest")
            lowest = rec.get("lowest")
            order_count = rec.get("orderCount")
            volume = rec.get("volume")
            if ts_ms is None:
                continue
            dt = datetime.utcfromtimestamp(ts_ms / 1000.0)
            date_utc = dt.strftime("%Y-%m-%d")
            skew = transaction_skew(average, highest, lowest)
            rows.append({
                "date_utc": date_utc,
                "average": average,
                "highest": highest,
                "lowest": lowest,
                "order_count": order_count,
                "volume": volume,
                "transaction_skew": skew,
            })
        except Exception as e:
            logger.debug("Skip record %s: %s", rec, e)
    return rows


def ensure_table(conn):
    """Create market_history_daily if missing."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS market_history_daily (
            region_id INTEGER NOT NULL,
            type_id INTEGER NOT NULL,
            date_utc TEXT NOT NULL,
            average REAL NOT NULL,
            highest REAL NOT NULL,
            lowest REAL NOT NULL,
            order_count INTEGER,
            volume INTEGER,
            transaction_skew REAL,
            PRIMARY KEY (region_id, type_id, date_utc),
            FOREIGN KEY (type_id) REFERENCES items(typeID)
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_market_history_type_date ON market_history_daily(type_id, date_utc)"
    )
    conn.commit()


def run_fetch(
    region_id=THE_FORGE_REGION_ID,
    type_ids=None,
    all_items=False,
    start=0,
    limit=None,
    delay_seconds=1.0,
    progress_interval=50,
):
    """
    Fetch market history for given type_ids (or reprocessable modules if type_ids is None).
    all_items: if True, use all typeIDs from items table (can be interrupted; restart with --start N).
    start: skip first N items (0-based). Use to resume after interrupt; item numbers are logged.
    limit: max number of types to fetch (None = no limit). delay_seconds: pause between API calls.
    progress_interval: log progress every N items.
    """
    if not Path(DB_FILE).exists():
        logger.error("Database not found: %s", DB_FILE)
        return
    conn = sqlite3.connect(DB_FILE)
    ensure_table(conn)
    if type_ids is not None:
        type_ids = list(type_ids)
    elif all_items:
        type_ids = get_all_type_ids(conn)
        logger.info("Fetched %s type IDs (all items from items table)", len(type_ids))
    else:
        type_ids = get_reprocessable_type_ids(conn, limit=limit)
        logger.info("Fetched %s reprocessable type IDs (limit=%s)", len(type_ids), limit)
    total_in_list = len(type_ids)
    if start > 0:
        if start >= total_in_list:
            logger.warning("start=%s >= total items %s; nothing to do", start, total_in_list)
            conn.close()
            return
        type_ids = type_ids[start:]
        logger.info("Resuming from item index %s (skipping first %s). Processing items %s to %s (of %s total).",
                    start, start, start + 1, start + len(type_ids), total_in_list)
    if limit:
        type_ids = type_ids[: int(limit)]
    total_to_process = len(type_ids)
    if not type_ids:
        logger.warning("No type IDs to fetch")
        conn.close()
        return
    total_rows = 0
    for i, type_id in enumerate(type_ids):
        item_number = start + i + 1  # 1-based index in full list (for --start when resuming)
        rows = fetch_history_for_type(region_id, type_id)
        if rows:
            for row in rows:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO market_history_daily
                    (region_id, type_id, date_utc, average, highest, lowest, order_count, volume, transaction_skew)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        region_id,
                        type_id,
                        row["date_utc"],
                        row["average"],
                        row["highest"],
                        row["lowest"],
                        row.get("order_count"),
                        row.get("volume"),
                        row.get("transaction_skew"),
                    ),
                )
            total_rows += len(rows)
        if (i + 1) % progress_interval == 0:
            conn.commit()
            logger.info(
                "Progress: item %s/%s (type_id %s), %s daily rows stored. To resume later: --start %s",
                item_number,
                start + total_to_process,
                type_id,
                total_rows,
                start + i + 1,
            )
        time.sleep(delay_seconds)
    conn.commit()
    conn.close()
    logger.info("Done. Processed %s items, total daily rows stored: %s", total_to_process, total_rows)


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(
        description="Fetch EVE Tycoon market history for Jita (The Forge). Use --all-items and --start N to resume."
    )
    p.add_argument("--all-items", action="store_true", help="Fetch for all items in items table (long run; use --start to resume)")
    p.add_argument("--start", type=int, default=0, metavar="N", help="Skip first N items (0-based). Use to resume after interrupt (see progress log).")
    p.add_argument("--limit", type=int, default=None, help="Max number of types to fetch (default: no limit)")
    p.add_argument("--delay", type=float, default=1.0, help="Seconds between API requests (default: 1)")
    p.add_argument("--progress", type=int, default=50, metavar="N", help="Log progress every N items (default: 50)")
    p.add_argument("--types", type=str, default=None, help="Comma-separated typeIDs to fetch (overrides reprocessable/all-items)")
    args = p.parse_args()
    type_ids = None
    if args.types:
        type_ids = [int(x.strip()) for x in args.types.split(",") if x.strip()]
    run_fetch(
        type_ids=type_ids,
        all_items=args.all_items,
        start=args.start,
        limit=args.limit,
        delay_seconds=args.delay,
        progress_interval=args.progress,
    )
