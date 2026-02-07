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
import sys
import time
import logging
from pathlib import Path
from datetime import datetime, timezone

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


def get_type_ids_from_prices(conn):
    """Return list of type_id from prices table (same set as Update All Prices), ordered by typeID."""
    cur = conn.execute("SELECT typeID FROM prices ORDER BY typeID")
    return [row[0] for row in cur.fetchall()]


# Mineral/item names for --scope blueprint_consensus_mineral (same as update_mineral_prices)
_MINERAL_AND_MATERIAL_NAMES = [
    "Tritanium", "Pyerite", "Mexallon", "Isogen", "Nocxium", "Zydrine", "Megacyte", "Morphite",
    "Armor Mutaplasmid Residue", "Astronautic Mutaplasmid Residue", "Crystalline Isogen-10",
    "Damage Control Mutaplasmid Residue", "Drone Mutaplasmid Residue", "Engineering Mutaplasmid Residue",
    "Large Mutaplasmid Residue", "Medium Mutaplasmid Residue", "Mutaplasmid Residue",
    "Shield Mutaplasmid Residue", "Small Mutaplasmid Residue", "Stasis Webifier Mutaplasmid Residue",
    "Warp Disruption Mutaplasmid Residue", "Weapon Upgrade Mutaplasmid Residue", "X-Large Mutaplasmid Residue",
    "Zero-Point Condensate",
]


def get_type_ids_blueprint_consensus_mineral(conn):
    """Return type_ids from blueprint + group_consensus (input_quantity_cache) plus minerals/materials by name."""
    # Blueprint and group_consensus from input_quantity_cache
    cur = conn.execute("""
        SELECT DISTINCT c.typeID
        FROM input_quantity_cache c
        WHERE c.source IN ('blueprint', 'group_consensus')
        ORDER BY c.typeID
    """)
    from_cache = [row[0] for row in cur.fetchall()]
    # Minerals/materials by name (same as update_mineral_prices)
    placeholders = ",".join("?" * len(_MINERAL_AND_MATERIAL_NAMES))
    cur = conn.execute(
        f"SELECT typeID FROM items WHERE typeName IN ({placeholders}) ORDER BY typeID",
        _MINERAL_AND_MATERIAL_NAMES,
    )
    from_minerals = [row[0] for row in cur.fetchall()]
    combined = sorted(set(from_cache) | set(from_minerals))
    return combined


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


def get_type_name(conn, type_id):
    """Return typeName from items for type_id, or None if not found."""
    cur = conn.execute("SELECT typeName FROM items WHERE typeID = ?", (type_id,))
    row = cur.fetchone()
    return row[0] if row else None


def fetch_history_for_type(region_id, type_id):
    """GET one type's history. Returns list of dicts with date_utc, average, highest, lowest, order_count, volume."""
    url = f"{EVETYCOON_BASE}/{region_id}/{type_id}"
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        logger.warning("Fetch %s: %s", url, e)
        return None
    return _parse_history_response(data, region_id, type_id) if isinstance(data, list) else None


def _create_table(conn):
    """Create market_history_daily table and index (shared by ensure_table and reset_table)."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS market_history_daily (
            region_id INTEGER NOT NULL,
            type_id INTEGER NOT NULL,
            type_name TEXT,
            date_utc TEXT NOT NULL,
            average REAL NOT NULL,
            highest REAL NOT NULL,
            lowest REAL NOT NULL,
            order_count INTEGER,
            volume INTEGER,
            PRIMARY KEY (region_id, type_id, date_utc),
            FOREIGN KEY (type_id) REFERENCES items(typeID)
        )
    """)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_market_history_type_date ON market_history_daily(type_id, date_utc)"
    )


def reset_table(conn):
    """Drop market_history_daily and recreate it blank (current schema with type_name)."""
    conn.execute("DROP TABLE IF EXISTS market_history_daily")
    _create_table(conn)
    conn.commit()
    logger.info("Dropped and recreated market_history_daily (empty).")


def ensure_table(conn):
    """Create market_history_daily if missing; migrate existing table (add type_name, drop transaction_skew)."""
    _create_table(conn)
    # Migration: add type_name and drop transaction_skew on existing tables
    cur = conn.execute("PRAGMA table_info(market_history_daily)")
    columns = [row[1] for row in cur.fetchall()]
    if "type_name" not in columns:
        try:
            conn.execute("ALTER TABLE market_history_daily ADD COLUMN type_name TEXT")
        except sqlite3.OperationalError as e:
            if "duplicate column" not in str(e).lower():
                raise
    if "transaction_skew" in columns:
        try:
            conn.execute("ALTER TABLE market_history_daily DROP COLUMN transaction_skew")
        except sqlite3.OperationalError:
            pass  # SQLite < 3.35 has no DROP COLUMN; column is unused
    conn.commit()


def run_fetch(
    region_id=THE_FORGE_REGION_ID,
    type_ids=None,
    all_items=False,
    scope="prices",
    start=0,
    limit=None,
    delay_seconds=1.0,
    progress_interval=50,
):
    """
    Fetch market history for given type_ids (or reprocessable modules if type_ids is None).
    all_items: if True, use type IDs from prices table (scope='prices') or blueprint+consensus+mineral (scope='blueprint_consensus_mineral').
    scope: when all_items, 'prices' = same set as Update All Prices; 'blueprint_consensus_mineral' = blueprint + group_consensus + mineral only.
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
        if scope == "blueprint_consensus_mineral":
            type_ids = get_type_ids_blueprint_consensus_mineral(conn)
            logger.info("Fetched %s type IDs (blueprint, group consensus, and mineral only)", len(type_ids))
        else:
            type_ids = get_type_ids_from_prices(conn)
            logger.info("Fetched %s type IDs (same set as Update All Prices)", len(type_ids))
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
            type_name = get_type_name(conn, type_id)
            for row in rows:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO market_history_daily
                    (region_id, type_id, type_name, date_utc, average, highest, lowest, order_count, volume)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        region_id,
                        type_id,
                        type_name,
                        row["date_utc"],
                        row["average"],
                        row["highest"],
                        row["lowest"],
                        row.get("order_count"),
                        row.get("volume"),
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


def _parse_history_response(data, region_id, type_id):
    """Parse API response (list of daily records) into rows. Same logic as fetch_history_for_type."""
    if not isinstance(data, list):
        return []
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
            # Accept Unix ms (number) or seconds (number)
            try:
                ts = float(ts_ms)
                if ts > 1e12:
                    ts = ts / 1000.0  # was ms
                dt = datetime.fromtimestamp(ts, tz=timezone.utc)
            except (TypeError, ValueError, OSError):
                continue
            date_utc = dt.strftime("%Y-%m-%d")
            rows.append({
                "date_utc": date_utc,
                "average": average,
                "highest": highest,
                "lowest": lowest,
                "order_count": order_count,
                "volume": volume,
            })
        except Exception as e:
            logger.debug("Skip record %s: %s", rec, e)
    return rows


def run_test(region_id=THE_FORGE_REGION_ID):
    """
    Test the API and optionally DB without full run.
    Fetches type 34 (Tritanium), prints result, and if DB exists writes one type.
    Tries region 10000002 (The Forge) if region_id returns no rows.
    """
    type_id = 34
    url = f"{EVETYCOON_BASE}/{region_id}/{type_id}"
    print(f"Testing EVE Tycoon API: {url}")
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"API error: {e}")
        return
    if not isinstance(data, list):
        print(f"Unexpected response (not a list): {type(data)}. Keys: {list(data.keys()) if isinstance(data, dict) else 'n/a'}")
        return
    if len(data) == 0:
        print(f"API returned empty list for region {region_id}. Trying region 10000002 (The Forge)...")
        region_id = 10000002
        url = f"{EVETYCOON_BASE}/{region_id}/{type_id}"
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"API error (fallback): {e}")
            return
        if not isinstance(data, list) or len(data) == 0:
            print("Fallback region also returned no data.")
            return
        print(f"Got {len(data)} raw records from region 10000002.")
    else:
        print(f"Got {len(data)} raw records from API.")
    if len(data) > 0:
        first = data[0]
        print(f"  First record keys: {list(first.keys()) if isinstance(first, dict) else type(first)}")
    rows = _parse_history_response(data, region_id, type_id)
    if not rows:
        print("No rows parsed from response (check date/average keys in records).")
        return
    print(f"OK: got {len(rows)} daily records for type_id={type_id}")
    sample = rows[-1]
    skew = transaction_skew(sample["average"], sample["highest"], sample["lowest"])
    print(f"  Sample (newest): date_utc={sample['date_utc']!r}, average={sample['average']}, volume={sample.get('volume')}, transaction_skew={skew}")
    if Path(DB_FILE).exists():
        conn = sqlite3.connect(DB_FILE)
        ensure_table(conn)
        type_name = get_type_name(conn, type_id)
        for row in rows:
            conn.execute(
                """
                INSERT OR REPLACE INTO market_history_daily
                (region_id, type_id, type_name, date_utc, average, highest, lowest, order_count, volume)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    region_id,
                    type_id,
                    type_name,
                    row["date_utc"],
                    row["average"],
                    row["highest"],
                    row["lowest"],
                    row.get("order_count"),
                    row.get("volume"),
                ),
            )
        conn.commit()
        conn.close()
        print(f"  Wrote {len(rows)} rows to {DB_FILE} -> market_history_daily (type_id={type_id}, type_name={type_name!r})")
    else:
        print(f"  Database {DB_FILE} not found; skipping write. Create DB first (e.g. run build_database.py).")


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(
        description="Fetch EVE Tycoon market history for Jita (The Forge). Use --all-items and --start N to resume."
    )
    p.add_argument("--test", action="store_true", help="Test API and DB: fetch type 34 only, print result, write to DB if exists")
    p.add_argument("--all-items", action="store_true", help="Fetch for types in prices table (same as Update All Prices); use --scope to narrow.")
    p.add_argument("--scope", choices=["prices", "blueprint_consensus_mineral"], default="prices", help="With --all-items: 'prices' = Update All Prices set (default); 'blueprint_consensus_mineral' = blueprint + group_consensus + mineral only.")
    p.add_argument("--start", type=int, default=0, metavar="N", help="Skip first N items (0-based). Use to resume after interrupt (see progress log).")
    p.add_argument("--limit", type=int, default=None, help="Max number of types to fetch (default: no limit)")
    p.add_argument("--delay", type=float, default=1.0, help="Seconds between API requests (default: 1)")
    p.add_argument("--progress", type=int, default=50, metavar="N", help="Log progress every N items (default: 50)")
    p.add_argument("--types", type=str, default=None, help="Comma-separated typeIDs to fetch (overrides reprocessable/all-items)")
    p.add_argument("--region", type=int, default=None, metavar="ID", help="Region ID (default: 44992). Use 10000002 for The Forge in EVE SDE.")
    p.add_argument("--reset", action="store_true", help="Drop market_history_daily and recreate it empty (then exit).")
    args = p.parse_args()
    if args.reset:
        if not Path(DB_FILE).exists():
            logger.error("Database not found: %s", DB_FILE)
            sys.exit(1)
        conn = sqlite3.connect(DB_FILE)
        reset_table(conn)
        conn.close()
        sys.exit(0)
    if args.test:
        run_test()
        sys.exit(0)
    type_ids = None
    if args.types:
        type_ids = [int(x.strip()) for x in args.types.split(",") if x.strip()]
    region_id = args.region if args.region is not None else THE_FORGE_REGION_ID
    run_fetch(
        region_id=region_id,
        type_ids=type_ids,
        all_items=args.all_items,
        scope=args.scope,
        start=args.start,
        limit=args.limit,
        delay_seconds=args.delay,
        progress_interval=args.progress,
    )
