"""
Analyze daily market history patterns by day-of-week for selected items.

Focus:
- Minerals: Tritanium (34), Pyerite (35), Mexallon (36), Isogen (37),
  Nocxium (38), Megacyte (40), Morphite (11399)
- Modules currently tracked in on_offer_items.

Output: for each item, average price/volume per weekday (Mon–Sun).
Run from the repo root:

    python analyze_market_patterns.py
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from collections import defaultdict
from datetime import datetime

from fetch_market_history import (
    DB_FILE as MARKET_DB_FILE,
    THE_FORGE_REGION_ID,
    expected_buy_order_volume_for_day,
)


DATABASE_FILE = MARKET_DB_FILE

# Minerals of interest
MINERAL_TYPE_IDS = [
    34,     # Tritanium
    35,     # Pyerite
    36,     # Mexallon
    37,     # Isogen
    38,     # Nocxium
    40,     # Megacyte
    11399,  # Morphite
]

WEEKDAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def get_conn(db_file: str = DATABASE_FILE) -> sqlite3.Connection:
    if not Path(db_file).exists():
        raise SystemExit(f"Database not found: {db_file}")
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    return conn


def get_item_name(conn: sqlite3.Connection, type_id: int) -> str:
    cur = conn.execute("SELECT typeName FROM items WHERE typeID = ?", (type_id,))
    row = cur.fetchone()
    return row["typeName"] if row else f"typeID {type_id}"


def get_on_offer_type_ids(conn: sqlite3.Connection) -> list[int]:
    """
    Return distinct typeIDs for items currently tracked in on_offer_items.
    Table uses module_type_id (and module_name), not type_id.
    """
    cur = conn.execute(
        """
        SELECT DISTINCT module_type_id
        FROM on_offer_items
        ORDER BY module_type_id
        """
    )
    return [int(r["module_type_id"]) for r in cur.fetchall()]


def get_history_rows(
    conn: sqlite3.Connection,
    region_id: int,
    type_id: int,
) -> list[sqlite3.Row]:
    cur = conn.execute(
        """
        SELECT date_utc, lowest, highest, average, volume
        FROM market_history_daily
        WHERE region_id = ? AND type_id = ?
        ORDER BY date_utc
        """,
        (region_id, type_id),
    )
    return cur.fetchall()


def group_by_weekday(rows: list[sqlite3.Row]):
    """
    Group raw history rows by weekday.
    Returns dict weekday_index -> list[dict].
    """
    grouped = defaultdict(list)
    for r in rows:
        date_str = r["date_utc"]
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
        except Exception:
            # Fallback if stored with time component
            try:
                dt = datetime.fromisoformat(date_str.split("T", 1)[0])
            except Exception:
                continue
        wd = dt.weekday()  # 0=Mon .. 6=Sun
        lowest = r["lowest"]
        highest = r["highest"]
        avg = r["average"]
        vol = r["volume"]
        ev = expected_buy_order_volume_for_day(lowest, highest, avg, vol)
        grouped[wd].append(
            {
                "lowest": lowest or 0.0,
                "highest": highest or 0.0,
                "average": avg or 0.0,
                "volume": vol or 0.0,
                "expected_buy_vol": ev or 0.0,
            }
        )
    return grouped


def summarize_weekday_group(values: list[dict]) -> dict:
    if not values:
        return {}
    n = len(values)
    def avg(key: str) -> float:
        return sum(float(v.get(key, 0.0) or 0.0) for v in values) / n

    return {
        "n_days": n,
        "avg_lowest": avg("lowest"),
        "avg_highest": avg("highest"),
        "avg_average": avg("average"),
        "avg_volume": avg("volume"),
        "avg_expected_buy_vol": avg("expected_buy_vol"),
    }


def print_item_weekday_summary(conn: sqlite3.Connection, type_id: int, header: str | None = None) -> None:
    name = get_item_name(conn, type_id)
    rows = get_history_rows(conn, THE_FORGE_REGION_ID, type_id)
    if not rows:
        print(f"\n=== {name} (typeID {type_id}) – no history data ===")
        return
    grouped = group_by_weekday(rows)
    if header:
        print(f"\n=== {header}: {name} (typeID {type_id}) ===")
    else:
        print(f"\n=== {name} (typeID {type_id}) ===")
    print("Weekday  Days  Avg Price(low/avg/high)      Avg Volume   Avg Exp. Buy Vol")
    print("-------  ----  ---------------------------  ----------   -----------------")
    for wd in range(7):
        vals = summarize_weekday_group(grouped.get(wd, []))
        if not vals:
            continue
        lbl = WEEKDAY_NAMES[wd]
        n = vals["n_days"]
        lo = vals["avg_lowest"]
        av = vals["avg_average"]
        hi = vals["avg_highest"]
        vol = vals["avg_volume"]
        ev = vals["avg_expected_buy_vol"]
        print(
            f"{lbl:>7}  {n:4d}  "
            f"{lo:10.2f}/{av:10.2f}/{hi:10.2f}  "
            f"{vol:10.1f}   {ev:15.1f}"
        )


def main():
    conn = get_conn(DATABASE_FILE)
    try:
        print("Analyzing day-of-week patterns for selected minerals and On Offer modules "
              f"(region_id={THE_FORGE_REGION_ID})")

        print("\n### Minerals ###")
        for tid in MINERAL_TYPE_IDS:
            print_item_weekday_summary(conn, tid, header="Mineral")

        print("\n### Modules on offer ###")
        on_offer_type_ids = get_on_offer_type_ids(conn)
        if not on_offer_type_ids:
            print("No items in on_offer_items table.")
        else:
            for tid in on_offer_type_ids:
                print_item_weekday_summary(conn, tid, header="On Offer")
    finally:
        conn.close()


if __name__ == "__main__":
    main()

