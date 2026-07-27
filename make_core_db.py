"""
Generate the git-tracked *core* database snapshot: eve_manufacturing_core.db

The full runtime DB (eve_manufacturing.db) is dominated (~99%) by the
market_history_daily table (millions of daily price rows). That table is too
large for GitHub and is not required to run the launcher — it only improves
market decisions and can be rebuilt locally via fetch_market_history.py.

This script copies the runtime DB, empties market_history_daily (schema kept so
queries still work), strips SSO tokens (secrets must never be committed), then
VACUUMs to produce a small, portable core snapshot suitable for committing.

On another computer, the launcher will bootstrap eve_manufacturing.db from this
core snapshot automatically (see ensure_runtime_database in eve_launcher.py),
so a fresh `git clone` is immediately usable. Rebuild market history locally
afterwards for full functionality.

Usage:
    python make_core_db.py
"""

import os
import shutil
import sqlite3
import sys

try:
    from fetch_market_history import _create_table as _create_market_history_table
except Exception:
    _create_market_history_table = None

RUNTIME_DB = "eve_manufacturing.db"
CORE_DB = "eve_manufacturing_core.db"

# GitHub hard limit is 100 MB; warn well before that.
WARN_MB = 90.0


def build_core(runtime_db=RUNTIME_DB, core_db=CORE_DB):
    if not os.path.exists(runtime_db):
        print(f"ERROR: runtime database not found: {runtime_db}")
        return False

    src_mb = os.path.getsize(runtime_db) / 1e6
    print(f"Source runtime DB: {src_mb:,.1f} MB")

    if os.path.exists(core_db):
        os.remove(core_db)

    print(f"Copying -> {core_db} ...")
    shutil.copyfile(runtime_db, core_db)

    conn = sqlite3.connect(core_db)
    try:
        # 1) Drop the huge market-history table, then recreate it empty so that
        #    all existing queries keep working (no "no such table" errors).
        print("Emptying market_history_daily (kept empty; rebuild locally) ...")
        conn.execute("DROP TABLE IF EXISTS market_history_daily")
        if _create_market_history_table is not None:
            _create_market_history_table(conn)
        else:
            # Fallback: recreate minimal schema if fetch_market_history is unavailable.
            conn.execute(
                """CREATE TABLE IF NOT EXISTS market_history_daily (
                       region_id INTEGER NOT NULL, type_id INTEGER NOT NULL,
                       type_name TEXT, date_utc TEXT NOT NULL, average REAL NOT NULL,
                       highest REAL NOT NULL, lowest REAL NOT NULL,
                       order_count INTEGER, volume INTEGER,
                       PRIMARY KEY (region_id, type_id, date_utc))"""
            )

        # 2) Strip SSO secrets — refresh/access tokens must never be committed.
        try:
            conn.execute(
                "UPDATE sso_character SET refresh_token=NULL, access_token=NULL, "
                "access_token_expires_at=NULL"
            )
            print("Cleared SSO tokens from sso_character (re-authenticate on the new PC).")
        except sqlite3.OperationalError:
            pass  # table may not exist in older DBs

        conn.commit()

        print("VACUUM ...")
        conn.execute("VACUUM")
        conn.commit()
    finally:
        conn.close()

    core_mb = os.path.getsize(core_db) / 1e6
    print(f"\nDone. Core snapshot: {core_db}  ({core_mb:,.1f} MB)")
    if core_mb > WARN_MB:
        print(
            f"WARNING: core is {core_mb:,.1f} MB, approaching GitHub's 100 MB limit. "
            "Consider trimming other large tables."
        )
    print(
        "\nCommit this file to share it. On another PC the launcher recreates "
        "eve_manufacturing.db from it automatically; rebuild market history there "
        "with fetch_market_history.py for full data."
    )
    return True


if __name__ == "__main__":
    ok = build_core()
    sys.exit(0 if ok else 1)
