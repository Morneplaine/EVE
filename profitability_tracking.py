"""
Production profitability tracking with FIFO cost basis.

Replays ESI-synced events (wallet transactions + finished industry jobs) in
chronological order to build a per-character lot ledger, then aggregates per-job
P/L for the Production P/L view in the launcher.

Tables created/owned by this module (in eve_manufacturing.db):
    profitability_lot           - one row per opened inventory lot
    profitability_consumption   - one row per qty drained from a lot
    profitability_job           - one row per processed industry job (cached aggregates)

The ledger is fully derived from raw ESI tables (esi_wallet_transactions,
esi_corporation_wallet_transactions for corp entities, esi_industry_jobs /
esi_corporation_industry_jobs) plus the static blueprints / manufacturing_materials
tables.
It is safe to rebuild any time; rebuilding is idempotent for a given character.
"""

from __future__ import annotations

import logging
import math
import sqlite3
from typing import Callable, Optional

from decryptor_profitability import (
    compare_decryptor_profitability,
    estimate_t2_invention_from_bindings,
)

logger = logging.getLogger(__name__)

# Industry activities we treat as "manufacturing" for output lots.
# 1 = Manufacturing, 9 = Reactions (output is a product like manufacturing).
MANUFACTURING_ACTIVITY_IDS = {1, 9}

# Statuses that imply the job consumed materials AND produced output.
# We do not process active/paused/cancelled jobs (no realized output yet).
COMPLETED_JOB_STATUSES = {"delivered"}

# The Forge - region id of Jita; matches MARKET_HISTORY_REGION_ID in eve_launcher.
MARKET_HISTORY_REGION_ID = 10000002


def ensure_profitability_tables(conn: sqlite3.Connection) -> None:
    """Create the profitability tables if missing. Safe to call repeatedly.

    The `character_id` column actually stores the **entity_id** which is either
    a character_id (entity_kind='character') or a corporation_id (entity_kind='corporation').
    The legacy column name is kept for back-compat; new code should use the entity_id alias.
    """
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS profitability_lot (
            lot_id INTEGER PRIMARY KEY AUTOINCREMENT,
            character_id INTEGER NOT NULL,
            type_id INTEGER NOT NULL,
            source_kind TEXT NOT NULL,    -- 'buy_tx' | 'manufacture' | 'unknown'
            source_id INTEGER,            -- transaction_id or job_id; NULL for synthetic
            opened_at TEXT NOT NULL,      -- iso utc
            qty_initial INTEGER NOT NULL,
            qty_remaining INTEGER NOT NULL,
            unit_cost REAL,               -- NULL = unknown basis (synthetic lot)
            note TEXT,
            entity_kind TEXT NOT NULL DEFAULT 'character'
        );
        CREATE INDEX IF NOT EXISTS idx_pl_lot_lookup
            ON profitability_lot(character_id, type_id, opened_at);

        CREATE TABLE IF NOT EXISTS profitability_consumption (
            consumption_id INTEGER PRIMARY KEY AUTOINCREMENT,
            lot_id INTEGER NOT NULL,
            consumer_kind TEXT NOT NULL,  -- 'sale_tx' | 'manufacture_input'
            consumer_id INTEGER,          -- transaction_id or job_id
            qty INTEGER NOT NULL,
            unit_cost REAL,               -- NULL when lot was synthetic/unknown
            consumed_at TEXT NOT NULL,
            FOREIGN KEY (lot_id) REFERENCES profitability_lot(lot_id)
        );
        CREATE INDEX IF NOT EXISTS idx_pl_cons_consumer
            ON profitability_consumption(consumer_kind, consumer_id);
        CREATE INDEX IF NOT EXISTS idx_pl_cons_lot
            ON profitability_consumption(lot_id);

        CREATE TABLE IF NOT EXISTS profitability_job (
            character_id INTEGER NOT NULL,
            job_id INTEGER NOT NULL,
            activity_id INTEGER,
            blueprint_type_id INTEGER,
            product_type_id INTEGER,
            product_name TEXT,
            runs INTEGER,
            output_qty INTEGER,
            start_date_utc TEXT,
            end_date_utc TEXT,
            status TEXT,
            job_fee REAL,
            materials_cost REAL,
            invention_cost REAL DEFAULT 0, -- amortized T2 invention (expected per BPC × runs used)
            facility_cost REAL DEFAULT 0,  -- optional production_cost_per_run × mfg runs
            materials_unknown_qty INTEGER,
            unit_cost REAL,                -- (job_fee + mats + invention + facility) / output_qty
            sold_qty INTEGER,
            revenue REAL,
            cogs_for_sold REAL,
            realized_profit REAL,
            unsold_qty INTEGER,
            unrealized_value REAL,         -- estimated (unsold_qty * current sell price)
            note TEXT,
            entity_kind TEXT NOT NULL DEFAULT 'character',
            PRIMARY KEY (character_id, job_id)
        );
        """
    )
    # Idempotent migration: add entity_kind to old DBs.
    for tbl in ("profitability_lot", "profitability_job"):
        cols = {row[1] for row in conn.execute(f"PRAGMA table_info({tbl})").fetchall()}
        if "entity_kind" not in cols:
            conn.execute(f"ALTER TABLE {tbl} ADD COLUMN entity_kind TEXT NOT NULL DEFAULT 'character'")
    pj_cols = {row[1] for row in conn.execute("PRAGMA table_info(profitability_job)").fetchall()}
    if "invention_cost" not in pj_cols:
        conn.execute("ALTER TABLE profitability_job ADD COLUMN invention_cost REAL DEFAULT 0")
    if "facility_cost" not in pj_cols:
        conn.execute("ALTER TABLE profitability_job ADD COLUMN facility_cost REAL DEFAULT 0")
    # Snapshots cache: stable cost-basis fallback per (type, date).
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS profitability_price_snapshot (
            type_id INTEGER NOT NULL,
            snapshot_date TEXT NOT NULL,        -- YYYY-MM-DD (UTC)
            unit_cost REAL NOT NULL,
            source TEXT NOT NULL,               -- 'market_history' | 'current_sell_min' | 'current_buy_max' | 'manual'
            captured_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (type_id, snapshot_date)
        )
        """
    )
    _ensure_blueprint_datacore_bindings(conn)
    conn.commit()


def _ensure_blueprint_datacore_bindings(conn: sqlite3.Connection) -> None:
    """Same table as launcher Decryptor tab (T2 invention parameters per blueprint)."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS blueprint_datacore_bindings (
            blueprint_type_id INTEGER PRIMARY KEY,
            dc1_name TEXT,
            dc1_qty INTEGER NOT NULL DEFAULT 0,
            dc2_name TEXT,
            dc2_qty INTEGER NOT NULL DEFAULT 0,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cols = {row[1] for row in conn.execute("PRAGMA table_info(blueprint_datacore_bindings)").fetchall()}
    for col, typ in [
        ("base_invention_chance_pct", "REAL"),
        ("invention_cost_per_attempt", "REAL"),
        ("base_bpc_runs", "INTEGER"),
        ("research_time_days", "REAL"),
        ("research_time_hours", "REAL"),
        ("research_time_minutes", "REAL"),
        ("production_cost_per_run", "REAL"),
    ]:
        if col not in cols:
            conn.execute(f"ALTER TABLE blueprint_datacore_bindings ADD COLUMN {col} {typ}")


def clear_price_snapshots(conn: sqlite3.Connection) -> int:
    """Wipe the snapshot cache so the next rebuild re-resolves all fallback prices."""
    ensure_profitability_tables(conn)
    n = conn.execute("SELECT COUNT(*) FROM profitability_price_snapshot").fetchone()[0]
    conn.execute("DELETE FROM profitability_price_snapshot")
    conn.commit()
    return int(n or 0)


def _material_qty_total(base_qty: int, runs: int, me_percent: float) -> int:
    """EVE formula: total_qty = max(runs, ceil(base * (1 - ME%) * runs)). Matches the launcher's other uses."""
    me_fraction = max(0.0, min(100.0, float(me_percent))) / 100.0
    return max(int(runs), int(math.ceil(base_qty * (1.0 - me_fraction) * int(runs))))


def _is_t2_blueprint(conn: sqlite3.Connection, blueprint_type_id: int) -> bool:
    row = conn.execute(
        "SELECT 1 FROM invention_recipes WHERE t2_blueprint_type_id = ? LIMIT 1",
        (int(blueprint_type_id),),
    ).fetchone()
    return row is not None


def _bp_lookup(conn: sqlite3.Connection, blueprint_type_id: int) -> Optional[dict]:
    row = conn.execute(
        "SELECT productTypeID, productName, outputQuantity FROM blueprints WHERE blueprintTypeID = ?",
        (int(blueprint_type_id),),
    ).fetchone()
    if not row:
        return None
    return {"productTypeID": int(row[0]), "productName": row[1], "outputQuantity": int(row[2])}


def _materials_for_blueprint(
    conn: sqlite3.Connection, blueprint_type_id: int, runs: int, me_percent: float
) -> list[dict]:
    """Return [{material_type_id, material_name, qty}] consumed for `runs` runs at the given ME."""
    rows = conn.execute(
        "SELECT materialTypeID, materialName, quantity FROM manufacturing_materials WHERE blueprintTypeID = ?",
        (int(blueprint_type_id),),
    ).fetchall()
    out = []
    for r in rows:
        out.append({
            "material_type_id": int(r[0]),
            "material_name": r[1],
            "qty": _material_qty_total(int(r[2]), int(runs), me_percent),
        })
    return out


def _open_lot(
    conn: sqlite3.Connection,
    entity_id: int,
    entity_kind: str,
    type_id: int,
    source_kind: str,
    source_id: Optional[int],
    opened_at: str,
    qty: int,
    unit_cost: Optional[float],
    note: Optional[str] = None,
) -> int:
    cur = conn.execute(
        """
        INSERT INTO profitability_lot
            (character_id, entity_kind, type_id, source_kind, source_id, opened_at,
             qty_initial, qty_remaining, unit_cost, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            int(entity_id), entity_kind, int(type_id), source_kind,
            int(source_id) if source_id is not None else None,
            opened_at, int(qty), int(qty), unit_cost, note,
        ),
    )
    return int(cur.lastrowid)


def _date_key(when_iso: str) -> str:
    """Return YYYY-MM-DD (UTC) from an ISO timestamp like '2026-04-15T12:34:56Z'."""
    if not when_iso:
        return ""
    s = str(when_iso)
    return s[:10] if len(s) >= 10 else s


def _get_or_capture_fallback_price(
    conn: sqlite3.Connection,
    type_id: int,
    when_iso: str,
) -> tuple[Optional[float], Optional[str]]:
    """
    Return a stable fallback unit cost for (type_id, date) and the source string.

    Resolution order (only the first time per (type_id, date)):
      1. profitability_price_snapshot - reuse already-captured value
      2. market_history_daily.average for the exact date  (source='market_history')
      3. prices.sell_min                                   (source='current_sell_min')
      4. prices.buy_max                                    (source='current_buy_max')
    The resolved (price, source) is stored in profitability_price_snapshot, so
    later rebuilds reuse it and produce stable numbers.

    Returns (None, None) if no price can be found at all.
    """
    date = _date_key(when_iso)
    if not date or not type_id:
        return None, None

    snap = conn.execute(
        "SELECT unit_cost, source FROM profitability_price_snapshot "
        " WHERE type_id = ? AND snapshot_date = ?",
        (int(type_id), date),
    ).fetchone()
    if snap is not None:
        return float(snap[0]), str(snap[1])

    hist = conn.execute(
        "SELECT average FROM market_history_daily "
        " WHERE region_id = ? AND type_id = ? AND date_utc = ?",
        (int(MARKET_HISTORY_REGION_ID), int(type_id), date),
    ).fetchone()
    if hist and hist[0]:
        price = float(hist[0])
        if price > 0:
            conn.execute(
                "INSERT OR REPLACE INTO profitability_price_snapshot "
                " (type_id, snapshot_date, unit_cost, source) VALUES (?, ?, ?, ?)",
                (int(type_id), date, price, "market_history"),
            )
            return price, "market_history"

    cur = conn.execute(
        "SELECT sell_min, buy_max FROM prices WHERE typeID = ?",
        (int(type_id),),
    ).fetchone()
    if cur:
        sell_min = float(cur[0]) if cur[0] else 0.0
        buy_max = float(cur[1]) if cur[1] else 0.0
        if sell_min > 0:
            conn.execute(
                "INSERT OR REPLACE INTO profitability_price_snapshot "
                " (type_id, snapshot_date, unit_cost, source) VALUES (?, ?, ?, ?)",
                (int(type_id), date, sell_min, "current_sell_min"),
            )
            return sell_min, "current_sell_min"
        if buy_max > 0:
            conn.execute(
                "INSERT OR REPLACE INTO profitability_price_snapshot "
                " (type_id, snapshot_date, unit_cost, source) VALUES (?, ?, ?, ?)",
                (int(type_id), date, buy_max, "current_buy_max"),
            )
            return buy_max, "current_buy_max"
    return None, None


def _consume_fifo(
    conn: sqlite3.Connection,
    entity_id: int,
    entity_kind: str,
    type_id: int,
    qty_to_consume: int,
    consumer_kind: str,
    consumer_id: Optional[int],
    consumed_at: str,
    use_market_fallback: bool = True,
) -> tuple[float, int, int]:
    """
    Drain `qty_to_consume` from oldest open lots first.
    If lots run out, opens a synthetic lot for the remainder so that later
    inflow doesn't get retro-attributed. The synthetic lot's unit cost is
    pulled from the current `prices` table (sell_min, falling back to buy_max)
    when `use_market_fallback` is True; otherwise it's NULL.

    Returns (total_cost, qty_with_known_cost, qty_with_truly_unknown_cost).
    Market-fallback units count as "known" because they have a real cost number,
    but they're tagged on the synthetic lot via the `source_kind` value.
    """
    remaining = int(qty_to_consume)
    total_cost = 0.0
    known_qty = 0
    unknown_qty = 0
    while remaining > 0:
        row = conn.execute(
            """
            SELECT lot_id, qty_remaining, unit_cost
              FROM profitability_lot
             WHERE character_id = ? AND entity_kind = ? AND type_id = ? AND qty_remaining > 0
             ORDER BY opened_at ASC, lot_id ASC
             LIMIT 1
            """,
            (int(entity_id), entity_kind, int(type_id)),
        ).fetchone()
        if not row:
            if use_market_fallback:
                mkt_price, mkt_source = _get_or_capture_fallback_price(conn, type_id, consumed_at)
            else:
                mkt_price, mkt_source = None, None
            if mkt_price is not None:
                synth_kind = "market_fallback"
                note = f"synthetic - {mkt_source} on {_date_key(consumed_at)} = {mkt_price:.2f}"
            else:
                synth_kind = "unknown"
                note = "synthetic - no buy lot and no market price"
            synth_lot = _open_lot(
                conn, entity_id, entity_kind, type_id, synth_kind, None,
                consumed_at, remaining, mkt_price, note=note,
            )
            conn.execute(
                "UPDATE profitability_lot SET qty_remaining = 0 WHERE lot_id = ?",
                (synth_lot,),
            )
            conn.execute(
                """
                INSERT INTO profitability_consumption
                    (lot_id, consumer_kind, consumer_id, qty, unit_cost, consumed_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (synth_lot, consumer_kind,
                 int(consumer_id) if consumer_id is not None else None,
                 int(remaining), mkt_price, consumed_at),
            )
            if mkt_price is not None:
                total_cost += mkt_price * remaining
                known_qty += remaining
            else:
                unknown_qty += remaining
            remaining = 0
            break
        lot_id, qty_remaining, unit_cost = int(row[0]), int(row[1]), row[2]
        take = min(remaining, qty_remaining)
        if unit_cost is not None:
            total_cost += float(unit_cost) * take
            known_qty += take
        else:
            unknown_qty += take
        conn.execute(
            "UPDATE profitability_lot SET qty_remaining = qty_remaining - ? WHERE lot_id = ?",
            (take, lot_id),
        )
        conn.execute(
            """
            INSERT INTO profitability_consumption
                (lot_id, consumer_kind, consumer_id, qty, unit_cost, consumed_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (lot_id, consumer_kind,
             int(consumer_id) if consumer_id is not None else None,
             int(take), unit_cost, consumed_at),
        )
        remaining -= take
    return total_cost, known_qty, unknown_qty


def _clear_entity_ledger(conn: sqlite3.Connection, entity_id: int, entity_kind: str) -> None:
    conn.execute(
        """
        DELETE FROM profitability_consumption
         WHERE lot_id IN (SELECT lot_id FROM profitability_lot
                          WHERE character_id = ? AND entity_kind = ?)
        """,
        (int(entity_id), entity_kind),
    )
    conn.execute(
        "DELETE FROM profitability_lot WHERE character_id = ? AND entity_kind = ?",
        (int(entity_id), entity_kind),
    )
    conn.execute(
        "DELETE FROM profitability_job WHERE character_id = ? AND entity_kind = ?",
        (int(entity_id), entity_kind),
    )


def rebuild_ledger(
    conn: sqlite3.Connection,
    entity_id: int,
    default_me_percent: float = 10.0,
    log: Optional[Callable[[str], None]] = None,
    entity_kind: str = "character",
    use_market_fallback: bool = True,
) -> dict:
    """
    Rebuild the FIFO ledger for one entity (character or corporation) from scratch.

    For entity_kind='character': wallet+jobs of that single character (personal wallet only).
    For entity_kind='corporation': union of personal wallets of all linked corp members
      (sso_character.corporation_id = entity), plus any SSO-linked character who appears as
      installer_id on that corp's industry jobs (covers missing/stale corporation_id on alts),
      corporation wallet market transactions (all divisions synced to DB), plus the
      corporation's industry jobs.

    Returns counts dict: {jobs_processed, lots_opened, sales_matched, unknown_units}.
    """
    if entity_kind not in ("character", "corporation"):
        raise ValueError(f"Unknown entity_kind: {entity_kind}")
    ensure_profitability_tables(conn)
    log = log or (lambda _msg: None)
    if entity_kind == "character":
        wallet_chars = [int(entity_id)]
    else:
        rows = conn.execute(
            "SELECT character_id FROM sso_character WHERE corporation_id = ?",
            (int(entity_id),),
        ).fetchall()
        wallet_chars = [int(r[0]) for r in rows]
        for r in conn.execute(
            """
            SELECT DISTINCT j.installer_id
              FROM esi_corporation_industry_jobs j
             WHERE j.corporation_id = ?
               AND j.installer_id IS NOT NULL
               AND j.installer_id IN (SELECT character_id FROM sso_character)
            """,
            (int(entity_id),),
        ).fetchall():
            iid = int(r[0])
            if iid not in wallet_chars:
                wallet_chars.append(iid)
        if len(wallet_chars) > len(rows):
            extra = len(wallet_chars) - len(rows)
            log(
                f"  Including {extra} extra linked character(s) from corp job installer_id "
                f"(personal wallet) not matched by sso_character.corporation_id."
            )
        if not wallet_chars:
            log(f"  No linked characters for corporation {entity_id}; corp jobs without buy lots will all be unknown-basis.")

    events: list[tuple[str, str, dict]] = []

    if wallet_chars:
        placeholders = ",".join("?" * len(wallet_chars))
        for r in conn.execute(
            f"""
            SELECT transaction_id, date_utc, type_id, quantity, unit_price, is_buy
              FROM esi_wallet_transactions
             WHERE character_id IN ({placeholders})
            """,
            wallet_chars,
        ).fetchall():
            ev = {
                "transaction_id": int(r[0]),
                "date_utc": r[1],
                "type_id": int(r[2]) if r[2] is not None else None,
                "quantity": int(r[3]) if r[3] is not None else 0,
                "unit_price": float(r[4]) if r[4] is not None else 0.0,
                "is_buy": int(r[5]) if r[5] is not None else 0,
            }
            if not ev["type_id"] or ev["quantity"] <= 0:
                continue
            events.append(("buy_tx" if ev["is_buy"] else "sale_tx", ev["date_utc"], ev))

    if entity_kind == "corporation":
        for r in conn.execute(
            """
            SELECT transaction_id, date_utc, type_id, quantity, unit_price, is_buy
              FROM esi_corporation_wallet_transactions
             WHERE corporation_id = ?
            """,
            (int(entity_id),),
        ).fetchall():
            ev = {
                "transaction_id": int(r[0]),
                "date_utc": r[1],
                "type_id": int(r[2]) if r[2] is not None else None,
                "quantity": int(r[3]) if r[3] is not None else 0,
                "unit_price": float(r[4]) if r[4] is not None else 0.0,
                "is_buy": int(r[5]) if r[5] is not None else 0,
            }
            if not ev["type_id"] or ev["quantity"] <= 0:
                continue
            events.append(("buy_tx" if ev["is_buy"] else "sale_tx", ev["date_utc"], ev))

    if entity_kind == "character":
        job_rows = conn.execute(
            """
            SELECT job_id, activity_id, blueprint_type_id, product_type_id, runs, cost,
                   status, start_date_utc, end_date_utc, completed_date_utc
              FROM esi_industry_jobs
             WHERE character_id = ?
            """,
            (int(entity_id),),
        ).fetchall()
    else:
        job_rows = conn.execute(
            """
            SELECT job_id, activity_id, blueprint_type_id, product_type_id, runs, cost,
                   status, start_date_utc, end_date_utc, completed_date_utc
              FROM esi_corporation_industry_jobs
             WHERE corporation_id = ?
            """,
            (int(entity_id),),
        ).fetchall()

    for r in job_rows:
        job = {
            "job_id": int(r[0]),
            "activity_id": int(r[1]) if r[1] is not None else None,
            "blueprint_type_id": int(r[2]) if r[2] is not None else None,
            "product_type_id": int(r[3]) if r[3] is not None else None,
            "runs": int(r[4]) if r[4] is not None else 0,
            "cost": float(r[5]) if r[5] is not None else 0.0,
            "status": r[6] or "",
            "start_date_utc": r[7],
            "end_date_utc": r[8],
            "completed_date_utc": r[9],
        }
        if job["activity_id"] not in MANUFACTURING_ACTIVITY_IDS:
            continue
        if (job["status"] or "").lower() not in COMPLETED_JOB_STATUSES:
            continue
        if not job["blueprint_type_id"] or job["runs"] <= 0:
            continue
        consume_at = job["start_date_utc"] or job["end_date_utc"] or job["completed_date_utc"]
        produce_at = job["completed_date_utc"] or job["end_date_utc"] or consume_at
        if not consume_at or not produce_at:
            continue
        events.append(("job_consume", consume_at, job))
        events.append(("job_produce", produce_at, job))

    invention_cache: dict[int, dict | None] = {}
    t2_bp_ids: set[int] = set()
    for _k, _w, ev in events:
        if _k in ("job_consume", "job_produce") and ev.get("blueprint_type_id"):
            bpid = int(ev["blueprint_type_id"])
            if _is_t2_blueprint(conn, bpid):
                t2_bp_ids.add(bpid)
    for bpid in t2_bp_ids:
        bp_info = _bp_lookup(conn, bpid)
        if bp_info:
            invention_cache[bpid] = estimate_t2_invention_from_bindings(
                conn, bpid, bp_info.get("productName") or "",
            )
        else:
            invention_cache[bpid] = None

    _clear_entity_ledger(conn, entity_id, entity_kind)

    kind_priority = {"buy_tx": 0, "job_produce": 1, "job_consume": 2, "sale_tx": 3}
    events.sort(key=lambda e: (e[1] or "", kind_priority.get(e[0], 9)))

    log(f"Replaying {len(events)} events for {entity_kind} {entity_id}...")

    job_state: dict[int, dict] = {}
    counts = {"jobs_processed": 0, "lots_opened": 0, "sales_matched": 0, "unknown_units": 0}

    for kind, when, ev in events:
        if kind == "buy_tx":
            _open_lot(
                conn, entity_id, entity_kind, ev["type_id"], "buy_tx",
                ev["transaction_id"], when, ev["quantity"], ev["unit_price"],
            )
            counts["lots_opened"] += 1
        elif kind == "sale_tx":
            cost, _known, unknown = _consume_fifo(
                conn, entity_id, entity_kind, ev["type_id"], ev["quantity"],
                "sale_tx", ev["transaction_id"], when,
                use_market_fallback=use_market_fallback,
            )
            counts["sales_matched"] += 1
            counts["unknown_units"] += unknown
        elif kind == "job_consume":
            bp = _bp_lookup(conn, ev["blueprint_type_id"])
            if not bp:
                log(f"  [job {ev['job_id']}] no blueprint row for bpid={ev['blueprint_type_id']}; skipping")
                continue
            me_for_job = float(default_me_percent)
            invention_cost = 0.0
            facility_cost = 0.0
            invention_note: str | None = None
            if _is_t2_blueprint(conn, ev["blueprint_type_id"]):
                inv = invention_cache.get(int(ev["blueprint_type_id"]))
                if inv:
                    me_for_job = float(inv["bpc_me"])
                    bpc_runs = max(1, int(inv["bpc_runs"]))
                    invention_cost = float(inv["expected_inv_cost"]) * (
                        int(ev["runs"]) / bpc_runs
                    )
                    ppr = inv.get("production_cost_per_run")
                    if ppr is not None and float(ppr) > 0:
                        facility_cost = float(ppr) * int(ev["runs"])
                    invention_note = inv.get("decryptor_name")
                else:
                    invention_note = "no invention binding (Decryptor tab)"

            mats = _materials_for_blueprint(
                conn, ev["blueprint_type_id"], ev["runs"], me_for_job
            )
            mats_cost = 0.0
            mats_unknown = 0
            for m in mats:
                c, _kn, un = _consume_fifo(
                    conn, entity_id, entity_kind, m["material_type_id"], m["qty"],
                    "manufacture_input", ev["job_id"], when,
                    use_market_fallback=use_market_fallback,
                )
                mats_cost += c
                mats_unknown += un
            output_qty = bp["outputQuantity"] * int(ev["runs"])
            unit_cost = (
                (float(ev["cost"] or 0.0) + mats_cost + invention_cost + facility_cost)
                / output_qty
                if output_qty > 0
                else 0.0
            )
            job_state[ev["job_id"]] = {
                "ev": ev,
                "bp": bp,
                "materials_cost": mats_cost,
                "materials_unknown_qty": mats_unknown,
                "invention_cost": invention_cost,
                "facility_cost": facility_cost,
                "invention_note": invention_note,
                "output_qty": output_qty,
                "unit_cost": unit_cost,
                "lot_id": None,
                "consume_at": when,
            }
        elif kind == "job_produce":
            st = job_state.get(ev["job_id"])
            if not st:
                continue
            if not st["ev"]["product_type_id"]:
                continue
            note_parts = [f"runs={st['ev']['runs']}"]
            if st.get("invention_note"):
                note_parts.append(f"inv={st['invention_note']}")
            lot_id = _open_lot(
                conn, entity_id, entity_kind, st["ev"]["product_type_id"], "manufacture",
                ev["job_id"], when, st["output_qty"], st["unit_cost"],
                note="; ".join(note_parts),
            )
            st["lot_id"] = lot_id
            counts["jobs_processed"] += 1
            counts["lots_opened"] += 1

    _populate_job_aggregates(conn, entity_id, entity_kind, wallet_chars, job_state)
    conn.commit()
    log(
        f"Done: jobs={counts['jobs_processed']}, lots opened={counts['lots_opened']}, "
        f"sales matched={counts['sales_matched']}, unknown-basis units={counts['unknown_units']}"
    )
    return counts


def _populate_job_aggregates(
    conn: sqlite3.Connection,
    entity_id: int,
    entity_kind: str,
    wallet_chars: list[int],
    job_state: dict[int, dict],
) -> None:
    """Write one summary row per processed job into profitability_job."""
    for job_id, st in job_state.items():
        ev = st["ev"]
        bp = st["bp"]
        lot_id = st.get("lot_id")
        sold_qty = 0
        revenue = 0.0
        cogs_for_sold = 0.0
        unsold_qty = st["output_qty"]
        if lot_id is not None:
            row = conn.execute(
                "SELECT qty_remaining FROM profitability_lot WHERE lot_id = ?",
                (int(lot_id),),
            ).fetchone()
            unsold_qty = int(row[0]) if row else 0
            join_parts: list[str] = []
            bind: list = []
            price_cols: list[str] = []
            if wallet_chars:
                ph = ",".join("?" * len(wallet_chars))
                join_parts.append(
                    f"""LEFT JOIN esi_wallet_transactions et
                            ON et.transaction_id = pc.consumer_id
                           AND et.character_id IN ({ph})"""
                )
                bind.extend(wallet_chars)
                price_cols.append("et.unit_price")
            if entity_kind == "corporation":
                join_parts.append(
                    """LEFT JOIN esi_corporation_wallet_transactions cwt
                            ON cwt.transaction_id = pc.consumer_id
                           AND cwt.corporation_id = ?"""
                )
                bind.append(int(entity_id))
                price_cols.append("cwt.unit_price")
            sale_price_sql = (
                "COALESCE(" + ", ".join(price_cols) + ")"
                if len(price_cols) > 1
                else (price_cols[0] if price_cols else "NULL")
            )
            joins = "\n".join(join_parts)
            sql = f"""
                SELECT pc.qty, pc.unit_cost, {sale_price_sql}
                  FROM profitability_consumption pc
                  {joins}
                 WHERE pc.lot_id = ? AND pc.consumer_kind = 'sale_tx'
            """
            bind.append(int(lot_id))
            sales = conn.execute(sql, bind).fetchall()
            for s in sales:
                q = int(s[0])
                uc = float(s[1]) if s[1] is not None else float(st["unit_cost"] or 0.0)
                up = float(s[2]) if s[2] is not None else 0.0
                sold_qty += q
                revenue += up * q
                cogs_for_sold += uc * q
        realized_profit = revenue - cogs_for_sold
        product_type_id = ev["product_type_id"] or bp["productTypeID"]
        unrealized_value = 0.0
        if product_type_id and unsold_qty > 0:
            pr = conn.execute(
                "SELECT sell_min FROM prices WHERE typeID = ?",
                (int(product_type_id),),
            ).fetchone()
            if pr and pr[0]:
                unrealized_value = float(pr[0]) * unsold_qty
        inv_cost = float(st.get("invention_cost") or 0.0)
        fac_cost = float(st.get("facility_cost") or 0.0)
        note = st.get("invention_note")
        if note and inv_cost <= 0 and _is_t2_blueprint(conn, ev["blueprint_type_id"]):
            note = f"{note}; invention cost not applied"
        conn.execute(
            """
            INSERT OR REPLACE INTO profitability_job
                (character_id, entity_kind, job_id, activity_id, blueprint_type_id, product_type_id,
                 product_name, runs, output_qty, start_date_utc, end_date_utc, status,
                 job_fee, materials_cost, invention_cost, facility_cost, materials_unknown_qty, unit_cost,
                 sold_qty, revenue, cogs_for_sold, realized_profit,
                 unsold_qty, unrealized_value, note)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(entity_id), entity_kind, int(job_id), ev["activity_id"], ev["blueprint_type_id"],
                product_type_id, bp["productName"], ev["runs"], st["output_qty"],
                ev["start_date_utc"], ev["completed_date_utc"] or ev["end_date_utc"],
                ev["status"], float(ev["cost"] or 0.0), float(st["materials_cost"]),
                inv_cost, fac_cost,
                int(st["materials_unknown_qty"]), float(st["unit_cost"]),
                int(sold_qty), float(revenue), float(cogs_for_sold), float(realized_profit),
                int(unsold_qty), float(unrealized_value), note,
            ),
        )


def list_production_pl(
    conn: sqlite3.Connection,
    entity_id: Optional[int] = None,
    entity_kind: Optional[str] = None,
    search: Optional[str] = None,
    limit: int = 2000,
) -> list[dict]:
    """Return production P/L rows, newest first.

    entity_id=None returns all entities. If entity_kind is also None, both
    character ledgers and corporation ledgers are included.
    """
    ensure_profitability_tables(conn)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sso_character (
            character_id INTEGER PRIMARY KEY,
            character_name TEXT,
            corporation_id INTEGER,
            corporation_name TEXT
        )
        """
    )
    where = []
    args: list = []
    if entity_id is not None:
        where.append("pj.character_id = ?")
        args.append(int(entity_id))
    if entity_kind is not None:
        where.append("pj.entity_kind = ?")
        args.append(entity_kind)
    if search:
        where.append("LOWER(pj.product_name) LIKE ?")
        args.append(f"%{search.lower().strip()}%")
    where_sql = ("WHERE " + " AND ".join(where)) if where else ""
    args.append(int(limit))
    rows = conn.execute(
        f"""
        SELECT pj.character_id, pj.entity_kind, sc.character_name,
               corp.corporation_name, pj.job_id, pj.activity_id,
               pj.product_type_id, pj.product_name, pj.runs, pj.output_qty, pj.end_date_utc,
               pj.status, pj.job_fee, pj.materials_cost, pj.invention_cost, pj.facility_cost,
               pj.materials_unknown_qty,
               pj.unit_cost, pj.sold_qty, pj.revenue, pj.cogs_for_sold,
               pj.realized_profit, pj.unsold_qty, pj.unrealized_value, pj.note
          FROM profitability_job pj
          LEFT JOIN sso_character sc
                 ON pj.entity_kind = 'character' AND sc.character_id = pj.character_id
          LEFT JOIN (
              SELECT corporation_id, MAX(corporation_name) AS corporation_name
                FROM sso_character
               WHERE corporation_id IS NOT NULL
               GROUP BY corporation_id
          ) corp
                 ON pj.entity_kind = 'corporation' AND corp.corporation_id = pj.character_id
        {where_sql}
        ORDER BY pj.end_date_utc DESC
        LIMIT ?
        """,
        args,
    ).fetchall()
    out = []
    for r in rows:
        kind = r[1] or "character"
        if kind == "corporation":
            owner_name = (r[3] or f"Corp {r[0]}") + " (corp)"
        else:
            owner_name = r[2] or "?"
        job_fee = float(r[12] or 0)
        mat_cost = float(r[13] or 0)
        inv_cost = float(r[14] or 0)
        fac_cost = float(r[15] or 0)
        mats_unknown = int(r[16] or 0)
        unit_cost = float(r[17] or 0)
        sold_qty = int(r[18] or 0)
        revenue = float(r[19] or 0)
        cogs = float(r[20] or 0)
        realized = float(r[21] or 0)
        unsold_qty = int(r[22] or 0)
        unrealized = float(r[23] or 0)
        note = r[24] or ""
        total_cost = job_fee + mat_cost + inv_cost + fac_cost
        total_pl = realized + unrealized - (unit_cost * unsold_qty)
        out.append({
            "entity_id": r[0],
            "entity_kind": kind,
            "owner_name": owner_name,
            "character_name": owner_name,
            "job_id": r[4],
            "activity_id": r[5],
            "product_type_id": int(r[6]) if r[6] is not None else None,
            "product_name": r[7] or "?",
            "runs": r[8] or 0,
            "output_qty": r[9] or 0,
            "end_date_utc": r[10] or "",
            "status": r[11] or "",
            "job_fee": job_fee,
            "materials_cost": mat_cost,
            "invention_cost": inv_cost,
            "facility_cost": fac_cost,
            "materials_unknown_qty": mats_unknown,
            "total_cost": total_cost,
            "unit_cost": unit_cost,
            "sold_qty": sold_qty,
            "revenue": revenue,
            "cogs_for_sold": cogs,
            "realized_profit": realized,
            "unsold_qty": unsold_qty,
            "unrealized_value": unrealized,
            "note": note,
            "total_estimated_pl": total_pl,
            "view_mode": "per_job",
            "job_count": 1,
            "job_ids": [int(r[4])],
            "week_key": _iso_week_key(r[10] or ""),
        })
    return out


def _iso_week_key(end_date_utc: str) -> str:
    """ISO calendar week key YYYY-Www from an ESI UTC timestamp."""
    if not end_date_utc:
        return "unknown"
    s = str(end_date_utc).replace("Z", "+00:00")
    try:
        from datetime import datetime
        dt = datetime.fromisoformat(s)
        y, w, _ = dt.isocalendar()
        return f"{y}-W{w:02d}"
    except ValueError:
        return s[:10] if len(s) >= 10 else "unknown"


def _group_product_key(row: dict) -> tuple:
    """Stable grouping key for the same manufactured product."""
    pt = row.get("product_type_id")
    if pt is not None:
        return ("type", int(pt))
    return ("name", (row.get("product_name") or "?").strip().lower())


def list_production_pl_weekly(
    conn: sqlite3.Connection,
    entity_id: Optional[int] = None,
    entity_kind: Optional[str] = None,
    search: Optional[str] = None,
    limit: int = 500,
) -> list[dict]:
    """
    Aggregate profitability_job rows by (entity, product, ISO week of completion).
    Numeric fields are summed; unit_cost is blended total_cost / output_qty.
    """
    jobs = list_production_pl(
        conn, entity_id=entity_id, entity_kind=entity_kind, search=search, limit=10000
    )
    groups: dict[tuple, dict] = {}
    for r in jobs:
        week = _iso_week_key(r.get("end_date_utc") or "")
        gkey = (
            int(r["entity_id"]),
            r["entity_kind"],
            _group_product_key(r),
            week,
        )
        g = groups.get(gkey)
        if not g:
            g = {
                "entity_id": r["entity_id"],
                "entity_kind": r["entity_kind"],
                "owner_name": r["owner_name"],
                "character_name": r["character_name"],
                "product_type_id": r.get("product_type_id"),
                "product_name": r["product_name"],
                "week_key": week,
                "end_date_utc": r.get("end_date_utc") or "",
                "end_date_min": r.get("end_date_utc") or "",
                "end_date_max": r.get("end_date_utc") or "",
                "job_ids": [],
                "job_count": 0,
                "runs": 0,
                "output_qty": 0,
                "job_fee": 0.0,
                "materials_cost": 0.0,
                "invention_cost": 0.0,
                "facility_cost": 0.0,
                "materials_unknown_qty": 0,
                "sold_qty": 0,
                "revenue": 0.0,
                "cogs_for_sold": 0.0,
                "realized_profit": 0.0,
                "unsold_qty": 0,
                "unrealized_value": 0.0,
                "notes": [],
                "view_mode": "weekly",
            }
            groups[gkey] = g
        g["job_ids"].append(int(r["job_id"]))
        g["job_count"] += 1
        g["runs"] += int(r.get("runs") or 0)
        g["output_qty"] += int(r.get("output_qty") or 0)
        g["job_fee"] += float(r.get("job_fee") or 0)
        g["materials_cost"] += float(r.get("materials_cost") or 0)
        g["invention_cost"] += float(r.get("invention_cost") or 0)
        g["facility_cost"] += float(r.get("facility_cost") or 0)
        g["materials_unknown_qty"] += int(r.get("materials_unknown_qty") or 0)
        g["sold_qty"] += int(r.get("sold_qty") or 0)
        g["revenue"] += float(r.get("revenue") or 0)
        g["cogs_for_sold"] += float(r.get("cogs_for_sold") or 0)
        g["realized_profit"] += float(r.get("realized_profit") or 0)
        g["unsold_qty"] += int(r.get("unsold_qty") or 0)
        g["unrealized_value"] += float(r.get("unrealized_value") or 0)
        ed = r.get("end_date_utc") or ""
        if ed:
            if not g["end_date_min"] or ed < g["end_date_min"]:
                g["end_date_min"] = ed
            if not g["end_date_max"] or ed > g["end_date_max"]:
                g["end_date_max"] = ed
                g["end_date_utc"] = ed
        note = (r.get("note") or "").strip()
        if note and note not in g["notes"]:
            g["notes"].append(note)

    out: list[dict] = []
    for g in groups.values():
        total_cost = (
            g["job_fee"] + g["materials_cost"] + g["invention_cost"] + g["facility_cost"]
        )
        oq = int(g["output_qty"] or 0)
        g["total_cost"] = total_cost
        g["unit_cost"] = total_cost / oq if oq > 0 else 0.0
        g["total_estimated_pl"] = (
            float(g["realized_profit"])
            + float(g["unrealized_value"])
            - float(g["unit_cost"]) * int(g["unsold_qty"] or 0)
        )
        g["job_id"] = g["job_ids"][0] if len(g["job_ids"]) == 1 else None
        g["note"] = "; ".join(g["notes"][:3])
        if len(g["notes"]) > 3:
            g["note"] += f" (+{len(g['notes']) - 3} more)"
        out.append(g)

    out.sort(key=lambda x: (x.get("end_date_max") or "", x.get("product_name") or ""), reverse=True)
    return out[: int(limit)]


def _fmt_isk(v: float | None) -> str:
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return "n/a"


def _fmt_dt(iso: str | None) -> str:
    if not iso:
        return "?"
    return str(iso).replace("T", " ").replace("Z", " UTC")[:19]


def _item_name(conn: sqlite3.Connection, type_id: int | None) -> str:
    if not type_id:
        return "?"
    row = conn.execute("SELECT typeName FROM items WHERE typeID = ?", (int(type_id),)).fetchone()
    return row[0] if row else f"type {type_id}"


def _sso_char_name(conn: sqlite3.Connection, character_id: int | None) -> str:
    if not character_id:
        return "?"
    row = conn.execute(
        "SELECT character_name FROM sso_character WHERE character_id = ?",
        (int(character_id),),
    ).fetchone()
    return (row[0] if row and row[0] else None) or f"char {character_id}"


def _wallet_chars_for_entity(
    conn: sqlite3.Connection, entity_id: int, entity_kind: str
) -> list[int]:
    if entity_kind == "character":
        return [int(entity_id)]
    rows = conn.execute(
        "SELECT character_id FROM sso_character WHERE corporation_id = ?",
        (int(entity_id),),
    ).fetchall()
    wallet_chars = [int(r[0]) for r in rows]
    for r in conn.execute(
        """
        SELECT DISTINCT installer_id FROM esi_corporation_industry_jobs
         WHERE corporation_id = ? AND installer_id IS NOT NULL
           AND installer_id IN (SELECT character_id FROM sso_character)
        """,
        (int(entity_id),),
    ).fetchall():
        iid = int(r[0])
        if iid not in wallet_chars:
            wallet_chars.append(iid)
    return wallet_chars


def _describe_lot_source(
    conn: sqlite3.Connection,
    lot: sqlite3.Row | tuple,
    wallet_chars: list[int],
    corp_id: int | None,
) -> list[str]:
    """Human-readable provenance lines for a profitability_lot row."""
    lot_id = int(lot[0])
    type_id = int(lot[2])
    source_kind = lot[3] or ""
    source_id = lot[4]
    opened_at = lot[5]
    unit_cost = lot[8]
    note = lot[9] or ""
    name = _item_name(conn, type_id)
    lines = [f"    Lot #{lot_id}: {name} opened {_fmt_dt(opened_at)}"]
    if unit_cost is not None:
        lines.append(f"      Unit cost in lot: {_fmt_isk(unit_cost)} ISK")
    if note:
        lines.append(f"      Note: {note}")

    if source_kind == "buy_tx" and source_id is not None and wallet_chars:
        ph = ",".join("?" * len(wallet_chars))
        tx = conn.execute(
            f"""
            SELECT character_id, transaction_id, date_utc, quantity, unit_price, is_buy, location_id
              FROM esi_wallet_transactions
             WHERE transaction_id = ? AND character_id IN ({ph})
             LIMIT 1
            """,
            [int(source_id)] + wallet_chars,
        ).fetchone()
        if tx:
            who = _sso_char_name(conn, int(tx[0]))
            side = "bought" if int(tx[5] or 0) else "sold"
            lines.append(
                f"      Source: personal wallet — {who} {side} {int(tx[3]):,} @ "
                f"{_fmt_isk(tx[4])} ISK on {_fmt_dt(tx[2])} (tx {tx[1]}, loc {tx[6]})"
            )
        else:
            lines.append(f"      Source: buy_tx id {source_id} (wallet row not in synced data)")
    elif source_kind == "manufacture" and source_id is not None:
        lines.append(f"      Source: manufactured by industry job {source_id}")
    elif source_kind in ("market_fallback", "unknown"):
        lines.append(f"      Source: {source_kind} (no matching buy lot before consumption)")
    else:
        lines.append(f"      Source: {source_kind} id {source_id}")
    return lines


def _wallet_sale_lines(
    conn: sqlite3.Connection,
    transaction_id: int,
    wallet_chars: list[int],
    corp_id: int | None,
) -> list[str]:
    """Resolve a sale_tx consumer_id to wallet transaction + journal fee lines."""
    lines: list[str] = []
    tx_row = None
    if wallet_chars:
        ph = ",".join("?" * len(wallet_chars))
        tx_row = conn.execute(
            f"""
            SELECT character_id, transaction_id, date_utc, quantity, unit_price,
                   client_id, location_id, journal_ref_id
              FROM esi_wallet_transactions
             WHERE transaction_id = ? AND character_id IN ({ph}) AND is_buy = 0
             LIMIT 1
            """,
            [int(transaction_id)] + wallet_chars,
        ).fetchone()
    if not tx_row and corp_id is not None:
        tx_row = conn.execute(
            """
            SELECT NULL, transaction_id, date_utc, quantity, unit_price,
                   client_id, location_id, journal_ref_id
              FROM esi_corporation_wallet_transactions
             WHERE transaction_id = ? AND corporation_id = ? AND is_buy = 0
             LIMIT 1
            """,
            (int(transaction_id), int(corp_id)),
        ).fetchone()
    if not tx_row:
        lines.append(f"      Sale tx {transaction_id}: not found in synced wallet data")
        return lines

    char_id = int(tx_row[0]) if tx_row[0] is not None else None
    qty = int(tx_row[3] or 0)
    unit_price = float(tx_row[4] or 0)
    gross = unit_price * qty
    who = _sso_char_name(conn, char_id) if char_id else f"corp wallet (corp {corp_id})"
    lines.append(
        f"      Sold by {who}: {qty:,} @ {_fmt_isk(unit_price)} ISK on {_fmt_dt(tx_row[2])} "
        f"(gross {_fmt_isk(gross)} ISK, tx {tx_row[1]}, loc {tx_row[6]})"
    )

    tax_total = 0.0
    tax_lines: list[str] = []
    if char_id is not None:
        jref = tx_row[7]
        jrows = []
        if jref is not None:
            jrows = conn.execute(
                """
                SELECT ref_id, date_utc, ref_type, amount, description
                  FROM esi_wallet_journal
                 WHERE character_id = ? AND ref_id = ?
                ORDER BY date_utc
                """,
                (char_id, int(jref)),
            ).fetchall()
        if not jrows:
            jrows = conn.execute(
                """
                SELECT ref_id, date_utc, ref_type, amount, description
                  FROM esi_wallet_journal
                 WHERE character_id = ?
                   AND context_id = ? AND context_id_type = 'market_transaction_id'
                ORDER BY date_utc
                """,
                (char_id, int(transaction_id)),
            ).fetchall()
        for jr in jrows:
            amt = float(jr[3] or 0)
            if amt >= 0:
                continue
            tax_total += abs(amt)
            tax_lines.append(
                f"        Journal {_fmt_dt(jr[1])} [{jr[2]}]: {_fmt_isk(amt)} ISK — {jr[4] or ''}"
            )
    if tax_lines:
        lines.append(f"      Wallet taxes/fees (from esi_wallet_journal): {_fmt_isk(tax_total)} ISK total")
        lines.extend(tax_lines)
        lines.append(f"      Net after journal fees: {_fmt_isk(gross - tax_total)} ISK (approx.)")
    else:
        try:
            from assumptions import BROKER_FEE, SALES_TAX
            est_tax = gross * (float(SALES_TAX) / 100.0)
            est_broker = gross * (float(BROKER_FEE) / 100.0)
            lines.append(
                f"      No linked journal rows — estimated using assumptions.py: "
                f"sales tax {SALES_TAX}% ≈ {_fmt_isk(est_tax)} ISK, "
                f"broker {BROKER_FEE}% ≈ {_fmt_isk(est_broker)} ISK "
                f"(not stored in wallet transactions API)"
            )
        except Exception:
            lines.append("      No linked journal rows; sync wallet journal for fee breakdown")
    return lines


def _datacore_cost_detail_lines(conn: sqlite3.Connection, datacores: list[tuple[str, int]]) -> list[str]:
    """Per-datacore ISK lines using prices table (sell_min, else buy_max)."""
    out: list[str] = []
    if not datacores:
        out.append("    (no datacores configured)")
        return out
    for name, qty in datacores:
        row = conn.execute("SELECT typeID FROM items WHERE typeName = ?", (name,)).fetchone()
        if not row:
            out.append(f"    {name}: {qty} — type not in DB, cost unknown")
            continue
        type_id = int(row[0])
        prow = conn.execute(
            "SELECT sell_min, buy_max FROM prices WHERE typeID = ?", (type_id,)
        ).fetchone()
        if not prow:
            out.append(f"    {name}: {qty} — no market price in DB")
            continue
        sell_min = float(prow[0] or 0)
        buy_max = float(prow[1] or 0)
        unit = sell_min if sell_min > 0 else buy_max
        src = "sell_min" if sell_min > 0 else "buy_max"
        if unit <= 0:
            out.append(f"    {name}: {qty} — price zero/missing")
            continue
        line = unit * int(qty)
        out.append(
            f"    {name}: {qty:,} x {_fmt_isk(unit)} ISK ({src}) = {_fmt_isk(line)} ISK"
        )
    return out


def _invention_breakdown_lines(
    conn: sqlite3.Connection,
    blueprint_type_id: int,
    product_name: str,
    mfg_job_runs: int,
    invention_cost_applied: float,
    facility_cost_applied: float,
) -> list[str]:
    """Detailed invention / research cost explanation for T2 manufacturing jobs."""
    lines: list[str] = ["", "--- Invention cost (T2 research) ---"]
    if not _is_t2_blueprint(conn, blueprint_type_id):
        lines.append("  Not a T2 blueprint — invention cost does not apply.")
        return lines

    _ensure_blueprint_datacore_bindings(conn)
    t1_row = conn.execute(
        """
        SELECT i.typeName, ir.t1_blueprint_type_id
          FROM invention_recipes ir
          JOIN items i ON i.typeID = ir.t1_blueprint_type_id
         WHERE ir.t2_blueprint_type_id = ?
         LIMIT 1
        """,
        (int(blueprint_type_id),),
    ).fetchone()

    bind = conn.execute(
        """
        SELECT dc1_name, dc1_qty, dc2_name, dc2_qty,
               base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs,
               production_cost_per_run,
               research_time_days, research_time_hours, research_time_minutes
          FROM blueprint_datacore_bindings
         WHERE blueprint_type_id = ?
        """,
        (int(blueprint_type_id),),
    ).fetchone()

    lines.append(f"  T2 blueprint: {product_name} (type {blueprint_type_id})")
    if t1_row:
        lines.append(f"  T1 copy for invention: {t1_row[0]} (type {t1_row[1]})")

    if not bind:
        lines.extend([
            "  No row in blueprint_datacore_bindings for this blueprint.",
            "  Set datacores, invention chance, and cost in the Decryptor comparison tab,",
            "  then Rebuild ledger — invention cost is 0 until configured.",
        ])
        return lines

    dc1, dq1, dc2, dq2 = bind[0], bind[1], bind[2], bind[3]
    base_chance_pct = float(bind[4]) if bind[4] is not None else 40.0
    inv_base = float(bind[5]) if bind[5] is not None else 0.0
    base_bpc_runs = int(bind[6]) if bind[6] is not None else 10
    if base_bpc_runs not in (1, 10):
        base_bpc_runs = 10
    prod_per_run = (
        float(bind[7]) if len(bind) > 7 and bind[7] is not None and float(bind[7]) > 0 else None
    )
    r_days = float(bind[8] or 0) if len(bind) > 8 else 0.0
    r_hrs = float(bind[9] or 0) if len(bind) > 9 else 0.0
    r_min = float(bind[10] or 0) if len(bind) > 10 else 0.0

    datacores: list[tuple[str, int]] = []
    if dc1 and (dq1 or 0) > 0:
        datacores.append((dc1, int(dq1)))
    if dc2 and (dq2 or 0) > 0:
        datacores.append((dc2, int(dq2)))

    lines.extend([
        "  Saved parameters (Decryptor tab / blueprint_datacore_bindings):",
        f"    Base invention chance: {base_chance_pct:.1f}%",
        f"    Invention cost per attempt (excl. datacores & decryptor): {_fmt_isk(inv_base)} ISK",
        "      (typically T1 BPC copy + other fixed ISK you enter per attempt)",
        f"    Base BPC runs (no decryptor): {base_bpc_runs}",
    ])
    if r_days or r_hrs or r_min:
        lines.append(
            f"    Research time (1 run): {int(r_days)}d {int(r_hrs)}h {int(r_min)}m "
            "(planning only; not added to ISK cost here)"
        )
    lines.append("  Datacores consumed per invention attempt (current DB prices):")
    lines.extend(_datacore_cost_detail_lines(conn, datacores))

    results = compare_decryptor_profitability(
        blueprint_name_or_product=product_name,
        base_invention_chance_pct=base_chance_pct,
        invention_cost_without_decryptor=inv_base,
        base_bpc_runs=base_bpc_runs,
        datacores=datacores if datacores else None,
        conn=conn,
    )
    valid = [x for x in results if not x.get("error")]
    if not valid:
        lines.append("  Could not run decryptor comparison (see errors in Decryptor tab).")
        return lines

    best = max(valid, key=lambda x: x.get("profit_per_bpc") or -1e99)
    dc_cost = float(best.get("datacore_cost") or 0)
    dec_price = float(best.get("decryptor_price") or 0)
    attempt_cost = float(best.get("attempt_cost") or (inv_base + dc_cost + dec_price))
    success_pct = float(best.get("success_prob_pct") or 0)
    expected_bpc = float(best.get("expected_inv_cost") or 0)
    bpc_runs = max(1, int(best.get("bpc_runs") or base_bpc_runs))
    bpc_me = float(best.get("bpc_me") or 0)
    dec_name = (best.get("decryptor_name") or "No decryptor").strip()

    lines.extend([
        "  Selected decryptor (highest profit_per_bpc, same rule as shopping list refresh):",
        f"    {dec_name}",
        f"    Decryptor market price: {_fmt_isk(dec_price)} ISK",
        f"    Effective success chance: {success_pct:.1f}%",
        "",
        "  Per invention attempt:",
        f"    {_fmt_isk(inv_base)} ISK  invention base (binding)",
        f"  + {_fmt_isk(dc_cost)} ISK  datacores",
        f"  + {_fmt_isk(dec_price)} ISK  decryptor",
        f"  = {_fmt_isk(attempt_cost)} ISK  total attempt cost",
        "",
        "  Expected cost per successful T2 BPC:",
        f"    {_fmt_isk(attempt_cost)} / {success_pct:.1f}% = {_fmt_isk(expected_bpc)} ISK",
        f"    Resulting BPC: {bpc_runs} manufacturing run(s) @ ME {bpc_me:.0f}%",
        "",
        "  Allocated to this manufacturing job:",
        f"    Job manufacturing runs: {mfg_job_runs:,}",
        f"    BPC runs from invention: {bpc_runs:,}",
    ])
    bpc_fraction = float(mfg_job_runs) / float(bpc_runs) if bpc_runs else 0.0
    calc_inv = expected_bpc * bpc_fraction
    lines.append(f"    BPC fraction used: {mfg_job_runs}/{bpc_runs} = {bpc_fraction:.4f}")
    lines.append(
        f"    Invention on job = {_fmt_isk(expected_bpc)} x {bpc_fraction:.4f} "
        f"= {_fmt_isk(calc_inv)} ISK"
    )
    lines.append(f"    Stored in ledger: {_fmt_isk(invention_cost_applied)} ISK")
    if abs(calc_inv - invention_cost_applied) > 1.0:
        lines.append("    (Rebuild ledger if stored value differs — bindings/prices may have changed.)")
    if prod_per_run is not None:
        calc_fac = prod_per_run * mfg_job_runs
        lines.extend([
            "",
            "  Facility / production cost (optional binding):",
            f"    {_fmt_isk(prod_per_run)} ISK per manufacturing run x {mfg_job_runs:,} runs "
            f"= {_fmt_isk(calc_fac)} ISK",
            f"    Stored in ledger: {_fmt_isk(facility_cost_applied)} ISK",
        ])
    lines.append(
        "  Source: decryptor_profitability.compare_decryptor_profitability + "
        "blueprint_datacore_bindings (not from ESI invention jobs)."
    )
    return lines


def get_job_pl_breakdown(
    conn: sqlite3.Connection,
    entity_id: int,
    entity_kind: str,
    job_id: int,
) -> str:
    """
    Build a detailed text breakdown for one profitability_job row:
    costs, material FIFO sources, sales, and journal taxes where available.
    """
    ensure_profitability_tables(conn)
    wallet_chars = _wallet_chars_for_entity(conn, entity_id, entity_kind)
    corp_id = int(entity_id) if entity_kind == "corporation" else None

    pj = conn.execute(
        """
        SELECT job_id, activity_id, blueprint_type_id, product_type_id, product_name,
               runs, output_qty, start_date_utc, end_date_utc, status,
               job_fee, materials_cost, invention_cost, facility_cost, materials_unknown_qty,
               unit_cost, sold_qty, revenue, cogs_for_sold, realized_profit,
               unsold_qty, unrealized_value, note
          FROM profitability_job
         WHERE character_id = ? AND entity_kind = ? AND job_id = ?
        """,
        (int(entity_id), entity_kind, int(job_id)),
    ).fetchone()
    if not pj:
        return f"No profitability_job row for {entity_kind} {entity_id} job {job_id}. Rebuild ledger first."

    product_name = pj[4] or _item_name(conn, pj[3])
    lines: list[str] = [
        f"=== Job P/L breakdown: {product_name} ===",
        f"Entity: {entity_kind} {entity_id}  |  Job ID: {job_id}",
        f"Completed: {_fmt_dt(pj[8])}  |  Runs: {pj[5]}  |  Output qty: {pj[6]:,}",
        "",
        "--- Cost basis (per unit and total) ---",
        f"  Job fee (ESI industry job 'cost'):     {_fmt_isk(pj[10])} ISK",
        f"  Materials (FIFO lots consumed):      {_fmt_isk(pj[11])} ISK",
        f"  Invention (expected, T2 bindings):   {_fmt_isk(pj[12])} ISK",
        f"  Facility / prod cost per run:        {_fmt_isk(pj[13])} ISK",
        f"  Total cost:                          {_fmt_isk(float(pj[10] or 0) + float(pj[11] or 0) + float(pj[12] or 0) + float(pj[13] or 0))} ISK",
        f"  Unit cost (/ output {pj[6]:,}):      {_fmt_isk(pj[15])} ISK",
    ]
    if pj[14]:
        lines.append(f"  ⚠ Materials with unknown basis: {pj[14]:,} units")

    bp_tid = pj[2]
    if bp_tid and _is_t2_blueprint(conn, int(bp_tid)):
        lines.extend(
            _invention_breakdown_lines(
                conn,
                int(bp_tid),
                product_name,
                int(pj[5] or 0),
                float(pj[12] or 0),
                float(pj[13] or 0),
            )
        )
    elif pj[22]:
        lines.append(f"  Invention note: {pj[22]}")

    # ESI job metadata
    if entity_kind == "character":
        ej = conn.execute(
            """
            SELECT installer_id, facility_id, blueprint_type_id, cost, completed_date_utc
              FROM esi_industry_jobs WHERE character_id = ? AND job_id = ?
            """,
            (int(entity_id), int(job_id)),
        ).fetchone()
    else:
        ej = conn.execute(
            """
            SELECT installer_id, facility_id, blueprint_type_id, cost, completed_date_utc
              FROM esi_corporation_industry_jobs WHERE corporation_id = ? AND job_id = ?
            """,
            (int(entity_id), int(job_id)),
        ).fetchone()
    if ej:
        lines.extend([
            "",
            "--- ESI industry job ---",
            f"  Installer: {_sso_char_name(conn, ej[0])} ({ej[0]})",
            f"  Facility ID: {ej[1]}  |  Blueprint type: {ej[2]}",
            f"  ESI cost field: {_fmt_isk(ej[3])} ISK  |  Completed: {_fmt_dt(ej[4])}",
        ])

    # Output lot
    out_lot = conn.execute(
        """
        SELECT lot_id, character_id, type_id, source_kind, source_id, opened_at,
               qty_initial, qty_remaining, unit_cost, note
          FROM profitability_lot
         WHERE character_id = ? AND entity_kind = ? AND source_kind = 'manufacture'
           AND source_id = ?
        """,
        (int(entity_id), entity_kind, int(job_id)),
    ).fetchone()

    lines.extend(["", "--- Output lot (manufactured) ---"])
    if out_lot:
        lines.extend(_describe_lot_source(conn, out_lot, wallet_chars, corp_id))
        lines.append(
            f"    Produced {int(out_lot[6]):,} units; {int(out_lot[7]):,} still in lot (unsold)"
        )
    else:
        lines.append("    (no manufacture lot found)")

    # Material inputs
    lines.extend(["", "--- Materials consumed (FIFO) ---"])
    mat_rows = conn.execute(
        """
        SELECT pc.qty, pc.unit_cost, pc.consumed_at, pl.lot_id, pl.type_id, pl.source_kind,
               pl.source_id, pl.opened_at, pl.unit_cost, pl.note
          FROM profitability_consumption pc
          JOIN profitability_lot pl ON pl.lot_id = pc.lot_id
         WHERE pc.consumer_kind = 'manufacture_input' AND pc.consumer_id = ?
         ORDER BY pl.type_id, pc.consumed_at
        """,
        (int(job_id),),
    ).fetchall()
    if not mat_rows:
        lines.append("    (none recorded)")
    else:
        by_type: dict[int, list] = {}
        for mr in mat_rows:
            by_type.setdefault(int(mr[4]), []).append(mr)
        for tid, mlist in sorted(by_type.items(), key=lambda x: _item_name(conn, x[0])):
            tqty = sum(int(m[0]) for m in mlist)
            tcost = sum(float(m[1] or 0) * int(m[0]) for m in mlist if m[1] is not None)
            lines.append(f"  {_item_name(conn, tid)}: {tqty:,} units, cost {_fmt_isk(tcost)} ISK")
            seen_lots: set[int] = set()
            for m in mlist:
                lid = int(m[3])
                if lid in seen_lots:
                    continue
                seen_lots.add(lid)
                lot_tuple = (lid, m[3], m[4], m[5], m[6], m[7], None, None, m[8], m[9])
                # Re-fetch full lot row for describe
                full = conn.execute(
                    """
                    SELECT lot_id, character_id, type_id, source_kind, source_id, opened_at,
                           qty_initial, qty_remaining, unit_cost, note
                      FROM profitability_lot WHERE lot_id = ?
                    """,
                    (lid,),
                ).fetchone()
                if full:
                    lines.extend(_describe_lot_source(conn, full, wallet_chars, corp_id))
                lines.append(
                    f"      Used {int(m[0]):,} @ {_fmt_isk(m[1])}/unit on {_fmt_dt(m[2])}"
                )

    # Sales from output lot
    lines.extend(["", "--- Sales matched to this job (FIFO) ---"])
    if out_lot:
        lot_id = int(out_lot[0])
        sale_rows = conn.execute(
            """
            SELECT pc.qty, pc.unit_cost, pc.consumed_at, pc.consumer_id
              FROM profitability_consumption pc
             WHERE pc.lot_id = ? AND pc.consumer_kind = 'sale_tx'
             ORDER BY pc.consumed_at
            """,
            (lot_id,),
        ).fetchall()
        if not sale_rows:
            lines.append("    (no sales consumed from this lot yet)")
        for sr in sale_rows:
            q, uc, cat, txid = int(sr[0]), sr[1], sr[2], sr[3]
            rev_line = ""
            if txid is not None:
                up = None
                if wallet_chars:
                    ph = ",".join("?" * len(wallet_chars))
                    row = conn.execute(
                        f"""
                        SELECT unit_price FROM esi_wallet_transactions
                         WHERE transaction_id = ? AND character_id IN ({ph}) AND is_buy = 0
                         LIMIT 1
                        """,
                        [int(txid)] + wallet_chars,
                    ).fetchone()
                    if row and row[0]:
                        up = float(row[0])
                if up is None and corp_id is not None:
                    row = conn.execute(
                        """
                        SELECT unit_price FROM esi_corporation_wallet_transactions
                         WHERE transaction_id = ? AND corporation_id = ? AND is_buy = 0
                         LIMIT 1
                        """,
                        (int(txid), int(corp_id)),
                    ).fetchone()
                    if row and row[0]:
                        up = float(row[0])
                if up is not None:
                    rev_line = f" — revenue {_fmt_isk(up * q)} ISK @ {_fmt_isk(up)}/unit"
            lines.append(
                f"  {q:,} units, COGS unit {_fmt_isk(uc)}{rev_line} on {_fmt_dt(cat)}"
            )
            if txid:
                lines.extend(_wallet_sale_lines(conn, int(txid), wallet_chars, corp_id))
    else:
        lines.append("    (cannot list sales without output lot)")

    lines.extend([
        "",
        "--- P/L summary (from profitability_job) ---",
        f"  Sold qty: {pj[16]:,}  |  Revenue (wallet unit_price × qty): {_fmt_isk(pj[17])} ISK",
        f"  COGS for sold: {_fmt_isk(pj[18])} ISK  |  Realized P/L: {_fmt_isk(pj[19])} ISK",
        f"  Unsold: {pj[20]:,}  |  Unrealized (@ current sell_min): {_fmt_isk(pj[21])} ISK",
        f"  Est. total P/L: {_fmt_isk(float(pj[19] or 0) + float(pj[21] or 0) - float(pj[15] or 0) * int(pj[20] or 0))} ISK",
        "",
        "Data sources: esi_wallet_transactions, esi_corporation_wallet_transactions,",
        "esi_wallet_journal, esi_*_industry_jobs, profitability_lot/consumption,",
        "blueprint_datacore_bindings (invention), profitability_price_snapshot (fallback costs).",
    ])
    return "\n".join(lines)


def get_group_pl_breakdown(
    conn: sqlite3.Connection,
    entity_id: int,
    entity_kind: str,
    job_ids: list[int],
    product_name: str,
    week_key: str,
) -> str:
    """Combined breakdown for a weekly group (same product, same ISO week, multiple jobs)."""
    if not job_ids:
        return "No jobs in this group."

    job_ids = sorted({int(j) for j in job_ids})
    ph = ",".join("?" * len(job_ids))
    rows = conn.execute(
        f"""
        SELECT job_id, runs, output_qty, end_date_utc, job_fee, materials_cost,
               invention_cost, facility_cost, materials_unknown_qty, unit_cost,
               sold_qty, revenue, cogs_for_sold, realized_profit, unsold_qty,
               unrealized_value, blueprint_type_id, product_type_id
          FROM profitability_job
         WHERE character_id = ? AND entity_kind = ? AND job_id IN ({ph})
         ORDER BY end_date_utc, job_id
        """,
        [int(entity_id), entity_kind] + job_ids,
    ).fetchall()

    if not rows:
        return (
            f"No profitability_job rows for {len(job_ids)} job id(s). Rebuild ledger first."
        )

    lines: list[str] = [
        f"=== Weekly group: {product_name} ===",
        f"Week: {week_key}  |  Entity: {entity_kind} {entity_id}  |  Jobs: {len(rows)}",
        "",
        "--- Group totals (sum of jobs) ---",
    ]
    tot = {
        "runs": 0,
        "output_qty": 0,
        "job_fee": 0.0,
        "materials_cost": 0.0,
        "invention_cost": 0.0,
        "facility_cost": 0.0,
        "materials_unknown_qty": 0,
        "sold_qty": 0,
        "revenue": 0.0,
        "cogs_for_sold": 0.0,
        "realized_profit": 0.0,
        "unsold_qty": 0,
        "unrealized_value": 0.0,
    }
    bp_tid = None
    ptid = None
    for r in rows:
        tot["runs"] += int(r[1] or 0)
        tot["output_qty"] += int(r[2] or 0)
        tot["job_fee"] += float(r[4] or 0)
        tot["materials_cost"] += float(r[5] or 0)
        tot["invention_cost"] += float(r[6] or 0)
        tot["facility_cost"] += float(r[7] or 0)
        tot["materials_unknown_qty"] += int(r[8] or 0)
        tot["sold_qty"] += int(r[10] or 0)
        tot["revenue"] += float(r[11] or 0)
        tot["cogs_for_sold"] += float(r[12] or 0)
        tot["realized_profit"] += float(r[13] or 0)
        tot["unsold_qty"] += int(r[14] or 0)
        tot["unrealized_value"] += float(r[15] or 0)
        if r[16] and bp_tid is None:
            bp_tid = int(r[16])
        if r[17] and ptid is None:
            ptid = int(r[17])

    total_cost = (
        tot["job_fee"] + tot["materials_cost"] + tot["invention_cost"] + tot["facility_cost"]
    )
    blended_unit = total_cost / tot["output_qty"] if tot["output_qty"] else 0.0
    total_pl = (
        tot["realized_profit"]
        + tot["unrealized_value"]
        - blended_unit * tot["unsold_qty"]
    )
    lines.extend([
        f"  Runs (sum): {tot['runs']:,}  |  Output qty: {tot['output_qty']:,}",
        f"  Job fees: {_fmt_isk(tot['job_fee'])}  |  Materials: {_fmt_isk(tot['materials_cost'])}",
        f"  Invention: {_fmt_isk(tot['invention_cost'])}  |  Facility: {_fmt_isk(tot['facility_cost'])}",
        f"  Total cost: {_fmt_isk(total_cost)}  |  Blended unit cost: {_fmt_isk(blended_unit)}",
        f"  Sold: {tot['sold_qty']:,}  |  Revenue: {_fmt_isk(tot['revenue'])}",
        f"  Realized P/L: {_fmt_isk(tot['realized_profit'])}  |  Unrealized: {_fmt_isk(tot['unrealized_value'])}",
        f"  Est. total P/L: {_fmt_isk(total_pl)}",
    ])
    if tot["materials_unknown_qty"]:
        lines.append(f"  ⚠ Unknown-basis material units (sum): {tot['materials_unknown_qty']:,}")

    if bp_tid and _is_t2_blueprint(conn, bp_tid):
        lines.extend(
            _invention_breakdown_lines(
                conn,
                bp_tid,
                product_name,
                tot["runs"],
                tot["invention_cost"],
                tot["facility_cost"],
            )
        )
        lines.append(
            "  (Invention section is per successful BPC recipe; totals above sum each job's allocation.)"
        )

    lines.extend(["", "--- Per-job summary ---"])
    for r in rows:
        jid = int(r[0])
        oq = int(r[2] or 0)
        uc = float(r[9] or 0)
        unsold = int(r[14] or 0)
        unv = float(r[15] or 0)
        real = float(r[13] or 0)
        job_pl = real + unv - uc * unsold
        lines.append(
            f"  Job {jid}  completed {_fmt_dt(r[3])}  |  runs {int(r[1] or 0):,}  out {oq:,}  |  "
            f"cost {_fmt_isk(float(r[4] or 0) + float(r[5] or 0) + float(r[6] or 0) + float(r[7] or 0))}  |  "
            f"rev {_fmt_isk(r[11])}  realized {_fmt_isk(real)}  est.P/L {_fmt_isk(job_pl)}"
        )

    lines.extend([
        "",
        "=" * 72,
        "DETAILED BREAKDOWN PER JOB (same as single-row double-click)",
        "=" * 72,
    ])
    for i, jid in enumerate(job_ids):
        if i > 0:
            lines.append("\n" + "#" * 72 + "\n")
        lines.append(get_job_pl_breakdown(conn, entity_id, entity_kind, jid))

    return "\n".join(lines)
