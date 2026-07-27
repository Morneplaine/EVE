"""
Jita -> nullsec arbitrage finder (default target: C-N4OD, Fountain).

Identifies high-volume items in Jita and compares the Jita sell price (plus a
landed/freight cost) against the cheapest sell order in a target nullsec
structure market, to surface resell opportunities.

Landed price model (per unit):
    landed = jita_sell * 1.1 + 1000 ISK * packaged_volume_m3
Opportunity margin:
    margin = (target_sell - landed) / jita_sell

Target structure sell prices come from the authenticated ESI endpoint
    GET /markets/structures/{structure_id}/   (scope: esi-markets.structure_markets.v1)
so it works for private alliance markets as long as one linked character has
docking access. Structures are discovered from ranged buy orders that appear in
the public region order book, then probed for read access with each token.

Usage:
    python arbitrage_finder.py --discover        # find structures + access only
    python arbitrage_finder.py                   # full run (top 500 by Jita ISK volume)
    python arbitrage_finder.py --top 300 --rank units --min-margin 0.1
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import sqlite3
import sys
import time
from pathlib import Path

import requests

from eve_sso_sync import get_valid_access_token

ESI = "https://esi.evetech.net/latest"
DB_FILE = "eve_manufacturing.db"
CREDENTIALS_FILE = Path(__file__).resolve().parent / "eve_sso_credentials.json"
STRUCTURE_CACHE = Path(__file__).resolve().parent / "arbitrage_structures.json"

# Defaults: Jita / The Forge  ->  C-N4OD / Fountain
JITA_REGION_ID = 10000002
TARGET_SYSTEM_ID = 30004600   # C-N4OD
TARGET_REGION_ID = 10000058   # Fountain
TARGET_NAME = "C-N4OD"

# Landed cost model
COLLATERAL_MULTIPLIER = 1.1   # jita_sell * 1.1
FREIGHT_ISK_PER_M3 = 1000.0   # + 1000 ISK per m3


def load_credentials() -> tuple[str, str]:
    cid = os.environ.get("EVE_SSO_CLIENT_ID", "").strip()
    sec = os.environ.get("EVE_SSO_CLIENT_SECRET", "").strip()
    if (not cid or not sec) and CREDENTIALS_FILE.exists():
        data = json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))
        cid = cid or str(data.get("client_id", "")).strip()
        sec = sec or str(data.get("client_secret", "")).strip()
    return cid, sec


def esi_get(url: str, *, params=None, token=None, timeout=30):
    headers = {"Authorization": f"Bearer {token}"} if token else {}
    return requests.get(url, params=params or {}, headers=headers, timeout=timeout)


# ---------------------------------------------------------------------------
# Structure discovery + access
# ---------------------------------------------------------------------------

def discover_target_structures(region_id: int, system_id: int) -> dict[int, int]:
    """
    Find structure IDs in the target system from ranged buy orders that DO appear
    in the public region order book. Returns {location_id: order_count}.
    """
    print(f"Sweeping region {region_id} buy orders to locate structures in system {system_id}...")
    locs: dict[int, int] = {}
    page = 1
    pages = 1
    while page <= pages:
        r = esi_get(
            f"{ESI}/markets/{region_id}/orders/",
            params={"datasource": "tranquility", "order_type": "buy", "page": page},
        )
        if r.status_code == 404:
            break
        r.raise_for_status()
        data = r.json()
        if not data:
            break
        pages = int(r.headers.get("X-Pages", page))
        for o in data:
            if o.get("system_id") == system_id:
                lid = o["location_id"]
                # NPC stations are < 100,000,000,000; Upwell structures are huge IDs
                if lid > 70_000_000_000:
                    locs[lid] = locs.get(lid, 0) + 1
        if page % 10 == 0 or page == pages:
            print(f"  page {page}/{pages} (structures so far: {len(locs)})")
        page += 1
    return locs


def find_access(conn, structure_ids: list[int], cid: str, sec: str) -> dict[int, dict]:
    """
    For each structure, find a character whose token can read its market.
    Returns {structure_id: {"character_id":, "character_name":, "name":, "pages":}}.
    """
    chars = conn.execute(
        "SELECT character_id, character_name FROM sso_character"
    ).fetchall()
    access: dict[int, dict] = {}
    tokens: dict[int, str] = {}

    for sid in structure_ids:
        for char_id, name in chars:
            tok = tokens.get(char_id)
            if tok is None:
                tok = get_valid_access_token(conn, char_id, cid, sec) or ""
                tokens[char_id] = tok
            if not tok:
                continue
            r = esi_get(
                f"{ESI}/markets/structures/{sid}/",
                params={"datasource": "tranquility", "page": 1},
                token=tok,
            )
            if r.status_code == 200:
                sname = ""
                sr = esi_get(
                    f"{ESI}/universe/structures/{sid}/",
                    params={"datasource": "tranquility"}, token=tok,
                )
                if sr.status_code == 200:
                    sname = sr.json().get("name", "")
                access[sid] = {
                    "character_id": char_id,
                    "character_name": name,
                    "name": sname,
                    "pages": int(r.headers.get("X-Pages", 1)),
                }
                print(f"  ACCESS structure {sid} via {name}: {sname or '(name hidden)'} ({access[sid]['pages']} pages)")
                break
        else:
            print(f"  no access to structure {sid}")
    return access


def fetch_structure_sell_min(conn, structure_id: int, char_id: int, cid: str, sec: str) -> dict[int, float]:
    """Fetch all orders in a structure; return {type_id: min_sell_price}."""
    sell_min: dict[int, float] = {}
    page = 1
    pages = 1
    while page <= pages:
        tok = get_valid_access_token(conn, char_id, cid, sec)
        if not tok:
            break
        r = esi_get(
            f"{ESI}/markets/structures/{structure_id}/",
            params={"datasource": "tranquility", "page": page}, token=tok,
        )
        if r.status_code != 200:
            break
        pages = int(r.headers.get("X-Pages", page))
        for o in r.json():
            if o.get("is_buy_order"):
                continue
            tid = o["type_id"]
            p = float(o["price"])
            if tid not in sell_min or p < sell_min[tid]:
                sell_min[tid] = p
        page += 1
        time.sleep(0.1)
    return sell_min


# ---------------------------------------------------------------------------
# Jita high-volume items
# ---------------------------------------------------------------------------

def top_high_volume_items(conn, region_id: int, top: int, days: int, rank: str) -> list[dict]:
    """
    Rank items by average daily Jita activity over the last `days` of history.
    rank='isk'  -> average daily ISK traded (volume * average price)
    rank='units'-> average daily units traded

    Aggregate market_history_daily alone first (uses the PK index, fast), then
    join item metadata in Python. Joining items inside the aggregation query
    triggers a pathological SQLite plan on this large table.
    """
    metric = "AVG(volume * average)" if rank == "isk" else "AVG(volume)"
    # Over-fetch so the packaged_volume>0 filter still leaves ~top rows
    agg = conn.execute(
        f"""
        SELECT type_id,
               AVG(volume)          AS avg_units,
               AVG(volume*average)  AS avg_isk,
               {metric}             AS rank_metric
          FROM market_history_daily
         WHERE region_id = ?
           AND date_utc >= date('now', ?)
         GROUP BY type_id
         HAVING avg_units > 0
         ORDER BY rank_metric DESC
         LIMIT ?
        """,
        (region_id, f"-{days} days", int(top * 1.5) + 50),
    ).fetchall()

    type_ids = [r[0] for r in agg]
    meta: dict[int, tuple] = {}
    for i in range(0, len(type_ids), 500):
        chunk = type_ids[i:i + 500]
        ph = ",".join("?" * len(chunk))
        for tid, name, vol in conn.execute(
            f"SELECT typeID, typeName, packaged_volume FROM items WHERE typeID IN ({ph})", chunk
        ):
            meta[tid] = (name, vol)

    out = []
    for r in agg:
        tid = r[0]
        name, vol = meta.get(tid, (None, None))
        if not name or not vol or vol <= 0:
            continue
        out.append({
            "type_id": tid,
            "name": name,
            "volume_m3": vol,
            "avg_units": r[1],
            "avg_isk": r[2],
        })
        if len(out) >= top:
            break
    return out


def jita_sell_prices(conn, type_ids: list[int]) -> dict[int, float]:
    """Jita sell_min from the prices table (Fuzzwork-populated)."""
    out: dict[int, float] = {}
    for i in range(0, len(type_ids), 500):
        chunk = type_ids[i:i + 500]
        ph = ",".join("?" * len(chunk))
        for tid, sell in conn.execute(
            f"SELECT typeID, sell_min FROM prices WHERE typeID IN ({ph})", chunk
        ):
            out[tid] = float(sell or 0)
    return out


# ---------------------------------------------------------------------------
# Reusable: C-N4OD sell prices (for the GUI Paste & Compare tab)
# ---------------------------------------------------------------------------

def load_cn4od_sell_prices(conn, cid: str, sec: str, progress=None) -> tuple[dict[int, float], str]:
    """Return ({type_id: min_sell in C-N4OD}, info_note).

    Reuses the cached structure access (arbitrage_structures.json); discovers +
    probes access if the cache is empty. Intended to be called from the GUI so
    Paste & Compare can show C-N4OD prices without re-running the full CLI.
    """
    def say(msg):
        if progress:
            progress(msg)

    access: dict[int, dict] = {}
    if STRUCTURE_CACHE.exists():
        try:
            cached = json.loads(STRUCTURE_CACHE.read_text(encoding="utf-8"))
            access = {int(k): v for k, v in cached.get("access", {}).items()}
        except Exception:
            access = {}
    if access:
        say(f"Using {len(access)} cached C-N4OD structure(s).")
    else:
        say("Discovering C-N4OD structures (first run may take ~30s)...")
        candidates = discover_target_structures(TARGET_REGION_ID, TARGET_SYSTEM_ID)
        if not candidates:
            return {}, "No C-N4OD structures discovered (no ranged buy orders visible)."
        access = find_access(conn, list(candidates.keys()), cid, sec)
        if access:
            STRUCTURE_CACHE.write_text(
                json.dumps({"system": TARGET_NAME, "access": {str(k): v for k, v in access.items()}}, indent=2),
                encoding="utf-8",
            )
    if not access:
        return {}, f"No linked character has market access to any {TARGET_NAME} structure."

    target_sell: dict[int, float] = {}
    for sid, info in access.items():
        say(f"Fetching C-N4OD market {sid} ({info.get('name') or 'hidden'})...")
        sm = fetch_structure_sell_min(conn, sid, info["character_id"], cid, sec)
        for tid, p in sm.items():
            if tid not in target_sell or p < target_sell[tid]:
                target_sell[tid] = p
    return target_sell, f"C-N4OD sell orders for {len(target_sell)} item types."


def landed_from_jita(jita_sell: float, volume_m3: float) -> float:
    """Landed cost in C-N4OD for one unit given Jita sell price and volume."""
    return jita_sell * COLLATERAL_MULTIPLIER + FREIGHT_ISK_PER_M3 * (volume_m3 or 0.0)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description="Jita -> C-N4OD arbitrage finder")
    ap.add_argument("--discover", action="store_true", help="Only discover structures + access, then exit")
    ap.add_argument("--top", type=int, default=500, help="Number of high-volume items (default 500)")
    ap.add_argument("--days", type=int, default=365, help="History window for volume ranking (default 365; Jita history may be stale)")
    ap.add_argument("--rank", choices=["isk", "units"], default="isk", help="Rank high-volume by ISK or units")
    ap.add_argument("--min-margin", type=float, default=0.0, help="Only output opportunities with margin >= this")
    ap.add_argument("--out", default="arbitrage_cn4od.csv", help="Output CSV path")
    ap.add_argument("--refresh-structures", action="store_true", help="Ignore cached structures and rediscover")
    ap.add_argument("--add-structure", type=int, default=None, help="Manually add a known structure ID to probe")
    args = ap.parse_args()

    cid, sec = load_credentials()
    if not cid or not sec:
        print("No SSO credentials (eve_sso_credentials.json or EVE_SSO_CLIENT_ID/SECRET).")
        return 1

    conn = sqlite3.connect(DB_FILE, timeout=60)
    conn.execute("PRAGMA busy_timeout=60000")

    # 1) Structures + access (cached)
    access: dict[int, dict] = {}
    if STRUCTURE_CACHE.exists() and not args.refresh_structures:
        cached = json.loads(STRUCTURE_CACHE.read_text(encoding="utf-8"))
        access = {int(k): v for k, v in cached.get("access", {}).items()}
        print(f"Loaded {len(access)} accessible structure(s) from cache.")

    if not access:
        candidates: dict[int, int] = {}
        if args.add_structure:
            candidates[args.add_structure] = 1
        candidates.update(discover_target_structures(TARGET_REGION_ID, TARGET_SYSTEM_ID))
        print(f"Found {len(candidates)} candidate structure(s) in {TARGET_NAME}.")
        if not candidates:
            print("No structures discovered (no ranged buy orders visible). "
                  "If you know the structure ID, pass --add-structure <id>.")
            return 1
        access = find_access(conn, list(candidates.keys()), cid, sec)
        STRUCTURE_CACHE.write_text(
            json.dumps({"system": TARGET_NAME, "access": {str(k): v for k, v in access.items()}},
                       indent=2),
            encoding="utf-8",
        )

    if not access:
        print(f"No linked character has docking/market access to any {TARGET_NAME} structure.")
        return 1

    if args.discover:
        print("\nAccessible structures:")
        for sid, info in access.items():
            print(f"  {sid}  {info.get('name') or '(hidden)'}  via {info['character_name']}")
        return 0

    # 2) Target sell prices (min across all accessible structures)
    target_sell: dict[int, float] = {}
    for sid, info in access.items():
        print(f"Fetching market for structure {sid} ({info.get('name') or 'hidden'})...")
        sm = fetch_structure_sell_min(conn, sid, info["character_id"], cid, sec)
        for tid, p in sm.items():
            if tid not in target_sell or p < target_sell[tid]:
                target_sell[tid] = p
    print(f"{TARGET_NAME} has sell orders for {len(target_sell)} item types.")

    # 3) High-volume Jita items
    items = top_high_volume_items(conn, JITA_REGION_ID, args.top, args.days, args.rank)
    print(f"Selected top {len(items)} Jita items by {args.rank} volume (last {args.days} days).")

    jita = jita_sell_prices(conn, [it["type_id"] for it in items])

    # 4) Compute opportunities
    rows = []
    for it in items:
        tid = it["type_id"]
        js = jita.get(tid, 0.0)
        cn = target_sell.get(tid, 0.0)
        vol = it["volume_m3"] or 0.0
        if js <= 0 or cn <= 0:
            continue
        landed = js * COLLATERAL_MULTIPLIER + FREIGHT_ISK_PER_M3 * vol
        margin = (cn - landed) / js
        rows.append({
            "type_id": tid,
            "name": it["name"],
            "jita_sell": round(js, 2),
            "cn4od_sell": round(cn, 2),
            "volume_m3": vol,
            "landed_price": round(landed, 2),
            "margin": round(margin, 4),
            "profit_per_unit": round(cn - landed, 2),
            "jita_avg_units_day": round(it["avg_units"], 1),
            "jita_avg_isk_day": round(it["avg_isk"], 0),
        })

    rows = [r for r in rows if r["margin"] >= args.min_margin]
    rows.sort(key=lambda r: r["margin"], reverse=True)

    # 5) Output
    fields = ["type_id", "name", "jita_sell", "cn4od_sell", "landed_price",
              "profit_per_unit", "margin", "volume_m3",
              "jita_avg_units_day", "jita_avg_isk_day"]
    with open(args.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    print(f"\nWrote {len(rows)} opportunities to {args.out}")
    print(f"\nTop 20 by margin (margin >= {args.min_margin}):")
    print(f"{'Item':<34}{'Jita':>12}{'C-N4OD':>12}{'Landed':>12}{'Margin':>9}")
    for r in rows[:20]:
        print(f"{r['name'][:33]:<34}{r['jita_sell']:>12,.0f}{r['cn4od_sell']:>12,.0f}"
              f"{r['landed_price']:>12,.0f}{r['margin']*100:>8.1f}%")

    conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
