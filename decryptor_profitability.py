"""
Compare profitability of T2 invention using different decryptors.

Flow: T1 BPC copy → Invention (optional decryptor) → T2 BPC (ME/TE/runs) → Manufacturing → T2 product.
Expected cost per successful BPC = (invention attempt cost including decryptor and datacores) / success probability.
Profit per BPC = manufacturing profit from that BPC's runs at its ME − expected invention cost.
"""

import sqlite3
from pathlib import Path

from calculate_blueprint_profitability import calculate_blueprint_profitability
from decryptors_data import (
    DECRYPTORS,
    BASE_ME_PCT,
    BASE_RUNS_MODULE,
    BASE_RUNS_SHIP,
)


DATABASE_FILE = "eve_manufacturing.db"

# Common datacores used in invention (static names, must match items.typeName).
# Full list of EVE datacores, alphabetically ordered.
DATACORE_NAMES = [
    "Datacore - Amarrian Starship Engineering",
    "Datacore - Caldari Starship Engineering",
    "Datacore - Core Subsystems Engineering",
    "Datacore - Defensive Subsystems Engineering",
    "Datacore - Electromagnetic Physics",
    "Datacore - Electronic Engineering",
    "Datacore - Gallentean Starship Engineering",
    "Datacore - Graviton Physics",
    "Datacore - High Energy Physics",
    "Datacore - Hydromagnetic Physics",
    "Datacore - Laser Physics",
    "Datacore - Mechanical Engineering",
    "Datacore - Minmatar Starship Engineering",
    "Datacore - Molecular Engineering",
    "Datacore - Nanite Engineering",
    "Datacore - Nuclear Physics",
    "Datacore - Offensive Subsystems Engineering",
    "Datacore - Plasma Physics",
    "Datacore - Propulsion Subsystems Engineering",
    "Datacore - Quantum Physics",
    "Datacore - Rocket Science",
    "Datacore - Triglavian Quantum Engineering",
    "Datacore - Upwell Starship Engineering",
]


def get_decryptor_price(conn, type_id, use_sell_min=True):
    """Return market price for decryptor (sell_min or buy_max) or 0 if missing."""
    cur = conn.execute(
        "SELECT sell_min, buy_max FROM prices WHERE typeID = ?",
        (type_id,),
    )
    row = cur.fetchone()
    if not row:
        return 0.0
    sell_min = float(row[0] or 0)
    buy_max = float(row[1] or 0)
    return sell_min if use_sell_min and sell_min > 0 else buy_max


def _estimate_datacore_cost_per_attempt(conn, datacores) -> float:
    """
    Estimate ISK cost per attempt from datacore names and quantities using market prices.
    Uses sell_min if available, otherwise buy_max.
    """
    total = 0.0
    if not datacores:
        return 0.0
    for name, qty in datacores:
        cur = conn.execute("SELECT typeID FROM items WHERE typeName = ?", (name,))
        row = cur.fetchone()
        if not row:
            continue
        type_id = int(row[0])
        cur2 = conn.execute("SELECT sell_min, buy_max FROM prices WHERE typeID = ?", (type_id,))
        prow = cur2.fetchone()
        if not prow:
            continue
        sell_min = float(prow[0] or 0)
        buy_max = float(prow[1] or 0)
        unit_price = sell_min or buy_max
        if unit_price <= 0:
            continue
        total += unit_price * qty
    return total


def compare_decryptor_profitability(
    blueprint_name_or_product,
    base_invention_chance_pct,
    invention_cost_without_decryptor,
    base_bpc_runs=10,
    input_price_type="buy_immediate",
    output_price_type="sell_immediate",
    system_cost_percent=8.61,
    region_id=None,
    db_file=DATABASE_FILE,
    datacores=None,
):
    """
    Compare profit per successful BPC for no-decryptor and each decryptor.

    base_invention_chance_pct: e.g. 40 for 40% base success (before decryptor multiplier).
    invention_cost_without_decryptor: ISK per attempt (T1 BPC copy + other fixed costs, excluding decryptor and datacores).
    base_bpc_runs: 10 for modules/ammo, 1 for ships/rigs (invention output runs when no decryptor).
    datacores: optional iterable of (name, qty) pairs for datacores per attempt.

    Returns list of dicts: decryptor_name, success_prob_pct, expected_inv_cost, bpc_me, bpc_runs,
    decryptor_price, manufacturing_profit, profit_per_bpc, error (if any).
    """
    if not Path(db_file).exists():
        return [{"error": f"Database not found: {db_file}"}]

    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    base_chance = max(0.01, min(1.0, float(base_invention_chance_pct) / 100.0))
    inv_cost_no_dec = max(0.0, float(invention_cost_without_decryptor))
    datacore_cost = _estimate_datacore_cost_per_attempt(conn, datacores)
    base_runs = int(base_bpc_runs) if base_bpc_runs in (1, 10) else 10

    out = []

    # No decryptor
    success_prob = base_chance * 1.0
    attempt_cost = inv_cost_no_dec + datacore_cost
    expected_inv = attempt_cost / success_prob if success_prob > 0 else 0.0
    me = BASE_ME_PCT
    runs = base_runs
    mfg = calculate_blueprint_profitability(
        blueprint_name_or_product=blueprint_name_or_product,
        input_price_type=input_price_type,
        output_price_type=output_price_type,
        system_cost_percent=system_cost_percent,
        material_efficiency=me,
        number_of_runs=runs,
        region_id=region_id,
        db_file=db_file,
    )
    if "error" in mfg:
        conn.close()
        return [{"decryptor_name": "No decryptor", "error": mfg["error"]}]
    profit_bpc = mfg["profit"] - expected_inv
    out.append({
        "decryptor_name": "No decryptor",
        "decryptor_type_id": None,
        "success_prob_pct": success_prob * 100,
        "attempt_cost": attempt_cost,
        "inv_cost_no_dec": inv_cost_no_dec,
        "datacore_cost": datacore_cost,
        "expected_inv_cost": expected_inv,
        "decryptor_price": 0.0,
        "bpc_me": me,
        "bpc_runs": runs,
        "manufacturing_profit": mfg["profit"],
        "profit_per_bpc": profit_bpc,
    })

    # Each decryptor
    for name, type_id, prob_mult, run_mod, me_mod, te_mod in DECRYPTORS:
        decryptor_price = get_decryptor_price(conn, type_id)
        attempt_cost = inv_cost_no_dec + datacore_cost + decryptor_price
        success_prob = base_chance * prob_mult
        if success_prob <= 0:
            success_prob = 0.01
        expected_inv = attempt_cost / success_prob
        me = max(0, min(10, BASE_ME_PCT + me_mod))
        runs = max(1, base_runs + run_mod)

        mfg = calculate_blueprint_profitability(
            blueprint_name_or_product=blueprint_name_or_product,
            input_price_type=input_price_type,
            output_price_type=output_price_type,
            system_cost_percent=system_cost_percent,
            material_efficiency=me,
            number_of_runs=runs,
            region_id=region_id,
            db_file=db_file,
        )
        if "error" in mfg:
            out.append({
                "decryptor_name": name,
                "decryptor_type_id": type_id,
                "error": mfg["error"],
            })
            continue
        profit_bpc = mfg["profit"] - expected_inv
        out.append({
            "decryptor_name": name,
            "decryptor_type_id": type_id,
            "success_prob_pct": success_prob * 100,
            "attempt_cost": attempt_cost,
            "inv_cost_no_dec": inv_cost_no_dec,
            "datacore_cost": datacore_cost,
            "expected_inv_cost": expected_inv,
            "decryptor_price": decryptor_price,
            "bpc_me": me,
            "bpc_runs": runs,
            "manufacturing_profit": mfg["profit"],
            "profit_per_bpc": profit_bpc,
        })

    conn.close()
    return out
