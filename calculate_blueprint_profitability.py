"""
Calculate manufacturing profitability for a single blueprint.

Given a blueprint (or product name), computes:
- Total input cost (materials × prices)
- Output revenue (product qty × sell price with transaction costs)
- System cost: system_cost_percent × EIV (Estimated Item Value). EIV uses CCP's
  adjusted_price for the output product (from ESI /markets/prices/), not market sell
  price — e.g. Leshak EIV ~258m vs sell ~494m; Triglavian/some items have different EIV.
- Profit and return %
"""

import math
import sqlite3
from pathlib import Path

import requests

from calculate_reprocessing_value import (
    sell_into_buy_order,
    sell_order_with_fees,
)
from assumptions import SALES_TAX

# ESI market prices (adjusted_price = CCP's value used for EIV, industry fees, etc.)
ESI_MARKETS_PRICES = "https://esi.evetech.net/latest/markets/prices/"
_adjusted_price_cache = {}

DATABASE_FILE = "eve_manufacturing.db"


def resolve_blueprint(conn, name):
    """
    Resolve blueprint or product name to blueprint row.
    Returns dict with blueprintTypeID, productTypeID, productName, outputQuantity, or None if not found.
    - Try as product name (blueprints.productName)
    - Try as blueprint type name (items.typeName = name where typeID in blueprints)
    - Try name without ' Blueprint' as product name
    """
    name = (name or "").strip()
    if not name:
        return None
    cur = conn.execute(
        "SELECT blueprintTypeID, productTypeID, productName, outputQuantity FROM blueprints WHERE productName = ?",
        (name,),
    )
    row = cur.fetchone()
    if row:
        return {"blueprintTypeID": row[0], "productTypeID": row[1], "productName": row[2], "outputQuantity": row[3]}
    cur = conn.execute(
        """SELECT b.blueprintTypeID, b.productTypeID, b.productName, b.outputQuantity
           FROM blueprints b JOIN items i ON b.blueprintTypeID = i.typeID WHERE i.typeName = ?""",
        (name,),
    )
    row = cur.fetchone()
    if row:
        return {"blueprintTypeID": row[0], "productTypeID": row[1], "productName": row[2], "outputQuantity": row[3]}
    if name.endswith(" Blueprint"):
        product_name = name[:-len(" Blueprint")].strip()
        cur = conn.execute(
            "SELECT blueprintTypeID, productTypeID, productName, outputQuantity FROM blueprints WHERE productName = ?",
            (product_name,),
        )
        row = cur.fetchone()
        if row:
            return {"blueprintTypeID": row[0], "productTypeID": row[1], "productName": row[2], "outputQuantity": row[3]}
    return None


def get_blueprint_materials(conn, blueprint_type_id):
    """Return list of (materialTypeID, materialName, quantity) for the blueprint."""
    cur = conn.execute(
        "SELECT materialTypeID, materialName, quantity FROM manufacturing_materials WHERE blueprintTypeID = ? ORDER BY materialName",
        (blueprint_type_id,),
    )
    return [{"materialTypeID": r[0], "materialName": r[1], "quantity": int(r[2])} for r in cur.fetchall()]


def _material_unit_price_raw(price_row, input_price_type):
    """Return raw unit price for one material (no tax/fees; taxes apply on sell side only)."""
    buy_max = float(price_row["buy_max"]) if price_row.get("buy_max") else 0.0
    sell_min = float(price_row["sell_min"]) if price_row.get("sell_min") else 0.0
    if input_price_type == "buy_offer":
        return buy_max if buy_max else 0.0
    return sell_min if sell_min else 0.0


def _output_price_after_costs(price_row, output_price_type):
    """Return unit price after costs for selling output (product)."""
    buy_max = float(price_row["buy_max"]) if price_row.get("buy_max") else 0.0
    sell_min = float(price_row["sell_min"]) if price_row.get("sell_min") else 0.0
    if output_price_type == "sell_immediate":
        return sell_into_buy_order(buy_max) if buy_max else 0.0
    return sell_order_with_fees(sell_min) if sell_min else 0.0


def get_adjusted_price(type_id: int) -> float | None:
    """
    Return CCP adjusted_price for type_id from ESI /markets/prices/ (used for EIV).
    Cached in-memory for the session. Returns None if not found or on error.
    """
    if type_id in _adjusted_price_cache:
        return _adjusted_price_cache[type_id]
    try:
        r = requests.get(ESI_MARKETS_PRICES, timeout=15)
        r.raise_for_status()
        data = r.json()
        for row in data:
            if row.get("type_id") == type_id:
                adj = float(row.get("adjusted_price") or 0)
                _adjusted_price_cache[type_id] = adj
                return adj
        _adjusted_price_cache[type_id] = None
        return None
    except Exception:
        _adjusted_price_cache[type_id] = None
        return None


def calculate_blueprint_profitability(
    blueprint_name_or_product=None,
    input_price_type="buy_immediate",
    output_price_type="sell_immediate",
    system_cost_percent=0.0,
    material_efficiency=0,
    number_of_runs=1,
    region_id=None,
    manufacturing_tax_rate=None,
    db_file=DATABASE_FILE,
):
    """
    Calculate manufacturing profitability for blueprint run(s).

    material_efficiency: ME level 0–10 (each level 4% reduction in material qty). Default 0.
    number_of_runs: number of runs (multiplies input and output quantities). Default 1.
    region_id: kept for compatibility; currently ignored for cost (only materials + system cost are used).
    manufacturing_tax_rate: deprecated and ignored; manufacturing tax is no longer applied.

    Returns dict with total_input_cost, system_cost, output_revenue, profit, etc.
    """
    if not Path(db_file).exists():
        return {"error": f"Database not found: {db_file}"}
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    try:
        bp = resolve_blueprint(conn, blueprint_name_or_product)
        if not bp:
            return {"error": f"Blueprint or product not found: {blueprint_name_or_product!r}"}
        blueprint_type_id = bp["blueprintTypeID"]
        product_type_id = bp["productTypeID"]
        product_name = bp["productName"]
        output_quantity = int(bp["outputQuantity"])
        materials = get_blueprint_materials(conn, blueprint_type_id)
        if not materials:
            return {"error": f"No manufacturing materials found for blueprint (product: {product_name})"}

        me_level = max(0, min(10, float(material_efficiency)))  # 0–10, each level 4% reduction
        runs = max(1, int(number_of_runs))
        me_fraction = me_level/100  # material efficiency as fraction (e.g. 40% → 0.4)
        # Per-run amount after ME: base*(1-ME). Total used = max(runs, ceil(base*(1-ME)*runs)); per item (per run) = total/runs (min 1 when base>=1).
        material_type_ids = [m["materialTypeID"] for m in materials]
        placeholders = ",".join("?" * len(material_type_ids))
        cur = conn.execute(
            f"SELECT typeID, buy_max, sell_min FROM prices WHERE typeID IN ({placeholders})",
            material_type_ids,
        )
        price_by_type = {int(row["typeID"]): dict(row) for row in cur.fetchall()}

        input_materials_out = []
        total_input_cost = 0.0
        materials_priced_at_zero = []
        for m in materials:
            tid = m["materialTypeID"]
            base_qty = m["quantity"]
            total_qty = max(runs, math.ceil(base_qty * (1.0 - me_fraction) * runs))
            per_run = total_qty / runs
            pr = price_by_type.get(tid)
            if not pr:
                unit_price = 0.0
                materials_priced_at_zero.append(m["materialName"])
            else:
                unit_price = _material_unit_price_raw(pr, input_price_type)
                if unit_price <= 0.0:
                    materials_priced_at_zero.append(m["materialName"])
            total_cost = unit_price * total_qty
            total_input_cost += total_cost
            input_materials_out.append({
                "materialName": m["materialName"],
                "base_quantity": base_qty,
                "quantity": total_qty,
                "quantity_per_run": per_run,
                "unit_price": unit_price,
                "total_cost": total_cost,
            })

        cur = conn.execute("SELECT buy_max, sell_min FROM prices WHERE typeID = ?", (product_type_id,))
        out_price_row = cur.fetchone()
        if out_price_row:
            out_price_row = dict(out_price_row)
        else:
            out_price_row = {}
        # EIV (Estimated Item Value): CCP adjusted_price for the output product × output qty × runs.
        # If no adjusted_price from ESI, fall back to market price (sell_min or buy_max).
        # System cost (job/facility fee) = system_cost_percent × EIV (not raw material cost).
        adjusted = get_adjusted_price(product_type_id)
        sell_min_out = float(out_price_row.get("sell_min") or 0)
        buy_max_out = float(out_price_row.get("buy_max") or 0)
        eiv_price_per_unit = adjusted if adjusted is not None and adjusted > 0 else (sell_min_out or buy_max_out)
        eiv = eiv_price_per_unit * output_quantity * runs
        system_cost_percent_fraction = float(system_cost_percent if system_cost_percent is not None else 0) / 100.0
        system_cost_isk = system_cost_percent_fraction * eiv

        # Manufacturing tax is no longer applied; only material cost + system cost (EIV-based) are used.
        manufacturing_tax_total_all_runs = 0.0
        tax_details = []

        output_unit_price = _output_price_after_costs(out_price_row, output_price_type)
        output_total_qty = output_quantity * runs
        output_revenue = output_unit_price * output_total_qty

        total_cost = total_input_cost + system_cost_isk + manufacturing_tax_total_all_runs
        profit = output_revenue - total_cost
        return_percent = (profit / total_cost * 100.0) if total_cost > 0 else 0.0

        items_produced = output_total_qty
        cost_per_item = total_cost / items_produced if items_produced else 0.0
        revenue_per_item = output_revenue / items_produced if items_produced else 0.0
        profit_per_item = profit / items_produced if items_produced else 0.0

        return {
            "blueprintTypeID": blueprint_type_id,
            "productTypeID": product_type_id,
            "productName": product_name,
            "outputQuantity": output_quantity,
            "number_of_runs": runs,
            "material_efficiency": me_level,
            "input_materials": input_materials_out,
            "materials_priced_at_zero": materials_priced_at_zero,
            "total_input_cost": total_input_cost,
            "eiv": eiv,
            "adjusted_price": adjusted,
            "eiv_price_per_unit": eiv_price_per_unit,
            "eiv_source": "adjusted_price" if (adjusted is not None and adjusted > 0) else "market",
            "system_cost": system_cost_isk,
            "system_cost_percent": float(system_cost_percent or 0),
            "manufacturing_tax": 0.0,
            "manufacturing_tax_rate": 0.0,
            "tax_details": tax_details,
            "output_unit_price": output_unit_price,
            "output_total_quantity": output_total_qty,
            "output_revenue": output_revenue,
            "profit": profit,
            "return_percent": return_percent,
            "items_produced": items_produced,
            "cost_per_item": cost_per_item,
            "revenue_per_item": revenue_per_item,
            "profit_per_item": profit_per_item,
            "input_price_type": input_price_type,
            "output_price_type": output_price_type,
        }
    finally:
        conn.close()


if __name__ == "__main__":
    # Test: Nanite Repair Paste, 10% ME, system cost 21.77%, The Forge, 10 runs (for line-by-line debugging)
    THE_FORGE_REGION_ID = 10000002
    result = calculate_blueprint_profitability(
        blueprint_name_or_product="Nanite Repair Paste",
        input_price_type="buy_immediate",
        output_price_type="sell_immediate",
        system_cost_percent=21.77,
        material_efficiency=10,   # 10% material efficiency -> me_fraction = 0.10
        number_of_runs=10,
        region_id=THE_FORGE_REGION_ID,
        db_file=DATABASE_FILE,
    )
    print("=== Test: Nanite Repair Paste, 10% ME, system 21.77%, The Forge, 10 runs ===\n")
    if "error" in result:
        print("ERROR:", result["error"])
    else:
        for key in (
            "productName", "outputQuantity", "number_of_runs", "material_efficiency",
            "total_input_cost", "system_cost", "system_cost_percent", "manufacturing_tax", "manufacturing_tax_rate",
            "output_unit_price", "output_total_quantity", "output_revenue",
            "profit", "return_percent",
            "items_produced", "cost_per_item", "revenue_per_item", "profit_per_item",
            "materials_priced_at_zero",
        ):
            val = result.get(key)
            if key == "input_materials":
                continue
            print(f"  {key}: {val}")
        print("\n  input_materials:")
        for m in result.get("input_materials", []):
            bq = m.get("base_quantity", "?")
            print(f"    {m['materialName']}: base_qty={bq} total_qty={m['quantity']} per_run={m['quantity_per_run']} unit_price={m['unit_price']} total_cost={m['total_cost']}")
        if result.get("tax_details"):
            print("  tax_details (per run):")
            for t in result["tax_details"]:
                print(f"    {t['materialName']}: avg={t['average']} qty={t['quantity']} tax={t['tax']}")
