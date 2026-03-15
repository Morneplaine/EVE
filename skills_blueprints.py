"""
Skills and blueprints: list unique manufacturing skills, filter blueprints by user skills,
run profitability and return top N by profit (ISK) and by return %.
T1 blueprints use 10%% ME, T2 use 0%% ME.
"""

import sqlite3
from pathlib import Path

from calculate_blueprint_profitability import calculate_blueprint_profitability

DATABASE_FILE = "eve_manufacturing.db"


def get_unique_skills(conn):
    """
    Return list of dicts with skillID, skillName from manufacturing_skills.
    Returns [] if table does not exist or is empty.
    """
    try:
        cur = conn.execute(
            "SELECT DISTINCT skillID, skillName FROM manufacturing_skills ORDER BY skillName"
        )
        return [{"skillID": row[0], "skillName": row[1]} for row in cur.fetchall()]
    except sqlite3.OperationalError:
        return []


def get_blueprint_requirements(conn):
    """
    Return dict: blueprintTypeID -> list of (skillID, required_level).
    """
    cur = conn.execute(
        "SELECT blueprintTypeID, skillID, level FROM manufacturing_skills"
    )
    reqs = {}
    for row in cur.fetchall():
        bid, sid, lvl = row[0], row[1], int(row[2])
        reqs.setdefault(bid, []).append((sid, lvl))
    return reqs


def is_t2_blueprint(conn, blueprint_type_id):
    """True if this blueprint is a T2 (invention output)."""
    row = conn.execute(
        "SELECT 1 FROM invention_recipes WHERE t2_blueprint_type_id = ? LIMIT 1",
        (blueprint_type_id,),
    ).fetchone()
    return row is not None


def get_available_blueprint_ids(conn, user_skill_levels):
    """
    user_skill_levels: dict skillID -> int (0-5, user's level).
    Return list of blueprintTypeID for which the user meets all required skills.
    """
    reqs = get_blueprint_requirements(conn)
    cur = conn.execute("SELECT blueprintTypeID FROM blueprints")
    all_bp_ids = [row[0] for row in cur.fetchall()]
    available = []
    for bid in all_bp_ids:
        needed = reqs.get(bid, [])
        if not needed:
            available.append(bid)
            continue
        if all(user_skill_levels.get(sid, 0) >= lvl for sid, lvl in needed):
            available.append(bid)
    return available


def run_profitability_analysis(
    db_file,
    blueprint_type_ids,
    input_price_type="buy_immediate",
    output_price_type="sell_immediate",
    system_cost_percent=8.61,
    progress_callback=None,
):
    """
    Run single-blueprint profitability for each blueprint. T1 = 10% ME, T2 = 0% ME.
    progress_callback(current, total) is called after each blueprint (from worker thread).
    Returns list of dicts: productName, blueprintTypeID, profit, return_percent, is_t2, ...
    """
    if not Path(db_file).exists():
        return []
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    results = []
    total = len(blueprint_type_ids)
    try:
        t2_set = set()
        cur = conn.execute(
            "SELECT t2_blueprint_type_id FROM invention_recipes"
        )
        for row in cur.fetchall():
            t2_set.add(row[0])
        cur = conn.execute(
            "SELECT blueprintTypeID, productName FROM blueprints"
        )
        bid_to_name = {row[0]: row[1] for row in cur.fetchall()}
        for tested, bid in enumerate(blueprint_type_ids, 1):
            product_name = bid_to_name.get(bid)
            if not product_name:
                if progress_callback:
                    progress_callback(tested, total)
                continue
            me = 10 if bid not in t2_set else 0
            r = calculate_blueprint_profitability(
                blueprint_name_or_product=product_name,
                input_price_type=input_price_type,
                output_price_type=output_price_type,
                system_cost_percent=system_cost_percent,
                material_efficiency=me,
                number_of_runs=1,
                db_file=db_file,
            )
            if r.get("error"):
                if progress_callback:
                    progress_callback(tested, total)
                continue
            results.append({
                "productName": product_name,
                "blueprintTypeID": bid,
                "profit": r.get("profit") or 0.0,
                "return_percent": r.get("return_percent") or 0.0,
                "is_t2": bid in t2_set,
                "material_efficiency": me,
                "total_input_cost": r.get("total_input_cost") or 0.0,
                "output_revenue": r.get("output_revenue") or 0.0,
            })
            if progress_callback:
                progress_callback(tested, total)
    finally:
        conn.close()
    return results


def top_n_by_profit(results, n=20):
    """Sort by profit descending, return first n."""
    return sorted(results, key=lambda x: x["profit"], reverse=True)[:n]


def top_n_by_return(results, n=20):
    """Sort by return_percent descending, return first n."""
    return sorted(results, key=lambda x: x["return_percent"], reverse=True)[:n]
