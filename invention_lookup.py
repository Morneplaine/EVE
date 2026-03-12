"""
Resolve T1 blueprint/product to possible T2 invention outputs using SDE invention data.
"""

import sqlite3
from pathlib import Path

from calculate_blueprint_profitability import resolve_blueprint

DATABASE_FILE = "eve_manufacturing.db"


def get_t2_products_from_t1(t1_blueprint_or_product_name, db_file=DATABASE_FILE):
    """
    Given a T1 blueprint or T1 product name, return the list of T2 products that can be invented from it.

    Returns list of dicts: t2_product_name, t2_blueprint_type_id, probability, quantity.
    Returns empty list if T1 not found or no invention data; returns [] on error.
    """
    if not Path(db_file).exists():
        return []
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    try:
        cur = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='invention_recipes'"
        )
        if not cur.fetchone():
            return []
        bp = resolve_blueprint(conn, t1_blueprint_or_product_name)
        if not bp:
            return []
        t1_blueprint_type_id = bp["blueprintTypeID"]
        cur = conn.execute("""
            SELECT r.t2_blueprint_type_id, r.quantity, r.probability, b.productName
            FROM invention_recipes r
            JOIN blueprints b ON b.blueprintTypeID = r.t2_blueprint_type_id
            WHERE r.t1_blueprint_type_id = ?
            ORDER BY b.productName
        """, (t1_blueprint_type_id,))
        rows = cur.fetchall()
        out = []
        for row in rows:
            out.append({
                "t2_product_name": row["productName"],
                "t2_blueprint_type_id": row["t2_blueprint_type_id"],
                "probability": row["probability"],
                "quantity": row["quantity"],
            })
        return out
    except Exception:
        return []
    finally:
        conn.close()
