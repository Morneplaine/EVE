"""
Skill-queue attribute planner.

For each SSO character, read the training queue and, for every queued skill,
add its SP to the running total of its PRIMARY attribute and SP/2 to its
SECONDARY attribute. The result is a per-attribute SP total that indicates the
best neural remap for the queued training.

SP per queue entry:
  --sp-mode remaining  (default) = level_end_sp - training_start_sp  (SP still to train)
  --sp-mode level                = level_end_sp - level_start_sp     (full level SP)

Usage:
  python skill_queue_attributes.py
  python skill_queue_attributes.py --character "Salvage Firn"
  python skill_queue_attributes.py --sp-mode level --json out.json
"""
from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
from pathlib import Path

from eve_sso_sync import (
    fetch_skill_queue,
    fetch_type_attributes,
    get_valid_access_token,
)

DB_FILE = "eve_manufacturing.db"
CREDENTIALS_FILE = Path(__file__).resolve().parent / "eve_sso_credentials.json"
ATTR_CACHE_FILE = Path(__file__).resolve().parent / "skill_attribute_cache.json"

PRIMARY_ATTR_DOGMA = 180
SECONDARY_ATTR_DOGMA = 181

# Character attribute type IDs -> name
ATTR_NAMES = {
    167: "Perception",
    166: "Memory",
    168: "Willpower",
    165: "Intelligence",
    164: "Charisma",
}
# Display / output order
ATTR_ORDER = ["Perception", "Memory", "Willpower", "Intelligence", "Charisma"]


def load_credentials() -> tuple[str, str]:
    cid = os.environ.get("EVE_SSO_CLIENT_ID", "").strip()
    sec = os.environ.get("EVE_SSO_CLIENT_SECRET", "").strip()
    if (not cid or not sec) and CREDENTIALS_FILE.exists():
        data = json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))
        cid = cid or str(data.get("client_id", "")).strip()
        sec = sec or str(data.get("client_secret", "")).strip()
    return cid, sec


def _load_attr_cache() -> dict:
    if ATTR_CACHE_FILE.exists():
        try:
            return json.loads(ATTR_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_attr_cache(cache: dict) -> None:
    try:
        ATTR_CACHE_FILE.write_text(json.dumps(cache, indent=0), encoding="utf-8")
    except Exception:
        pass


def get_skill_attributes(type_id: int, cache: dict) -> tuple[str | None, str | None, str]:
    """Return (primary_name, secondary_name, skill_name). Cached (attributes never change)."""
    key = str(type_id)
    if key in cache:
        c = cache[key]
        return c.get("primary"), c.get("secondary"), c.get("name", f"Type {type_id}")
    info = fetch_type_attributes(type_id)
    attrs = {a["attribute_id"]: a["value"] for a in info.get("dogma_attributes", [])}
    prim = ATTR_NAMES.get(int(attrs.get(PRIMARY_ATTR_DOGMA, 0)))
    sec = ATTR_NAMES.get(int(attrs.get(SECONDARY_ATTR_DOGMA, 0)))
    name = info.get("name", f"Type {type_id}")
    cache[key] = {"primary": prim, "secondary": sec, "name": name}
    return prim, sec, name


def entry_sp(entry: dict, sp_mode: str) -> int:
    end = entry.get("level_end_sp") or 0
    if sp_mode == "level":
        start = entry.get("level_start_sp") or 0
    else:  # remaining
        start = entry.get("training_start_sp")
        if start is None:
            start = entry.get("level_start_sp") or 0
    return max(0, int(end) - int(start))


def compute_attribute_totals(
    conn: sqlite3.Connection,
    cid: str,
    sec: str,
    sp_mode: str = "remaining",
    character_ids: list[int] | None = None,
) -> dict:
    """
    Returns:
      {
        "sp_mode": ...,
        "per_char": {name: {attr: sp, ...}, ...},
        "totals": {attr: sp, ...},
        "details": {name: [ {skill, sp, primary, secondary, level}, ... ]},
        "errors": {name: "msg"},
      }
    """
    cache = _load_attr_cache()
    rows = conn.execute(
        "SELECT character_id, character_name FROM sso_character ORDER BY character_name COLLATE NOCASE"
    ).fetchall()
    if character_ids:
        rows = [r for r in rows if r[0] in set(character_ids)]

    per_char: dict[str, dict[str, float]] = {}
    details: dict[str, list] = {}
    errors: dict[str, str] = {}
    totals = {a: 0.0 for a in ATTR_ORDER}

    for char_id, name in rows:
        name = name or str(char_id)
        token = get_valid_access_token(conn, char_id, cid, sec)
        if not token:
            errors[name] = "no valid token (re-login in EVE SSO Sync)"
            continue
        try:
            queue = fetch_skill_queue(char_id, token)
        except Exception as e:
            errors[name] = f"skill queue fetch failed: {e}"
            continue

        ctotals = {a: 0.0 for a in ATTR_ORDER}
        clist = []
        for entry in queue:
            skill_id = entry.get("skill_id")
            if skill_id is None:
                continue
            sp = entry_sp(entry, sp_mode)
            if sp <= 0:
                continue
            prim, secn, sname = get_skill_attributes(skill_id, cache)
            if prim:
                ctotals[prim] += sp
            if secn:
                ctotals[secn] += sp / 2.0
            clist.append({
                "skill": sname,
                "level": entry.get("finished_level"),
                "sp": sp,
                "primary": prim,
                "secondary": secn,
            })
        per_char[name] = ctotals
        details[name] = clist
        for a in ATTR_ORDER:
            totals[a] += ctotals[a]

    _save_attr_cache(cache)
    return {
        "sp_mode": sp_mode,
        "per_char": per_char,
        "totals": totals,
        "details": details,
        "errors": errors,
    }


def _fmt(n: float) -> str:
    return f"{n:,.0f}"


def main() -> int:
    ap = argparse.ArgumentParser(description="Skill-queue attribute (remap) planner")
    ap.add_argument("--character", help="Only this character (name or id)")
    ap.add_argument("--sp-mode", choices=["remaining", "level"], default="remaining",
                    help="remaining = SP left to train (default); level = full level SP")
    ap.add_argument("--details", action="store_true", help="Print per-skill breakdown")
    ap.add_argument("--json", metavar="PATH", help="Write result JSON to PATH")
    args = ap.parse_args()

    cid, sec = load_credentials()
    if not cid or not sec:
        print("No SSO credentials (eve_sso_credentials.json or EVE_SSO_CLIENT_ID/SECRET).")
        return 1

    conn = sqlite3.connect(DB_FILE, timeout=60)
    conn.execute("PRAGMA busy_timeout=60000")

    character_ids = None
    if args.character:
        if args.character.isdigit():
            character_ids = [int(args.character)]
        else:
            row = conn.execute(
                "SELECT character_id FROM sso_character WHERE character_name = ? COLLATE NOCASE",
                (args.character,),
            ).fetchone()
            if not row:
                print(f"Character not found: {args.character}")
                return 1
            character_ids = [row[0]]

    result = compute_attribute_totals(conn, cid, sec, args.sp_mode, character_ids)
    conn.close()

    hdr = f"{'Character':<22}" + "".join(f"{a:>14}" for a in ATTR_ORDER)
    print(f"\nSP mode: {result['sp_mode']}  (primary += SP, secondary += SP/2)\n")
    print(hdr)
    print("-" * len(hdr))
    for name, ct in result["per_char"].items():
        print(f"{name[:21]:<22}" + "".join(f"{_fmt(ct[a]):>14}" for a in ATTR_ORDER))
    print("-" * len(hdr))
    tot = result["totals"]
    print(f"{'TOTAL':<22}" + "".join(f"{_fmt(tot[a]):>14}" for a in ATTR_ORDER))

    if tot:
        ranked = sorted(ATTR_ORDER, key=lambda a: tot[a], reverse=True)
        print(f"\nSuggested remap priority: {' > '.join(ranked)}")

    if args.details:
        for name, clist in result["details"].items():
            print(f"\n=== {name} ===")
            for d in clist:
                print(f"  {d['skill']} L{d['level']}: {_fmt(d['sp'])} SP "
                      f"[P:{d['primary']} / S:{d['secondary']}]")

    if result["errors"]:
        print("\nErrors:")
        for name, msg in result["errors"].items():
            print(f"  {name}: {msg}")

    if args.json:
        Path(args.json).write_text(json.dumps(result, indent=2), encoding="utf-8")
        print(f"\nWrote JSON to {args.json}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
