"""
Push eve_manufacturing.db to git, keeping only the last 4 versions.

Rotates backups: current db -> v1, v1 -> v2, v2 -> v3 (v3 overwritten).
Then commits and pushes the 4 files. Run this when you want to sync the DB to git.
"""

import shutil
import subprocess
import sys
from pathlib import Path

DB_NAME = "eve_manufacturing.db"
V1 = "eve_manufacturing_v1.db"
V2 = "eve_manufacturing_v2.db"
V3 = "eve_manufacturing_v3.db"


def main():
    root = Path(__file__).resolve().parent
    db = root / DB_NAME
    v1 = root / V1
    v2 = root / V2
    v3 = root / V3

    if not db.exists():
        print(f"Database not found: {db}")
        sys.exit(1)

    # Rotate: v2 -> v3, v1 -> v2, db -> v1 (so we keep 4 versions; oldest v3 is overwritten)
    if v2.exists():
        shutil.copy2(v2, v3)
    if v1.exists():
        shutil.copy2(v1, v2)
    shutil.copy2(db, v1)

    # Ensure v2/v3 exist on first run (copy from v1)
    if not v2.exists():
        shutil.copy2(v1, v2)
    if not v3.exists():
        shutil.copy2(v2, v3)

    for f in (db, v1, v2, v3):
        subprocess.run(["git", "add", str(f)], check=True, cwd=root)
    r = subprocess.run(
        ["git", "commit", "-m", "Update database (rotate backups, keep last 4 versions)"],
        cwd=root,
    )
    if r.returncode != 0:
        print("Nothing to commit (no changes) or commit failed.")
        sys.exit(r.returncode)
    subprocess.run(["git", "push"], check=True, cwd=root)
    print("Pushed database and 3 backups (4 versions total) to git.")


if __name__ == "__main__":
    main()
