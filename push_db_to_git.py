"""
Push eve_manufacturing.db to git.

Only the main database file is committed; versioned backup copies are ignored by .gitignore.
Run this whenever you want to sync the DB to git.
"""

import subprocess
import sys
from pathlib import Path

DB_NAME = "eve_manufacturing.db"


def main():
    root = Path(__file__).resolve().parent
    db = root / DB_NAME

    if not db.exists():
        print(f"Database not found: {db}")
        sys.exit(1)

    subprocess.run(["git", "add", "-f", str(db)], check=True, cwd=root)
    r = subprocess.run(
        ["git", "commit", "-m", "Update database"],
        cwd=root,
    )
    if r.returncode != 0:
        print("Nothing to commit (no changes) or commit failed.")
        sys.exit(r.returncode)
    subprocess.run(["git", "push"], check=True, cwd=root)
    print("Pushed eve_manufacturing.db to git.")


if __name__ == "__main__":
    main()
