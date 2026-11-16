import sqlite3
import uuid
from pathlib import Path

from deo_build_state import (
    copy_database,
    read_state,
    save_state,
)


def main():
    base = Path(__file__).resolve().parent
    folder = base / "database"
    database = folder / "2026-11-10-deo.db"

    active = [
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".db"
    ]

    if len(active) != 1 or active[0] != database:
        raise RuntimeError(
            "Expected only 2026-11-10-deo.db directly inside database/."
        )

    # Do not overwrite an existing tracking file.
    for path in (
        base / "config" / "build_state.json",
        folder / "build_state.json",
    ):
        if path.exists():
            raise RuntimeError(
                f"Tracking file already exists: {path}\n"
                "Stop and inspect it before initializing."
            )

    pending = list(
        (base / "google_sheet_writeback").glob("*_pending.json")
    )

    if pending:
        raise RuntimeError(
            "Pending Google writeback files exist. "
            "Resolve them before initializing."
        )

    conn = sqlite3.connect(
        database.resolve().as_uri() + "?mode=ro",
        uri=True,
    )

    try:
        tables = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }

        if "deo" not in tables:
            raise RuntimeError("The database has no deo table.")

        if "deo_build_state" in tables:
            raise RuntimeError(
                "This database already has a build-state table. "
                "Inspect its state instead of resetting it."
            )
    finally:
        conn.close()

    # Preserve a snapshot before adding tracking metadata.
    baseline_id = "baseline-" + uuid.uuid4().hex
    backup_folder = folder / "previous_versions" / baseline_id
    backup_folder.mkdir(parents=True)
    backup = backup_folder / database.name
    copy_database(database, backup)

    state = {
        "build_id": baseline_id,
        "status": "complete",
        "new_db": str(database),
        "old_db": None,
        "original_db": str(database),
        "pending_path": None,
        "writeback_started": False,
        "initialization": (
            "User-confirmed stable November 10 baseline; "
            "no pending action-log changes."
        ),
    }

    save_state(base, state)

    if read_state(base) != state:
        raise RuntimeError("Could not verify initialized build state.")

    print("Baseline initialized successfully.")
    print(f"Database: {database}")
    print(f"Backup: {backup}")
    print("Tracking saved inside SQLite and config/build_state.json.")
    print("Reopen the app, check references, then build the database.")


if __name__ == "__main__":
    main()