import json
import os
import re
import sqlite3
import uuid
from pathlib import Path


def state_path(base):
    return Path(base) / "config" / "build_state.json"


def _write_json(base, state):
    path = state_path(base)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")

    with temporary.open("w", encoding="utf-8") as stream:
        json.dump(state, stream, indent=2)
        stream.flush()
        os.fsync(stream.fileno())

    os.replace(temporary, path)


def _state_database(state):
    if state["status"] == "rolled_back":
        return (
            Path(state["original_db"])
            if state["original_db"]
            else None
        )

    target = Path(state["new_db"])

    if target.is_file():
        return target

    if state["status"] == "building" and state["original_db"]:
        return Path(state["original_db"])

    return target


def _read_database_state(path):
    conn = sqlite3.connect(
        path.resolve().as_uri() + "?mode=ro",
        uri=True,
    )

    try:
        exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' "
            "AND name='deo_build_state'"
        ).fetchone()

        if not exists:
            return None

        row = conn.execute(
            "SELECT state_json FROM deo_build_state WHERE singleton = 1"
        ).fetchone()

        return json.loads(row[0]) if row else None

    finally:
        conn.close()


def save_state(base, state):
    # Save SQLite first so a failed JSON write cannot lose the new status.
    database = _state_database(state)

    if database is not None:
        if not database.is_file():
            if state["status"] != "building" or state["original_db"]:
                raise RuntimeError(
                    f"Build-state database is missing: {database}"
                )

            database.parent.mkdir(parents=True, exist_ok=True)

        conn = sqlite3.connect(database)

        try:
            with conn:
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS deo_build_state (
                        singleton INTEGER PRIMARY KEY CHECK (singleton = 1),
                        state_json TEXT NOT NULL
                    )
                """)

                conn.execute(
                    "INSERT OR REPLACE INTO deo_build_state "
                    "(singleton, state_json) VALUES (1, ?)",
                    (json.dumps(state),),
                )

        finally:
            conn.close()

    _write_json(base, state)


def read_state(base):
    base = Path(base).resolve()
    path = state_path(base)
    legacy = base / "database" / "build_state.json"
    active = sorted((base / "database").glob("*.db"))

    # Use the old location only when the new JSON does not exist.
    candidate = path if path.exists() else legacy
    state = None

    if candidate.exists():
        try:
            state = json.loads(candidate.read_text(encoding="utf-8"))

            required = {
                "build_id",
                "status",
                "new_db",
                "old_db",
                "original_db",
                "pending_path",
                "writeback_started",
            }

            if not isinstance(state, dict) or not required.issubset(state):
                state = None

        except (ValueError, OSError):
            state = None

    if state is not None:
        database = _state_database(state)

        if database is not None:
            if not database.is_file():
                raise RuntimeError(
                    "Recorded build database is missing; recovery required."
                )

            stored = _read_database_state(database)

            if stored is not None:
                if stored["build_id"] != state["build_id"]:
                    raise RuntimeError(
                        "Build tracking disagrees; recovery required."
                    )

                # SQLite has the authoritative status for this build.
                state = stored

            elif candidate == legacy:
                # First upgraded launch: preserve the existing workflow state.
                save_state(base, state)

            else:
                raise RuntimeError(
                    "Database build tracking is missing; recovery required."
                )

        elif active:
            raise RuntimeError(
                "Unexpected active database; recovery required."
            )

    else:
        # Without usable JSON, recover only from one identifiable active DB.
        if len(active) != 1:
            archive = base / "database" / "previous_versions"

            if active or candidate.exists() or archive.exists():
                raise RuntimeError(
                    "Cannot recover build tracking automatically."
                )

            # No databases, tracking file, or archive: a fresh installation.
            return None

        state = _read_database_state(active[0])

        if state is None:
            raise RuntimeError(
                "An existing database has no build tracking. "
                "Restore its build_state.json before continuing."
            )

        expected = _state_database(state)

        if expected is None or expected.resolve() != active[0].resolve():
            raise RuntimeError(
                "Recorded database location changed; recovery required."
            )

    if state["status"] == "ready_for_log":
        if Path(state["pending_path"]).exists():
            raise RuntimeError(
                "Google writeback is pending; recovery required."
            )

        if state["old_db"] and not Path(state["old_db"]).is_file():
            raise RuntimeError(
                "The comparison snapshot is missing; recovery required."
            )

    # Recreate or refresh the supporting JSON.
    _write_json(base, state)

    # Remove the obsolete copy only after the replacement is saved.
    legacy.unlink(missing_ok=True)

    return state


def copy_database(source, destination):
    """Create a consistent SQLite snapshot, including committed WAL data."""
    source = Path(source)
    destination = Path(destination)

    src = sqlite3.connect(
        source.resolve().as_uri() + "?mode=ro",
        uri=True,
    )

    try:
        dst = sqlite3.connect(destination)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()


def prepare_build(base, day, pending_path):
    base = Path(base).resolve()
    previous = read_state(base)

    if previous and previous["status"] not in (
        "complete",
        "rolled_back",
    ):
        raise RuntimeError(
            "The previous build is not finished. "
            "Update its action log, or resolve its recovery state "
            "before building again."
        )

    if Path(pending_path).exists():
        raise RuntimeError(
            "Resolve the pending Google writeback before building."
        )

    folder = base / "database"
    folder.mkdir(parents=True, exist_ok=True)

    sources = sorted(folder.glob("*.db"))

    if len(sources) > 1:
        raise RuntimeError(
            "Expected one active DEO database in database/."
        )

    source = sources[0] if sources else None

    if source and not re.fullmatch(
        r"\d{4}-\d{2}-\d{2}-deo\.db",
        source.name,
    ):
        raise RuntimeError(
            "The active database must be named YYYY-MM-DD-deo.db."
        )

    if source and source.name[:10] > day:
        raise RuntimeError(
            "The active database is dated after today's build date."
        )

    run_id = uuid.uuid4().hex
    archive = folder / "previous_versions" / run_id
    archive.mkdir(parents=True)

    baseline = archive / source.name if source else None

    if source:
        copy_database(source, baseline)

    target = folder / f"{day}-deo.db"

    state = {
        "build_id": run_id,
        "status": "building",
        "new_db": str(target),
        "old_db": str(baseline) if baseline else None,
        "original_db": str(source) if source else None,
        "pending_path": str(Path(pending_path).resolve()),
        "writeback_started": False,
    }

    save_state(base, state)

    try:
        if source and source != target:
            copy_database(baseline, target)

        # Copying the baseline also copies its older tracking table.
        # Replace that metadata with this build's state.
        save_state(base, state)

    except Exception:
        fail_build(base)
        raise

    return state


def mark_writeback_started(base):
    state = read_state(base)

    if not state or state["status"] != "building":
        raise RuntimeError(
            "Missing active build state; Google writeback stopped."
        )

    state["writeback_started"] = True
    save_state(base, state)


def finish_build(base):
    state = read_state(base)

    if Path(state["pending_path"]).exists():
        raise RuntimeError("Google writeback is still pending.")

    original = (
        Path(state["original_db"])
        if state["original_db"]
        else None
    )
    target = Path(state["new_db"])

    # The consistent prior snapshot is already in previous_versions/.
    # Retire the old active file only after the build succeeds.
    if original and original != target:
        original.unlink()

        for suffix in ("-wal", "-shm"):
            Path(str(original) + suffix).unlink(missing_ok=True)

    state["status"] = "ready_for_log"
    save_state(base, state)


def fail_build(base):
    state = read_state(base)

    if state["writeback_started"]:
        state["status"] = "recovery_required"
        save_state(base, state)

        return (
            "Database preserved. Google writeback may have occurred; "
            "recovery is required."
        )

    target = Path(state["new_db"])
    original = (
        Path(state["original_db"])
        if state["original_db"]
        else None
    )

    if original == target:
        # Same-day build: restore the pre-build contents.
        copy_database(state["old_db"], target)
    else:
        # New-day build: the previous active file was left untouched.
        target.unlink(missing_ok=True)

        for suffix in ("-wal", "-shm"):
            Path(str(target) + suffix).unlink(missing_ok=True)

    # prepare_build verified this journal did not exist beforehand.
    Path(state["pending_path"]).unlink(missing_ok=True)

    state["status"] = "rolled_back"
    save_state(base, state)

    return "Previous database restored. No action log was run."