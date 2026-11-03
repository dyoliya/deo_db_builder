import os
import re
import json
import sqlite3
from glob import glob
from typing import List, Dict, Any, Optional

# ============================================================
# CONFIG
# ============================================================
OLD_DIR = "old"
NEW_DIR = "new"

# Set to None to auto-detect one common table between old/new DBs.
# Or set explicitly, e.g. TABLE_NAME = "deo"
TABLE_NAME = None

PRIMARY_KEY = "deo_id"
IGNORE_FIELDS = {"date_uploaded"}

LOG_TABLE_NAME = "deo_action_log"


# ============================================================
# HELPERS
# ============================================================
def find_single_db(folder: str) -> str:
    db_files = sorted(glob(os.path.join(folder, "*.db")))
    if not db_files:
        raise FileNotFoundError(f"No .db file found in folder: {folder}")
    if len(db_files) > 1:
        raise ValueError(
            f"Expected only 1 .db file in folder '{folder}', found {len(db_files)}:\n" +
            "\n".join(db_files)
        )
    return db_files[0]


def extract_date_from_filename(path: str) -> Optional[str]:
    """
    Extract YYYY-MM-DD from filename like:
    2026-03-12-deo.db
    """
    filename = os.path.basename(path)
    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if match:
        return match.group(1)
    return None


def connect_db(path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def list_user_tables(conn: sqlite3.Connection) -> List[str]:
    cur = conn.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
          AND name NOT LIKE 'sqlite_%'
        ORDER BY name
    """)
    return [row["name"] for row in cur.fetchall()]


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    cur = conn.execute("""
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table'
          AND name = ?
        LIMIT 1
    """, (table_name,))
    return cur.fetchone() is not None


def resolve_table_name(old_conn: sqlite3.Connection, new_conn: sqlite3.Connection) -> str:
    if TABLE_NAME:
        return TABLE_NAME

    old_tables = set(list_user_tables(old_conn))
    new_tables = set(list_user_tables(new_conn))

    # ignore log table when auto-detecting
    old_tables.discard(LOG_TABLE_NAME)
    new_tables.discard(LOG_TABLE_NAME)

    common_tables = sorted(old_tables & new_tables)

    if not common_tables:
        raise ValueError("No common user table found between old and new databases.")

    if len(common_tables) > 1:
        raise ValueError(
            "Multiple common tables found. Please set TABLE_NAME explicitly.\n"
            f"Common tables: {common_tables}"
        )

    return common_tables[0]


def get_table_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    cur = conn.execute(f'PRAGMA table_info("{table_name}")')
    return [row["name"] for row in cur.fetchall()]


def fetch_rows_by_pk(conn: sqlite3.Connection, table_name: str, pk: str) -> Dict[Any, Dict[str, Any]]:
    cur = conn.execute(f'SELECT * FROM "{table_name}"')
    rows = {}
    for row in cur.fetchall():
        row_dict = dict(row)
        pk_value = row_dict.get(pk)
        if pk_value is None:
            continue
        if pk_value in rows:
            raise ValueError(f"Duplicate {pk} found in table '{table_name}': {pk_value}")
        rows[pk_value] = row_dict
    return rows


def normalize_value(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, bytes):
        try:
            return value.decode("utf-8")
        except Exception:
            return repr(value)

    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value

    return value


def values_equal(a: Any, b: Any) -> bool:
    return normalize_value(a) == normalize_value(b)


def stringify(value: Any) -> Optional[str]:
    value = normalize_value(value)
    if value is None:
        return None
    return str(value)


def row_to_json(row: Dict[str, Any]) -> str:
    normalized = {k: normalize_value(v) for k, v in row.items()}
    return json.dumps(normalized, ensure_ascii=False, default=str)


# ============================================================
# LOG TABLE
# ============================================================
def create_log_table_if_needed(conn: sqlite3.Connection) -> None:
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS {LOG_TABLE_NAME} (
            action_id INTEGER PRIMARY KEY AUTOINCREMENT,
            deo_id TEXT,
            field TEXT,
            old_value TEXT,
            new_value TEXT,
            status TEXT CHECK(status IN ('added', 'modified', 'deleted')),
            modified_date TEXT
        )
    """)
    conn.commit()


def get_log_table_row_count(conn: sqlite3.Connection) -> int:
    cur = conn.execute(f"SELECT COUNT(*) AS cnt FROM {LOG_TABLE_NAME}")
    return cur.fetchone()["cnt"]


def copy_old_logs_to_new_if_needed(old_conn: sqlite3.Connection, new_conn: sqlite3.Connection) -> int:
    """
    Copy historical logs from old DB into new DB only if:
    - old DB has the log table
    - new DB log table is currently empty

    Returns number of copied rows.
    """
    if not table_exists(old_conn, LOG_TABLE_NAME):
        return 0

    new_count = get_log_table_row_count(new_conn)
    if new_count > 0:
        # assume history is already present; do not duplicate
        return 0

    old_columns = get_table_columns(old_conn, LOG_TABLE_NAME)
    required_columns = ["deo_id", "field", "old_value", "new_value", "status", "modified_date"]

    missing = [col for col in required_columns if col not in old_columns]
    if missing:
        raise ValueError(
            f"Old DB log table '{LOG_TABLE_NAME}' is missing required columns: {missing}"
        )

    cur = old_conn.execute(f"""
        SELECT deo_id, field, old_value, new_value, status, modified_date
        FROM {LOG_TABLE_NAME}
        ORDER BY action_id
    """)
    rows = cur.fetchall()

    if not rows:
        return 0

    new_conn.executemany(f"""
        INSERT INTO {LOG_TABLE_NAME} (
            deo_id, field, old_value, new_value, status, modified_date
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, [
        (
            row["deo_id"],
            row["field"],
            row["old_value"],
            row["new_value"],
            row["status"],
            row["modified_date"],
        )
        for row in rows
    ])
    new_conn.commit()
    return len(rows)


def insert_log(
    conn: sqlite3.Connection,
    deo_id: Optional[Any],
    field: str,
    old_value: Optional[str],
    new_value: Optional[str],
    status: str,
    modified_date: Optional[str],
) -> None:
    conn.execute(f"""
        INSERT INTO {LOG_TABLE_NAME} (
            deo_id, field, old_value, new_value, status, modified_date
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        None if deo_id is None else str(deo_id),
        field,
        old_value,
        new_value,
        status,
        modified_date
    ))


# ============================================================
# MAIN COMPARISON
# ============================================================
def compare_databases():
    old_db_path = find_single_db(OLD_DIR)
    new_db_path = find_single_db(NEW_DIR)

    comparison_date = extract_date_from_filename(new_db_path)
    if comparison_date is None:
        raise ValueError(
            f"Could not extract date from new DB filename: {os.path.basename(new_db_path)}\n"
            "Expected format like: 2026-03-12-deo.db"
        )

    print(f"Old DB: {old_db_path}")
    print(f"New DB: {new_db_path}")
    print(f"Comparison date (from new DB filename): {comparison_date}")
    print(f"Log table will be stored inside the NEW DB: {LOG_TABLE_NAME}")

    old_conn = connect_db(old_db_path)
    new_conn = connect_db(new_db_path)

    try:
        create_log_table_if_needed(new_conn)

        copied_history_count = copy_old_logs_to_new_if_needed(old_conn, new_conn)
        print(f"Copied prior log rows from OLD DB: {copied_history_count}")

        table_name = resolve_table_name(old_conn, new_conn)
        print(f"Using table: {table_name}")

        old_columns = get_table_columns(old_conn, table_name)
        new_columns = get_table_columns(new_conn, table_name)

        if PRIMARY_KEY not in old_columns:
            raise ValueError(f"'{PRIMARY_KEY}' not found in OLD table '{table_name}'")
        if PRIMARY_KEY not in new_columns:
            raise ValueError(f"'{PRIMARY_KEY}' not found in NEW table '{table_name}'")

        old_col_set = set(old_columns)
        new_col_set = set(new_columns)

        added_columns = sorted(new_col_set - old_col_set)
        deleted_columns = sorted(old_col_set - new_col_set)

        comparable_columns = sorted((old_col_set & new_col_set) - IGNORE_FIELDS)

        old_rows = fetch_rows_by_pk(old_conn, table_name, PRIMARY_KEY)
        new_rows = fetch_rows_by_pk(new_conn, table_name, PRIMARY_KEY)

        old_ids = set(old_rows.keys())
        new_ids = set(new_rows.keys())

        added_ids = sorted(new_ids - old_ids, key=lambda x: str(x))
        deleted_ids = sorted(old_ids - new_ids, key=lambda x: str(x))
        common_ids = sorted(old_ids & new_ids, key=lambda x: str(x))

        added_count = 0
        deleted_count = 0
        modified_count = 0
        schema_count = len(added_columns) + len(deleted_columns)

        # log schema changes
        for col in added_columns:
            insert_log(
                conn=new_conn,
                deo_id=None,
                field=col,
                old_value=None,
                new_value=json.dumps({"column_added": col}),
                status="added",
                modified_date=comparison_date
            )

        for col in deleted_columns:
            insert_log(
                conn=new_conn,
                deo_id=None,
                field=col,
                old_value=json.dumps({"column_deleted": col}),
                new_value=None,
                status="deleted",
                modified_date=comparison_date
            )

        # added rows
        for deo_id in added_ids:
            new_row = new_rows[deo_id]
            insert_log(
                conn=new_conn,
                deo_id=deo_id,
                field="all_fields",
                old_value=None,
                new_value=row_to_json(new_row),
                status="added",
                modified_date=comparison_date
            )
            added_count += 1

        # deleted rows
        for deo_id in deleted_ids:
            old_row = old_rows[deo_id]
            insert_log(
                conn=new_conn,
                deo_id=deo_id,
                field="all_fields",
                old_value=row_to_json(old_row),
                new_value=None,
                status="deleted",
                modified_date=comparison_date
            )
            deleted_count += 1

        # modified fields
        for deo_id in common_ids:
            old_row = old_rows[deo_id]
            new_row = new_rows[deo_id]

            for col in comparable_columns:
                if col == PRIMARY_KEY:
                    continue

                old_val = old_row.get(col)
                new_val = new_row.get(col)

                if not values_equal(old_val, new_val):
                    insert_log(
                        conn=new_conn,
                        deo_id=deo_id,
                        field=col,
                        old_value=stringify(old_val),
                        new_value=stringify(new_val),
                        status="modified",
                        modified_date=comparison_date
                    )
                    modified_count += 1

        new_conn.commit()

        print("\nDone.")
        print(f"Copied historical logs : {copied_history_count}")
        print(f"Schema changes logged  : {schema_count}")
        print(f"Added rows logged      : {added_count}")
        print(f"Deleted rows logged    : {deleted_count}")
        print(f"Modified fields logged : {modified_count}")
        print(f"Total new log entries  : {schema_count + added_count + deleted_count + modified_count}")

    finally:
        old_conn.close()
        new_conn.close()


if __name__ == "__main__":
    compare_databases()