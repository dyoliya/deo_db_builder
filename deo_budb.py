import re
import sqlite3
from pathlib import Path
from collections import defaultdict
from typing import Optional, Tuple, Set

import pandas as pd


# =========================================================
# NORMALIZATION
# =========================================================

owner_changer = {
    r"\bCO\b": "COMPANY",
    r"\bCORP\b": "CORPORATION",
    r"\bEST\b": "ESTATE",
    r"\bFAM\b": "FAMILY",
    r"\bFMLY\b": "FAMILY",
    r"\bINC\b": "INCORPORATED",
    r"\bIRR\b": "IRREVOCABLE",
    r"\bIRREV\b": "IRREVOCABLE",
    r"\bIRRV\b": "IRREVOCABLE",
    r"\bIRRVCABLE\b": "IRREVOCABLE",
    r"\bIRV\b": "IRREVOCABLE",
    r"\bLIV\b": "LIVING",
    r"\bLLC\b": "LIMITED LIABILITY COMPANY",
    r"\bLP\b": "LIMITED PARTNERSHIP",
    r"\bLTD\b": "LIMITED",
    r"\bLVG\b": "LIVING",
    r"\bREV\b": "REVOCABLE",
    r"\bREVC\b": "REVOCABLE",
    r"\bREVOC\b": "REVOCABLE",
    r"\bRLT\b": "REVOCABLE LIVING TRUST",
    r"\bTR\b": "TRUST",
    r"\bTRST\b": "TRUST",
    r"\bTRT\b": "TRUST",
    r"\bTRTEE\b": "TRUSTEE",
    r"\bTST\b": "TRUST",
    r"\bTSTE\b": "TRUSTEE",
    r"\bTSTEE\b": "TRUSTEE",
    r"\bTSTEES\b": "TRUSTEES",
    r"\bTTEE\b": "TRUSTEE",
}

address_replacements = {
    r"\bSTATE HWY\b": "HWY",
    r"\bNORTHEAST\b": "NE",
    r"\bNORTHWEST\b": "NW",
    r"\bSOUTHEAST\b": "SE",
    r"\bSOUTHWEST\b": "SW",
    r"\bAPARTMENT\b": "APT",
    r"\bAVENUE\b": "AVE",
    r"\bBOULEVARD\b": "BLVD",
    r"\bCIRCLE\b": "CR",
    r"\bCOURT\b": "CT",
    r"\bDRIVE\b": "DR",
    r"\bEAST\b": "E",
    r"\bHIGHWAY\b": "HWY",
    r"\bLANE\b": "LN",
    r"\bNORTH\b": "N",
    r"\bPARKWAY\b": "PKWY",
    r"\bROAD\b": "RD",
    r"\bSOUTH\b": "S",
    r"\bSTREET\b": "ST",
    r"\bSUITE\b": "STE",
    r"\bTRAIL\b": "TRL",
    r"\bWEST\b": "W",
    r"\bP\.?\s*O\.?\b": "PO",
    r"\bFREEWAY\b": "FWY",
}

OWNER_REPLACEMENTS = [
    (re.compile(pattern, flags=re.IGNORECASE), repl)
    for pattern, repl in owner_changer.items()
]

ADDRESS_REPLACEMENTS = [
    (re.compile(pattern, flags=re.IGNORECASE), repl)
    for pattern, repl in address_replacements.items()
]

_BASIC_CACHE = {}
_OWNER_CACHE = {}
_ADDRESS_CACHE = {}
_CITYSTATE_CACHE = {}
_MATCH_KEY_CACHE = {}

def normalize_basic_text(value) -> str:
    if value is None:
        return ""

    raw = str(value).strip()
    if raw.lower() in {"", "nan", "none", "null"}:
        return ""

    cached = _BASIC_CACHE.get(raw)
    if cached is not None:
        return cached

    s = raw.upper()
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    _BASIC_CACHE[raw] = s
    return s


def apply_replacements(text: str, compiled_replacements) -> str:
    if not text:
        return ""

    for pattern, repl in compiled_replacements:
        text = pattern.sub(repl, text)

    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_owner_text(value) -> str:
    if value is None:
        return ""

    raw = str(value).strip()
    if raw.lower() in {"", "nan", "none", "null"}:
        return ""

    cached = _OWNER_CACHE.get(raw)
    if cached is not None:
        return cached

    s = normalize_basic_text(raw)
    if not s:
        return ""

    s = apply_replacements(s, OWNER_REPLACEMENTS)
    _OWNER_CACHE[raw] = s
    return s


def normalize_address_text(value) -> str:
    if value is None:
        return ""

    raw = str(value).strip()
    if raw.lower() in {"", "nan", "none", "null"}:
        return ""

    cached = _ADDRESS_CACHE.get(raw)
    if cached is not None:
        return cached

    s = normalize_basic_text(raw)
    if not s:
        return ""

    s = apply_replacements(s, ADDRESS_REPLACEMENTS)
    _ADDRESS_CACHE[raw] = s
    return s


def normalize_city_state_text(value) -> str:
    if value is None:
        return ""

    raw = str(value).strip()
    if raw.lower() in {"", "nan", "none", "null"}:
        return ""

    cached = _CITYSTATE_CACHE.get(raw)
    if cached is not None:
        return cached

    s = normalize_basic_text(raw)
    _CITYSTATE_CACHE[raw] = s
    return s

def build_match_key(owner, address, city, state) -> Optional[Tuple[str, str, str, str]]:
    raw_key = (
        "" if owner is None else str(owner).strip(),
        "" if address is None else str(address).strip(),
        "" if city is None else str(city).strip(),
        "" if state is None else str(state).strip(),
    )

    cached = _MATCH_KEY_CACHE.get(raw_key)
    if cached is not None:
        return cached

    owner_n = normalize_owner_text(raw_key[0])
    address_n = normalize_address_text(raw_key[1])
    city_n = normalize_city_state_text(raw_key[2])
    state_n = normalize_city_state_text(raw_key[3])

    if not address_n:
        _MATCH_KEY_CACHE[raw_key] = None
        return None

    result = (owner_n, address_n, city_n, state_n)
    _MATCH_KEY_CACHE[raw_key] = result
    return result

def serialize_key(key: Tuple[str, str, str, str]) -> str:
    return "||".join(key)


def deserialize_key(key_str: str) -> Tuple[str, str, str, str]:
    parts = key_str.split("||")
    if len(parts) != 4:
        raise ValueError(f"Invalid cached key format: {key_str}")
    return tuple(parts)


# =========================================================
# BUDB FILE DISCOVERY
# =========================================================

def get_single_budb_path(base_dir: str | Path) -> Path:
    base_dir = Path(base_dir)
    budb_dir = base_dir / "budb"

    if not budb_dir.exists():
        raise FileNotFoundError(f"BUDB folder not found: {budb_dir}")

    db_files = sorted([p for p in budb_dir.glob("*.db") if p.is_file()])

    if len(db_files) == 0:
        raise FileNotFoundError(f"No .db file found in BUDB folder: {budb_dir}")

    if len(db_files) > 1:
        names = "\n  - ".join(p.name for p in db_files)
        raise ValueError(
            "Expected exactly 1 BUDB .db file in budb folder, but found multiple:\n"
            f"  - {names}"
        )

    return db_files[0]


# =========================================================
# SQLITE HELPERS
# =========================================================

def connect_sqlite_readonly(db_path: str | Path) -> sqlite3.Connection:
    return sqlite3.connect(str(db_path))


def get_existing_columns(conn: sqlite3.Connection, table_name: str) -> list[str]:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info([{table_name}])")
    return [row[1] for row in cur.fetchall()]


def validate_budb_columns(existing_cols: list[str]):
    required = {
        "id",
        "contact_group_id",
        "Owner",
        "Owner (Standardized)",
        "Input: Address",
        "Input: City",
        "Input: State",
        "md_address",
        "md_city",
        "md_state",
    }
    existing = set(existing_cols)
    missing = sorted(required - existing)

    if missing:
        raise ValueError(
            "BUDB is missing required columns:\n  - " + "\n  - ".join(missing)
        )


# =========================================================
# CACHE HELPERS
# =========================================================

def get_cache_path(base_dir: str | Path) -> Path:
    base_dir = Path(base_dir)
    cache_dir = base_dir / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / "budb_lookup_cache.db"


def get_budb_file_signature(budb_path: str | Path) -> dict:
    budb_path = Path(budb_path)
    stat = budb_path.stat()
    return {
        "budb_path": str(budb_path.resolve()),
        "budb_mtime": stat.st_mtime,
        "budb_size": stat.st_size,
    }


def connect_cache_db(cache_path: str | Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(cache_path))
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def ensure_cache_tables(conn: sqlite3.Connection):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cache_meta (
            meta_key TEXT PRIMARY KEY,
            meta_value TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS lookup_cache (
            combo_name TEXT NOT NULL,
            match_key TEXT NOT NULL,
            budb_id TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_lookup_cache_combo_key
        ON lookup_cache (combo_name, match_key)
    """)

    conn.commit()


def read_cache_meta(conn: sqlite3.Connection) -> dict:
    cur = conn.cursor()
    cur.execute("SELECT meta_key, meta_value FROM cache_meta")
    rows = cur.fetchall()
    return {k: v for k, v in rows}


def write_cache_meta(conn: sqlite3.Connection, meta: dict):
    conn.execute("DELETE FROM cache_meta")
    conn.executemany(
        "INSERT INTO cache_meta (meta_key, meta_value) VALUES (?, ?)",
        [(k, str(v)) for k, v in meta.items()]
    )
    conn.commit()


def clear_lookup_cache(conn: sqlite3.Connection):
    conn.execute("DELETE FROM lookup_cache")
    conn.commit()


def cache_is_current(conn: sqlite3.Connection, budb_signature: dict) -> bool:
    meta = read_cache_meta(conn)
    if not meta:
        return False

    try:
        cached_path = meta.get("budb_path", "")
        cached_mtime = float(meta.get("budb_mtime", "0"))
        cached_size = int(meta.get("budb_size", "0"))
    except Exception:
        return False

    return (
        cached_path == budb_signature["budb_path"]
        and cached_mtime == budb_signature["budb_mtime"]
        and cached_size == budb_signature["budb_size"]
    )


def save_lookup_maps_to_cache(conn: sqlite3.Connection, lookup_maps: dict, budb_signature: dict, metadata: dict):
    clear_lookup_cache(conn)

    rows_to_insert = []
    for combo_name, combo_map in lookup_maps.items():
        for key_tuple, budb_ids in combo_map.items():
            key_str = serialize_key(key_tuple)
            for budb_id in budb_ids:
                rows_to_insert.append((combo_name, key_str, str(budb_id)))

    conn.executemany(
        "INSERT INTO lookup_cache (combo_name, match_key, budb_id) VALUES (?, ?, ?)",
        rows_to_insert
    )

    meta = {
        "budb_path": budb_signature["budb_path"],
        "budb_mtime": budb_signature["budb_mtime"],
        "budb_size": budb_signature["budb_size"],
        "table_name": metadata.get("table_name", "bottoms_up"),
        "rows_loaded": metadata.get("rows_loaded", 0),
        "cache_mode": "rebuilt",
    }
    write_cache_meta(conn, meta)
    conn.commit()


def load_lookup_maps_from_cache(conn: sqlite3.Connection) -> tuple[dict, dict]:
    combo_1 = defaultdict(set)
    combo_2 = defaultdict(set)
    combo_3 = defaultdict(set)
    combo_4 = defaultdict(set)

    combo_map_lookup = {
        "combo_1_owner_input": combo_1,
        "combo_2_ownerstd_input": combo_2,
        "combo_3_ownerstd_md": combo_3,
        "combo_4_owner_md": combo_4,
    }

    cur = conn.cursor()
    cur.execute("""
        SELECT combo_name, match_key, budb_id
        FROM lookup_cache
    """)
    rows = cur.fetchall()

    for combo_name, match_key, budb_id in rows:
        target_map = combo_map_lookup.get(combo_name)
        if target_map is None:
            continue
        target_map[deserialize_key(match_key)].add(str(budb_id))

    meta = read_cache_meta(conn)

    lookup_maps = {
        "combo_1_owner_input": combo_1,
        "combo_2_ownerstd_input": combo_2,
        "combo_3_ownerstd_md": combo_3,
        "combo_4_owner_md": combo_4,
    }

    metadata = {
        "budb_path": meta.get("budb_path", ""),
        "table_name": meta.get("table_name", "bottoms_up"),
        "rows_loaded": int(meta.get("rows_loaded", "0")),
        "cache_mode": "loaded_from_cache",
    }

    return lookup_maps, metadata


# =========================================================
# LOAD BUDB + BUILD LOOKUP MAPS
# =========================================================

def load_budb_dataframe(budb_path: str | Path) -> tuple[pd.DataFrame, str]:
    conn = connect_sqlite_readonly(budb_path)
    try:
        table_name = "bottoms_up"

        existing_cols = get_existing_columns(conn, table_name)
        if not existing_cols:
            raise ValueError(f"Table '{table_name}' not found in BUDB.")

        validate_budb_columns(existing_cols)

        query = f"""
            SELECT
                [id] AS budb_id,
                [contact_group_id] AS contact_group_id,
                [Owner] AS owner,
                [Owner (Standardized)] AS owner_standardized,
                [Input: Address] AS input_address,
                [Input: City] AS input_city,
                [Input: State] AS input_state,
                [md_address] AS md_address,
                [md_city] AS md_city,
                [md_state] AS md_state
            FROM [{table_name}]
        """
        df = pd.read_sql_query(query, conn)
        return df, table_name
    finally:
        conn.close()


def build_lookup_maps_from_budb_df(budb_df: pd.DataFrame) -> dict:
    combo_1 = defaultdict(set)
    combo_2 = defaultdict(set)
    combo_3 = defaultdict(set)
    combo_4 = defaultdict(set)

    id_to_group = {}
    group_to_ids = defaultdict(set)

    for row in budb_df.itertuples(index=False):
        budb_id = "" if pd.isna(row.budb_id) else str(row.budb_id).strip()
        if not budb_id:
            continue

        # Build contact_group_id mappings
        contact_group_id = "" if pd.isna(row.contact_group_id) else str(row.contact_group_id).strip()
        if contact_group_id and contact_group_id.lower() not in {"nan", "none", "null"}:
            id_to_group[budb_id] = contact_group_id
            group_to_ids[contact_group_id].add(budb_id)

        key1 = build_match_key(
            row.owner,
            row.input_address,
            row.input_city,
            row.input_state,
        )
        if key1:
            combo_1[key1].add(budb_id)

        key2 = build_match_key(
            row.owner_standardized,
            row.input_address,
            row.input_city,
            row.input_state,
        )
        if key2:
            combo_2[key2].add(budb_id)

        key3 = build_match_key(
            row.owner_standardized,
            row.md_address,
            row.md_city,
            row.md_state,
        )
        if key3:
            combo_3[key3].add(budb_id)

        key4 = build_match_key(
            row.owner,
            row.md_address,
            row.md_city,
            row.md_state,
        )
        if key4:
            combo_4[key4].add(budb_id)

    return {
        "combo_1_owner_input": combo_1,
        "combo_2_ownerstd_input": combo_2,
        "combo_3_ownerstd_md": combo_3,
        "combo_4_owner_md": combo_4,
        "id_to_group": id_to_group,
        "group_to_ids": group_to_ids,
    }

def load_budb_lookup_maps(base_dir: str | Path):
    """
    Safe cache strategy:
    - if BUDB file path + mtime + size are unchanged, load lookup maps from cache
    - otherwise rebuild lookup maps from BUDB and refresh cache
    """
    base_dir = Path(base_dir)
    budb_path = get_single_budb_path(base_dir)
    budb_signature = get_budb_file_signature(budb_path)

    cache_path = get_cache_path(base_dir)
    cache_conn = connect_cache_db(cache_path)

    try:
        ensure_cache_tables(cache_conn)

        if cache_is_current(cache_conn, budb_signature):
            lookup_maps, metadata = load_lookup_maps_from_cache(cache_conn)

            # Rebuild contact_group expansion maps because old cache does not store these.
            budb_df, table_name = load_budb_dataframe(budb_path)

            id_to_group = {}
            group_to_ids = defaultdict(set)

            for row in budb_df.itertuples(index=False):
                budb_id = "" if pd.isna(row.budb_id) else str(row.budb_id).strip()
                contact_group_id = "" if pd.isna(row.contact_group_id) else str(row.contact_group_id).strip()

                if budb_id and contact_group_id and contact_group_id.lower() not in {"nan", "none", "null"}:
                    id_to_group[budb_id] = contact_group_id
                    group_to_ids[contact_group_id].add(budb_id)

            lookup_maps["id_to_group"] = id_to_group
            lookup_maps["group_to_ids"] = group_to_ids

            metadata["cache_path"] = str(cache_path)
            metadata["table_name"] = table_name
            metadata["rows_loaded"] = len(budb_df)

            return lookup_maps, metadata

        budb_df, table_name = load_budb_dataframe(budb_path)
        lookup_maps = build_lookup_maps_from_budb_df(budb_df)

        metadata = {
            "budb_path": str(budb_path),
            "table_name": table_name,
            "rows_loaded": len(budb_df),
            "cache_mode": "rebuilt_from_budb",
            "cache_path": str(cache_path),
        }

        save_lookup_maps_to_cache(cache_conn, lookup_maps, budb_signature, metadata)
        return lookup_maps, metadata

    finally:
        cache_conn.close()


# =========================================================
# MATCHING
# =========================================================

def match_budb_ids_for_deo_row(row, lookup_maps: dict) -> Optional[str]:
    deo_combos = [
        ("owner", "input_address", "input_city", "input_state"),
        ("owner_standardized", "input_address", "input_city", "input_state"),
        ("owner", "md_address", "md_city", "md_state"),
        ("owner_standardized", "md_address", "md_city", "md_state"),
        # ("owner_name", "address2", "city2", "state2"),
        # ("owner_name", "add_address1", "add_address1_city", "add_address1_state"),
    ]

    direct_matches: Set[str] = set()

    for owner_col, address_col, city_col, state_col in deo_combos:
        key = build_match_key(
            getattr(row, owner_col, ""),
            getattr(row, address_col, ""),
            getattr(row, city_col, ""),
            getattr(row, state_col, ""),
        )

        if not key:
            continue

        for combo_name in (
            "combo_1_owner_input",
            "combo_2_ownerstd_input",
            "combo_3_ownerstd_md",
            "combo_4_owner_md",
        ):
            combo_map = lookup_maps.get(combo_name, {})
            direct_matches.update(combo_map.get(key, set()))

        if direct_matches:
            break

    if not direct_matches:
        return None

    id_to_group = lookup_maps.get("id_to_group", {})
    group_to_ids = lookup_maps.get("group_to_ids", {})

    expanded_matches: Set[str] = set(direct_matches)

    for budb_id in direct_matches:
        contact_group_id = id_to_group.get(budb_id)
        if contact_group_id:
            expanded_matches.update(group_to_ids.get(contact_group_id, set()))

    def sort_key(x: str):
        return (0, int(x)) if x.isdigit() else (1, x)

    return " | ".join(sorted(expanded_matches, key=sort_key))


def populate_budb_ids_in_df(df: pd.DataFrame, lookup_maps: dict, progress_callback=None) -> pd.DataFrame:
    df = df.copy()
    total = len(df)
    budb_ids = []

    for idx, row in enumerate(df.itertuples(index=False), start=1):
        budb_ids.append(match_budb_ids_for_deo_row(row, lookup_maps))
        if progress_callback and total:
            progress_callback(idx / total)

    df["budb_id"] = budb_ids
    return df