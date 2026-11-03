import re
import sqlite3
from pathlib import Path
from collections import defaultdict
from typing import Optional

import pandas as pd

# ============================================
# CONFIG
# ============================================

PIPEDRIVE_LOCAL_CACHE_FILENAME = "pipedrive_local_cache.csv"

# ============================================
# SQLITE CACHE HELPERS
# ============================================

def get_cache_path(base_dir: str | Path) -> Path:
    base_dir = Path(base_dir)
    cache_dir = base_dir / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / "pipedrive_lookup_cache.db"


def connect_cache_db(cache_path: str | Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(cache_path))
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def ensure_cache_tables(conn: sqlite3.Connection):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pipedrive_deals_cache (
            deal_id TEXT,
            title TEXT,
            contact_name TEXT,
            mailing_address_raw TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cache_meta (
            meta_key TEXT PRIMARY KEY,
            meta_value TEXT
        )
    """)
    conn.commit()


def clear_pipedrive_cache(conn: sqlite3.Connection):
    conn.execute("DELETE FROM pipedrive_deals_cache")
    conn.execute("DELETE FROM cache_meta")
    conn.commit()


def save_pipedrive_df_to_cache(conn: sqlite3.Connection, df: pd.DataFrame):
    clear_pipedrive_cache(conn)

    rows = []
    for row in df.itertuples(index=False):
        rows.append((
            None if pd.isna(row.deal_id) else str(row.deal_id),
            "" if pd.isna(row.title) else str(row.title),
            "" if pd.isna(row.contact_name) else str(row.contact_name),
            "" if pd.isna(row.mailing_address_raw) else str(row.mailing_address_raw),
        ))

    conn.executemany("""
        INSERT INTO pipedrive_deals_cache (
            deal_id,
            title,
            contact_name,
            mailing_address_raw
        )
        VALUES (?, ?, ?, ?)
    """, rows)

    conn.executemany("""
        INSERT INTO cache_meta (meta_key, meta_value)
        VALUES (?, ?)
    """, [
        ("rows_loaded", str(len(df))),
        ("cache_mode", "rebuilt_from_pipedrive"),
    ])

    conn.commit()


def load_pipedrive_df_from_cache(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query("""
        SELECT
            deal_id,
            title,
            contact_name,
            mailing_address_raw
        FROM pipedrive_deals_cache
    """, conn)

# ============================================
# NORMALIZATION
# ============================================

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

OWNER_REPLACEMENTS = [(re.compile(p, flags=re.IGNORECASE), r) for p, r in owner_changer.items()]
ADDRESS_REPLACEMENTS = [(re.compile(p, flags=re.IGNORECASE), r) for p, r in address_replacements.items()]


def normalize_basic_text(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in {"", "nan", "none", "null"}:
        return ""
    s = s.upper()
    s = re.sub(r"[^A-Z0-9\s,]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def apply_replacements(text: str, replacements) -> str:
    if not text:
        return ""
    for pattern, repl in replacements:
        text = pattern.sub(repl, text)
    text = re.sub(r"[^A-Z0-9\s,]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_owner_text(val) -> str:
    s = normalize_basic_text(val)
    if not s:
        return ""
    return apply_replacements(s, OWNER_REPLACEMENTS)


def clean_address(addr: str) -> str:
    if not addr:
        return ""

    s = normalize_basic_text(addr)
    s = apply_replacements(s, ADDRESS_REPLACEMENTS)

    # remove ZIP codes
    s = re.sub(r"\b\d{5}(?:-\d{4})?\b", "", s)

    # remove trailing USA
    s = re.sub(r"\bUSA\b", "", s)

    s = re.sub(r"\s+", " ", s).strip(" ,")
    return s


def extract_address_city_state(addr: str):
    """
    Best-effort split of mailing address.
    Expected common pattern:
    '123 MAIN ST, HOUSTON, TX 77001, USA'
    """
    cleaned = clean_address(addr)
    if not cleaned:
        return "", "", ""

    parts = [p.strip() for p in cleaned.split(",") if p.strip()]

    if len(parts) >= 3:
        address = parts[0]
        city = parts[-2]
        state = parts[-1]
        state = re.sub(r"\b[A-Z]{2}\s*$", "", state).strip() or parts[-1].split()[-1]
        return address, city, state

    return cleaned, "", ""

def normalize_person_name(val) -> str:
    """
    For person names like Deal - Contact person:
    - keep full name as-is conceptually
    - only normalize spaces / special chars / case
    - do NOT split into parts
    """
    return normalize_owner_text(val)


def build_name_address_key(name, address, city, state) -> Optional[tuple[str, str, str, str]]:
    """
    Key for person-name + mailing-address matching.
    This does NOT apply owner/trust replacements to the name.
    """
    name = normalize_person_name(name)
    address = clean_address(address)
    city = normalize_basic_text(city)
    state = normalize_basic_text(state)

    if not name or not address:
        return None

    return (name, address, city, state)

def build_key(owner, address, city, state) -> Optional[tuple[str, str, str, str]]:
    owner = normalize_owner_text(owner)
    address = clean_address(address)
    city = normalize_basic_text(city)
    state = normalize_basic_text(state)

    if not address:
        return None

    return (owner, address, city, state)

def split_pipe_values(value):
    if value is None:
        return []

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return []

    return [x.strip() for x in s.split("|") if x.strip()]


def normalize_id_value(value):
    if value is None:
        return ""

    s = str(value).strip()

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    return s


def normalize_serial_value(value):
    if value is None:
        return ""

    s = str(value).strip().upper()

    if s.lower() in {"", "nan", "none", "null"}:
        return ""

    # ignore dashes/spaces/special chars
    s = re.sub(r"[^A-Z0-9]", "", s)

    return s

# ============================================
# FETCH FROM PIPEDRIVE
# ============================================

def load_pipedrive_lookup_maps(base_dir: str | Path):
    """
    Load Pipedrive lookup source from local CSV only.
    Also overwrite the local sqlite cache DB from that CSV
    so the cache remains inspectable/debuggable.
    """
    base_dir = Path(base_dir)
    csv_path = base_dir / "cache" / PIPEDRIVE_LOCAL_CACHE_FILENAME

    if not csv_path.exists():
        raise FileNotFoundError(f"Pipedrive local cache CSV not found: {csv_path}")

    df_raw = pd.read_csv(csv_path, dtype=str, keep_default_na=False, na_filter=False).fillna("")
    # Exclude Test Deals from Pipedrive matching
    if "Deal - Label" in df_raw.columns:
        df_raw = df_raw[
            ~df_raw["Deal - Label"]
                .astype(str)
                .str.contains("Test Deal", case=False, na=False)
        ].copy()

    required_cols = {
        "Deal - ID",
        "Deal - Title",
        "Deal - Contact person",
        "Person - Mailing Address",
        "Deal - Serial Number",
    }
    missing = sorted(required_cols - set(df_raw.columns))
    if missing:
        raise ValueError(
            "Pipedrive local cache CSV is missing required columns:\n  - "
            + "\n  - ".join(missing)
        )

    df = pd.DataFrame({
        "deal_id": df_raw["Deal - ID"].astype(str).str.strip(),
        "title": df_raw["Deal - Title"].astype(str).str.strip(),
        "contact_name": df_raw["Deal - Contact person"].astype(str).str.strip(),
        "mailing_address_raw": df_raw["Person - Mailing Address"].astype(str).str.strip(),
        "deal_deo_id": df_raw.get("Deal - DEO ID", pd.Series("", index=df_raw.index)).astype(str).str.strip(),
        "deal_serial_number": df_raw["Deal - Serial Number"].astype(str).str.strip(),
    })

    # overwrite local sqlite cache from CSV
    cache_path = get_cache_path(base_dir)
    conn = connect_cache_db(cache_path)
    try:
        ensure_cache_tables(conn)
        save_pipedrive_df_to_cache(conn, df)
    finally:
        conn.close()

    lookup_maps = build_lookup_maps(df)

    metadata = {
        "rows_loaded": len(df),
        "cache_mode": "rebuilt_from_local_csv",
        "cache_path": str(cache_path),
        "source": str(csv_path),
    }

    return lookup_maps, metadata

def build_lookup_maps(df: pd.DataFrame):
    contact_full_combo = defaultdict(set)
    title_address_rows = []

    deo_id_lookup = defaultdict(set)
    serial_number_lookup = defaultdict(set)

    for row in df.itertuples(index=False):
        if pd.isna(row.deal_id):
            continue

        deal_id = str(row.deal_id).strip()
        if not deal_id:
            continue

        address, city, state = extract_address_city_state(row.mailing_address_raw)

        address_norm = clean_address(address)
        city_norm = normalize_basic_text(city)
        state_norm = normalize_basic_text(state)

        key_contact_full = build_name_address_key(
            row.contact_name,
            address_norm,
            city_norm,
            state_norm,
        )
        if key_contact_full:
            contact_full_combo[key_contact_full].add(deal_id)

        title_norm = normalize_owner_text(row.title)
        if title_norm and address_norm:
            title_address_rows.append({
                "deal_id": deal_id,
                "title_norm": title_norm,
                "address": address_norm,
                "city": city_norm,
                "state": state_norm,
            })

        # Deal - DEO ID lookup
        for deo_id in split_pipe_values(getattr(row, "deal_deo_id", "")):
            deo_id_norm = normalize_id_value(deo_id)
            if deo_id_norm:
                deo_id_lookup[deo_id_norm].add(deal_id)

        # Deal - Serial Number lookup
        for serial in split_pipe_values(getattr(row, "deal_serial_number", "")):
            serial_norm = normalize_serial_value(serial)
            if serial_norm:
                serial_number_lookup[serial_norm].add(deal_id)

    return {
        "contact_full_combo": contact_full_combo,
        "title_address_rows": title_address_rows,
        "deo_id_lookup": deo_id_lookup,
        "serial_number_lookup": serial_number_lookup,
    }


def match_deal_ids(row, lookup_maps):
    matches = set()

    input_address = clean_address(getattr(row, "address", ""))
    input_city = normalize_basic_text(getattr(row, "city", ""))
    input_state = normalize_basic_text(getattr(row, "state", ""))

    lessor_owner = normalize_owner_text(getattr(row, "lessor_owner", ""))
    first_name = normalize_owner_text(getattr(row, "first_name", ""))
    last_name = normalize_owner_text(getattr(row, "last_name", ""))
    first_last_name = " ".join(part for part in [first_name, last_name] if part).strip()

    # --------------------------------------------------
    # 1) Full Deal - Contact person + mailing address
    #    compared against:
    #    - lessor_owner + address/city/state
    #    - first_name + last_name + address/city/state
    # --------------------------------------------------
    key_owner = build_name_address_key(
        lessor_owner,
        input_address,
        input_city,
        input_state,
    )
    if key_owner:
        matches |= lookup_maps["contact_full_combo"].get(key_owner, set())

    key_name = build_name_address_key(
        first_last_name,
        input_address,
        input_city,
        input_state,
    )
    if key_name:
        matches |= lookup_maps["contact_full_combo"].get(key_name, set())

    # --------------------------------------------------
    # 2) Deal title contains name
    #    If lessor_owner or first+last is found in deal title,
    #    then require same address/city/state vs mailing address
    # --------------------------------------------------
    if input_address:
        for rec in lookup_maps["title_address_rows"]:
            if rec["address"] != input_address:
                continue
            if rec["city"] != input_city:
                continue
            if rec["state"] != input_state:
                continue

            if lessor_owner and lessor_owner in rec["title_norm"]:
                matches.add(rec["deal_id"])
                continue

            if first_last_name and first_last_name in rec["title_norm"]:
                matches.add(rec["deal_id"])
                continue

    if not matches:
        return None

    return " | ".join(sorted(matches))


def populate_deal_ids_in_df(df, lookup_maps, progress_callback=None):
    df = df.copy()
    total = len(df)
    results = []

    for i, row in enumerate(df.itertuples(index=False), start=1):
        results.append(match_deal_ids(row, lookup_maps))
        if progress_callback and total:
            progress_callback(i / total)

    df["deal_id"] = results
    return df