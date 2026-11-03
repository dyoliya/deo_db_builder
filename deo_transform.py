import re
import pandas as pd

NULL_LIKE_RAW = {
    "", "#N/A", "N/A", "NA", "NULL", "nan", "#REF",
    "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"
}
NULL_LIKE = {x.strip().casefold() for x in NULL_LIKE_RAW}

def sanitize_colname(name: str) -> str:
    """
    DB-friendly header:
    - normalize dashes to space
    - spaces -> underscore
    - remove non-alphanumeric except underscore
    - collapse multiple underscores
    - strip underscores
    - lowercase
    - avoid starting with digit
    """
    if name is None:
        name = ""

    s = str(name).strip()

    # Normalize dash types to space (prevents double underscores)
    s = s.replace("–", " ").replace("—", " ").replace("-", " ")
    s = s.replace("\u00A0", " ")

    # Spaces -> underscore
    s = re.sub(r"\s+", "_", s)

    # Keep only alnum + underscore
    s = re.sub(r"[^0-9a-zA-Z_]", "", s)

    # Collapse underscores + trim
    s = re.sub(r"_+", "_", s).strip("_")

    s = s.lower()

    if not s:
        s = "col"

    if s[0].isdigit():
        s = f"c_{s}"

    # Allow incoming DEO ID from the file to map to the real SQLite deo_id
    if s == "deo_id":
        return "deo_id"

    return s


def norm(name: str) -> str:
    # Uses sanitize_colname (now defined above ✅)
    return sanitize_colname(name)


EXPLICIT_DATE_COLS = {
    norm("ADD: Address1 First Seen"),
    norm("ADD: Address1 Last Seen"),
    norm("Phone 1 First Seen"),
    norm("Phone 1 Last Seen"),
    norm("Phone 2 First Seen"),
    norm("Phone 2 Last Seen"),
    norm("Phone 3 First Seen"),
    norm("Phone 3 Last Seen"),
    norm("Phone 4 First Seen"),
    norm("Phone 4 Last Seen"),
    norm("Phone 5 First Seen"),
    norm("Phone 5 Last Seen"),
    norm("Email 1 First Seen"),
    norm("Email 1 Last Seen"),
    norm("Email 2 First Seen"),
    norm("Email 2 Last Seen"),
    norm("Email 3 First Seen"),
    norm("Email 3 Last Seen"),
    norm("Email 4 First Seen"),
    norm("Email 4 Last Seen"),
    norm("Email 5 First Seen"),
    norm("Email 5 Last Seen"),
    norm("Report Date"),
}

EXPLICIT_PHONE_COLS = {
    norm("Phone1"),
    norm("Phone2"),
    norm("Phone3"),
    norm("Phone4"),
    norm("CLEANED PHONE1"),
    norm("CLEANED PHONE2"),
    norm("CLEANED PHONE3"),
    norm("CLEANED PHONE4"),
    norm("Phone 1"),
    norm("Phone 2"),
    norm("Phone 3"),
    norm("Phone 4"),
    norm("Phone 5"),
}


def make_unique(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        if n not in seen:
            seen[n] = 0
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
    return out


def coerce_nulls(x):
    if x is None:
        return None
    s = str(x).strip()
    if s.casefold() in NULL_LIKE:
        return None
    return s

def normalize_phone_value(x):
    """Keep phone values as TEXT and avoid float artifacts like 12345.0, preserve NULLs."""
    if x is None:
        return None
    s = str(x).strip()
    if s.casefold() in NULL_LIKE:
        return None

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    if re.search(r"[eE]", s):
        try:
            s = str(int(float(s)))
        except Exception:
            pass

    return s


def standardize_date_series(series: pd.Series) -> pd.Series:
    """Convert parseable values to YYYY-MM-DD; keep original if parse fails."""
    cleaned = series.map(coerce_nulls)
    dt = pd.to_datetime(cleaned, errors="coerce")

    out = []
    for orig, parsed in zip(cleaned.tolist(), dt.tolist()):
        if orig is None:
            out.append(None)
        elif parsed is pd.NaT or parsed is None:
            out.append(orig)
        else:
            out.append(parsed.date().strftime("%Y-%m-%d"))
    return pd.Series(out, index=series.index)


def clean_and_prepare_df(df_raw: pd.DataFrame):
    ignored_cols = {"assigned_to", "date", "date_added_to_db"}

    # Decide which input columns to keep.
    keep_cols = [
        sanitize_colname(c) not in ignored_cols
        for c in df_raw.columns
    ]

    # Copy only the columns we want to import.
    df = df_raw.loc[:, keep_cols].copy(deep=True)

    original_cols = list(df.columns)
    sanitized = make_unique([sanitize_colname(c) for c in original_cols])
    df.columns = sanitized

    audit_records = []

    def _as_text_series(s: pd.Series) -> pd.Series:
        return s.map(lambda x: "" if pd.isna(x) else str(x))

    def audit_stage(stage: str, before_df: pd.DataFrame, after_df: pd.DataFrame, cols: list[str], reason: str):
        for c in cols:
            if c not in before_df.columns or c not in after_df.columns:
                continue

            before_col = _as_text_series(before_df[c])
            after_col = after_df[c]

            # suspicious: had something (non-empty) then became NULL/None
            after_missing = after_df[c].map(lambda v: v is None or pd.isna(v))
            mask = (before_col.str.strip() != "") & (after_missing)

            if mask.any():
                for idx in after_df.index[mask]:
                    audit_records.append({
                        "row_index": int(idx) + 2,  # assumes header is row 1
                        "column_sanitized": c,
                        "original_value": before_col.loc[idx],
                        "cleaned_value": None,
                        "stage": stage,
                        "reason": reason,
                    })

    # ---------------- Stage 1: NULL-like -> None ----------------
    before1 = df.copy(deep=True)
    for c in df.columns:
        df[c] = df[c].map(coerce_nulls)

    audit_stage(
        stage="coerce_nulls",
        before_df=before1,
        after_df=df,
        cols=list(df.columns),
        reason="Non-empty value became NULL via NULL_LIKE/coerce_nulls",
    )

    # ---------------- Stage 2: phones ----------------
    phone_cols = [c for c in df.columns if c in EXPLICIT_PHONE_COLS]
    if phone_cols:
        before2 = df.copy(deep=True)
        for c in phone_cols:
            df[c] = df[c].map(normalize_phone_value)

        audit_stage(
            stage="normalize_phone_value",
            before_df=before2,
            after_df=df,
            cols=phone_cols,
            reason="Non-empty value became NULL during phone normalization",
        )

    # ---------------- Stage 3: dates ----------------
    date_cols = [c for c in df.columns if c in EXPLICIT_DATE_COLS]
    if date_cols:
        before3 = df.copy(deep=True)
        for c in date_cols:
            df[c] = standardize_date_series(df[c])

        audit_stage(
            stage="standardize_date_series",
            before_df=before3,
            after_df=df,
            cols=date_cols,
            reason="Non-empty value became NULL during date standardization",
        )

    audit_df = pd.DataFrame(audit_records)
    return df, original_cols, sanitized, phone_cols, date_cols, audit_df