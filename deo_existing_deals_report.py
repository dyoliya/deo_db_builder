# deo_existing_deals_report.py

import json
import sqlite3
from pathlib import Path
import pandas as pd


PIPEDRIVE_LOCAL_CACHE_FILENAME = "pipedrive_local_cache.csv"
OPERATOR_PARTNERSHIP_PIPELINE = "Operator Partnership Pipeline"


def split_deal_ids(value):
    if value is None:
        return []

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return []

    return [x.strip() for x in s.split("|") if x.strip()]


def _load_pipedrive_base_url(base_dir: Path) -> str:
    config_path = Path(base_dir) / "config" / "pipedrive_settings.json"
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    return config["base_url"]


def generate_deo_associated_deals_report(db_path, base_dir, table="deo"):
    db_path = Path(db_path)
    base_dir = Path(base_dir)
    pipedrive_base_url = _load_pipedrive_base_url(base_dir)

    cache_path = base_dir / "cache" / PIPEDRIVE_LOCAL_CACHE_FILENAME
    if not cache_path.exists():
        raise FileNotFoundError(f"Pipedrive local cache CSV not found: {cache_path}")

    pd_df = pd.read_csv(
        cache_path,
        dtype=str,
        keep_default_na=False,
        na_filter=False
    ).fillna("")

    required_pd_cols = {
        "Deal - ID",
        "Deal - Pipeline",
        "Deal - Stage",
    }

    missing_pd_cols = sorted(required_pd_cols - set(pd_df.columns))
    if missing_pd_cols:
        raise ValueError(
            "Pipedrive local cache CSV is missing required columns:\n  - "
            + "\n  - ".join(missing_pd_cols)
        )

    pd_lookup = {}

    for _, row in pd_df.iterrows():
        deal_id = str(row["Deal - ID"]).strip()
        if not deal_id:
            continue

        pipeline = str(row["Deal - Pipeline"]).strip()
        stage = str(row["Deal - Stage"]).strip()

        pd_lookup[deal_id] = {
            "deal_pipeline": pipeline,
            "deal_stage": stage,
            "is_operator_partnership_pipeline": (
                "Y" if pipeline == OPERATOR_PARTNERSHIP_PIPELINE else "N"
            ),
        }

    conn = sqlite3.connect(db_path)
    try:
        deo_df = pd.read_sql_query(
            f"""
            SELECT
                deo_id,
                deal_id
            FROM {table}
            WHERE deal_id IS NOT NULL
              AND TRIM(deal_id) <> ''
            """,
            conn
        )
    finally:
        conn.close()

    output_rows = []

    for _, row in deo_df.iterrows():
        deo_id = row["deo_id"]
        deal_ids = split_deal_ids(row["deal_id"])

        for deal_id in deal_ids:
            deal_info = pd_lookup.get(deal_id, {})

            output_rows.append({
                "deo_id": deo_id,
                "deal_id": deal_id,
                "deal_link": f"{pipedrive_base_url}/deal/{deal_id}",
                "Deal - Pipeline": deal_info.get("deal_pipeline", ""),
                "Deal - Stage": deal_info.get("deal_stage", ""),
                "Operator Partnership Pipeline": deal_info.get(
                    "is_operator_partnership_pipeline",
                    "N"
                ),
            })

    out_df = pd.DataFrame(output_rows)

    output_dir = base_dir / "existing_deals"
    output_dir.mkdir(parents=True, exist_ok=True)

    # db filename example: 2026-05-18-deo.db
    date_prefix = db_path.name[:10]
    output_path = output_dir / f"{date_prefix}-deo-associated-deals.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="associated_deals", index=False)

    return {
        "output_path": str(output_path),
        "rows_exported": len(out_df),
        "unique_deo_rows": len(deo_df),
    }