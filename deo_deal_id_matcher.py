# deo_deal_id_matcher.py

import pandas as pd
from pipedrive_lookup import (
    load_pipedrive_lookup_maps,
    build_name_address_key,
    clean_address,
    normalize_basic_text,
    normalize_owner_text,
    normalize_id_value,
    normalize_serial_value,
)

DEO_ADDRESS_SETS = [
    ("input_address", "input_city", "input_state"),
    # ("address2", "city2", "state2"),
    # ("add_address1", "add_address1_city", "add_address1_state"),
]


def get_deo_names(row):
    owner_name = normalize_owner_text(getattr(row, "owner", ""))
    owner_std_name = normalize_owner_text(getattr(row, "owner_standardized", ""))
    first_name = normalize_owner_text(getattr(row, "first_name", ""))
    last_name = normalize_owner_text(getattr(row, "last_name", ""))
    first_last_name = " ".join(
        part for part in [first_name, last_name] if part
    ).strip()

    return [name for name in [owner_name, first_last_name] if name]


def get_deo_address_sets(row):
    address_sets = []

    for address_col, city_col, state_col in DEO_ADDRESS_SETS:
        address = clean_address(getattr(row, address_col, ""))
        city = normalize_basic_text(getattr(row, city_col, ""))
        state = normalize_basic_text(getattr(row, state_col, ""))

        if address:
            address_sets.append((address, city, state))

    return address_sets


def match_deo_deal_ids(row, lookup_maps):
    match_methods = {}

    def add_matches(deal_ids, method):
        for deal_id in deal_ids:
            match_methods.setdefault(deal_id, set()).add(method)

    # --------------------------------------------------
    # 1) DEO ID match
    # --------------------------------------------------
    deo_id = normalize_id_value(getattr(row, "deo_id", ""))

    if deo_id:
        deo_matches = lookup_maps["deo_id_lookup"].get(deo_id, set())
        add_matches(deo_matches, "DEO ID")

    # --------------------------------------------------
    # 2) Serial number match
    # --------------------------------------------------
    deo_serial_number = normalize_serial_value(getattr(row, "deo_serial_number", ""))

    if deo_serial_number:
        serial_matches = lookup_maps["serial_number_lookup"].get(deo_serial_number, set())
        add_matches(serial_matches, "Serial Number")

    # --------------------------------------------------
    # 3) Deal - Contact person + Person - Mailing Address
    # --------------------------------------------------
    names = get_deo_names(row)
    address_sets = get_deo_address_sets(row)

    for name in names:
        for address, city, state in address_sets:
            key = build_name_address_key(name, address, city, state)
            if key:
                name_address_matches = lookup_maps["contact_full_combo"].get(key, set())
                add_matches(name_address_matches, "Name + Address")

    # --------------------------------------------------
    # 4) Deal - Title contains name + same mailing address/city/state
    # --------------------------------------------------
    for address, city, state in address_sets:
        for rec in lookup_maps["title_address_rows"]:
            if rec["address"] != address:
                continue
            if rec["city"] != city:
                continue
            if rec["state"] != state:
                continue

            for name in names:
                if name and name in rec["title_norm"]:
                    add_matches({rec["deal_id"]}, "Title + Address")

    if not match_methods:
        return {
            "deal_id": None,
            "deal_id_match_method": None,
            "deal_id_matched_by_deo_id": None,
            "deal_id_matched_by_deo_serial_number": None,
            "deal_id_matched_by_name_address": None,
            "deal_id_matched_by_title_address": None,
        }

    deal_ids = sorted(match_methods.keys())

    return {
        "deal_id": " | ".join(deal_ids),
        "deal_id_match_method": " | ".join(
            f"{deal_id}: {', '.join(sorted(methods))}"
            for deal_id, methods in sorted(match_methods.items())
        ),
        "deal_id_matched_by_deo_id": " | ".join(
            deal_id for deal_id in deal_ids
            if "DEO ID" in match_methods[deal_id]
        ) or None,
        "deal_id_matched_by_deo_serial_number": " | ".join(
            deal_id for deal_id in deal_ids
            if "Serial Number" in match_methods[deal_id]
        ) or None,
        "deal_id_matched_by_name_address": " | ".join(
            deal_id for deal_id in deal_ids
            if "Name + Address" in match_methods[deal_id]
        ) or None,
        "deal_id_matched_by_title_address": " | ".join(
            deal_id for deal_id in deal_ids
            if "Title + Address" in match_methods[deal_id]
        ) or None,
    }


def populate_deo_deal_ids_in_df(df, base_dir, progress_callback=None):
    lookup_maps, metadata = load_pipedrive_lookup_maps(base_dir)

    df = df.copy()
    total = len(df)

    deal_id_results = []
    audit_rows = []

    for i, row in enumerate(df.itertuples(index=False), start=1):
        result = match_deo_deal_ids(row, lookup_maps)

        deal_id_results.append(result["deal_id"])

        if result["deal_id"]:
            audit_rows.append({
                "deo_id": getattr(row, "deo_id", ""),
                "deo_serial_number": getattr(row, "deo_serial_number", ""),
                "deal_id": result["deal_id"],
                "deal_id_match_method": result["deal_id_match_method"],
                "deal_id_matched_by_deo_id": result["deal_id_matched_by_deo_id"],
                "deal_id_matched_by_deo_serial_number": result["deal_id_matched_by_deo_serial_number"],
                "deal_id_matched_by_name_address": result["deal_id_matched_by_name_address"],
                "deal_id_matched_by_title_address": result["deal_id_matched_by_title_address"],
            })

        if progress_callback and total:
            progress_callback(i / total)

    # only this column goes to DB
    df["deal_id"] = deal_id_results

    # separate audit only
    audit_df = pd.DataFrame(audit_rows)

    return df, metadata, audit_df