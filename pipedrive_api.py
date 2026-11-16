import json
import time
import tempfile
from pathlib import Path

import pandas as pd
import requests

# ============================================
# CONFIG
# ============================================

SETTINGS_FILENAME = "pipedrive_settings.json"
CACHE_DIRNAME = "cache"
CSV_FILENAME = "pipedrive_local_cache.csv"

DEO_ID_COLUMN = "Deal - DEO ID"

# Must match the real Pipedrive export header exactly (order doesn't matter
# functionally, but keeping it matches cache/pipedrive_local_cache.manual_baseline.csv).
CSV_COLUMNS = [
    "Deal - ID",
    "Deal - Title",
    "Deal - County",
    "Deal - Serial Number",
    "Deal - Contact Group ID",
    "Deal - Unique Database ID",
    "Deal - BU Database ID",
    "Deal - NCOP ID",
    "Deal - Pooling Order ID",
    "Deal - Contact person",
    "Person - Mailing Address",
    "Person - Phone - Work",
    "Person - Phone - Home",
    "Person - Phone - Mobile",
    "Person - Phone - Other",
    "Person - Email - Work",
    "Person - Email - Home",
    "Person - Email - Other",
    "Deal - Pipeline",
    "Deal - Stage",
    "Deal - Label",
    "Deal - Offer Generated Date",
    DEO_ID_COLUMN,
]

PERSON_PHONE_LABELS = ("Work", "Home", "Mobile", "Other")
PERSON_EMAIL_LABELS = ("Home", "Work", "Other")


class PipedriveError(RuntimeError):
    pass


# ============================================
# SETTINGS HELPERS
# ============================================

def _settings_path(base_dir: str | Path) -> Path:
    return Path(base_dir) / "config" / SETTINGS_FILENAME


def load_settings(base_dir: str | Path) -> dict:
    path = _settings_path(base_dir)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_settings(base_dir: str | Path, **updates) -> dict:
    path = _settings_path(base_dir)
    settings = {}
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            settings = json.load(f)
    settings.update(updates)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)
    return settings


def cache_csv_path(base_dir: str | Path) -> Path:
    return Path(base_dir) / CACHE_DIRNAME / CSV_FILENAME


def check_cache_write_access(base_dir: str | Path) -> None:
    path = cache_csv_path(base_dir)

    try:
        path.parent.mkdir(parents=True, exist_ok=True)

        if path.exists():
            # Test write access without changing the contents or timestamp.
            # Do not use "w": that would erase the existing cache.
            with path.open("r+b"):
                pass

        # Also verify that the cache folder allows file creation.
        with tempfile.TemporaryFile(dir=path.parent):
            pass

    except OSError as exc:
        raise PipedriveError(
            "Cannot write to the Pipedrive cache CSV or its folder.\n\n"
            f"File: {path}\n\n"
            "Close this CSV in Excel or any other application, "
            "then click Check references again.\n"
            "If it is already closed, check whether the file or folder "
            "is read-only or you lack write permission.\n\n"
            f"Details: {exc}"
        ) from exc


def cache_status(base_dir: str | Path) -> dict:
    path = cache_csv_path(base_dir)
    if not path.exists():
        return {"exists": False}

    stat = path.stat()
    rows = None
    try:
        with open(path, "r", encoding="utf-8") as f:
            rows = sum(1 for _ in f) - 1
    except OSError:
        pass

    return {
        "exists": True,
        "mtime": stat.st_mtime,
        "age_seconds": time.time() - stat.st_mtime,
        "rows": rows,
    }


# ============================================
# PIPEDRIVE API CLIENT
# ============================================

class Pipedrive:
    def __init__(self, token: str, base_url: str):
        self.token = token
        self.root = base_url.rstrip("/") + "/api/v2/"
        self.session = requests.Session()

    def get(self, endpoint: str, params: dict) -> dict:
        url = self.root + endpoint
        headers = {"x-api-token": self.token, "Accept": "application/json"}

        for attempt in range(5):
            try:
                response = self.session.get(url, params=params, headers=headers, timeout=90)
            except requests.exceptions.RequestException:
                if attempt == 4:
                    raise PipedriveError(
                        "Pipedrive connection failed. Check your internet connection and try again."
                    ) from None
                time.sleep(2 ** attempt)
                continue

            if response.status_code in (429, 500, 502, 503, 504):
                if attempt == 4:
                    raise PipedriveError("Pipedrive is busy or its request limit was reached. Try again later.")
                try:
                    delay = max(1, float(response.headers.get("Retry-After", 2 ** attempt)))
                except ValueError:
                    delay = 2 ** attempt
                if delay > 60:
                    raise PipedriveError("Pipedrive requested a longer wait. Try again in a few minutes.")
                time.sleep(delay)
                continue

            if response.status_code != 200:
                raise PipedriveError(
                    f"Pipedrive returned HTTP {response.status_code}. "
                    "Check your API token and access to the configured filter."
                )

            try:
                data = response.json()
            except ValueError:
                raise PipedriveError("Pipedrive returned an unreadable response.") from None

            if data.get("success") is not True:
                raise PipedriveError("Pipedrive could not complete this request.")

            return data

        raise PipedriveError("Pipedrive request failed.")

    def all(self, endpoint: str, log=None, **params) -> list:
        rows = []
        seen = set()
        params = {k: v for k, v in params.items() if v is not None}
        params = dict(params, limit=500)

        while True:
            payload = self.get(endpoint, params)
            page = payload.get("data") or []
            if not isinstance(page, list):
                raise PipedriveError("Unexpected Pipedrive response structure.")
            rows.extend(page)
            if log:
                log(f"[PIPEDRIVE] {endpoint}: {len(rows):,} records downloaded...\n")

            cursor = (payload.get("additional_data") or {}).get("next_cursor")
            if not cursor:
                break
            if cursor in seen:
                raise PipedriveError("Pipedrive pagination repeated a page; download stopped.")
            seen.add(cursor)
            params["cursor"] = cursor

        return rows

    def field_name_map(self, endpoint: str) -> dict:
        """Map casefolded field display-name -> list of {key, options} (list, to detect ambiguity)."""
        by_name = {}
        for field in self.all(endpoint):
            name = str(field.get("field_name") or "").strip()
            key = field.get("field_code")
            if not name or not key:
                continue

            options = {}
            for opt in field.get("options") or []:
                opt_id = opt.get("id")
                label = opt.get("label")
                if opt_id is not None and label is not None:
                    options[str(opt_id)] = str(label)

            by_name.setdefault(name.casefold(), []).append({"key": key, "options": options})
        return by_name

    def resolve_field(self, by_name: dict, display_name: str, required: bool = True):
        matches = by_name.get(str(display_name).strip().casefold(), [])
        if len(matches) == 1:
            return matches[0]
        if not matches:
            if required:
                raise PipedriveError(
                    f'Pipedrive field "{display_name}" was not found. '
                    "Check the field name in config/pipedrive_settings.json or in Pipedrive."
                )
            return None
        raise PipedriveError(
            f'Pipedrive field "{display_name}" is ambiguous '
            f"({len(matches)} fields share this name)."
        )


# ============================================
# VALUE RESOLUTION
# ============================================

def _unwrap(value):
    if isinstance(value, dict):
        return value.get("value", "")
    return value


def _lookup_id(value):
    if isinstance(value, dict):
        return value.get("value")
    return value


def _resolve_options(value, options_map: dict) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, list):
        raw_ids = value
    else:
        text = str(value)
        raw_ids = [p.strip() for p in text.split(",") if p.strip()] if "," in text else [text]
    labels = [options_map.get(str(raw_id), str(raw_id)) for raw_id in raw_ids]
    return ", ".join(labels)


def custom_value(item: dict, field_info) -> str:
    if field_info is None:
        return ""

    key = field_info["key"]
    fields = item.get("custom_fields") or {}
    if key in fields:
        value = fields[key]
    elif key in item:
        value = item[key]
    else:
        return ""

    value = _unwrap(value)
    if value is None:
        return ""

    options = field_info.get("options")
    if options:
        return _resolve_options(value, options)

    return str(value)


def build_person_row(person, address_field) -> dict:
    row = {
        "Deal - Contact person": "",
        "Person - Mailing Address": "",
        "Person - Phone - Work": "",
        "Person - Phone - Home": "",
        "Person - Phone - Mobile": "",
        "Person - Phone - Other": "",
        "Person - Email - Work": "",
        "Person - Email - Home": "",
        "Person - Email - Other": "",
    }
    if person is None:
        return row

    row["Deal - Contact person"] = str(person.get("name") or "")
    row["Person - Mailing Address"] = custom_value(person, address_field)

    for number in person.get("phones") or []:
        label = str(number.get("label") or "Other").strip().title()
        if label not in PERSON_PHONE_LABELS:
            label = "Other"
        col = f"Person - Phone - {label}"
        value = str(number.get("value") or "")
        row[col] = ", ".join(filter(None, [row[col], value]))

    for email in person.get("emails") or []:
        label = str(email.get("label") or "Other").strip().title()
        if label not in PERSON_EMAIL_LABELS:
            label = "Other"
        col = f"Person - Email - {label}"
        value = str(email.get("value") or "")
        row[col] = ", ".join(filter(None, [row[col], value]))

    return row


# ============================================
# DOWNLOAD
# ============================================

def download_deals(base_dir: str | Path, settings: dict, token: str, log=None, progress=None) -> pd.DataFrame:
    def _log(msg):
        if log:
            log(msg)

    def _progress(frac):
        if progress:
            progress(frac)

    base_dir = Path(base_dir)

    # Recheck immediately before starting any API requests.
    check_cache_write_access(base_dir)

    base_url = settings["base_url"]
    filter_id = settings["deals_filter_id"]
    deal_custom_field_labels = settings.get("deal_custom_field_labels") or {}

    api = Pipedrive(token, base_url)
    try:
        _log("[PIPEDRIVE] Resolving custom field names...\n")
        deal_field_map = api.field_name_map("dealFields")
        person_field_map = api.field_name_map("personFields")

        custom_field_info = {}
        for csv_column, pipedrive_name in deal_custom_field_labels.items():
            required = csv_column != DEO_ID_COLUMN
            info = api.resolve_field(deal_field_map, pipedrive_name, required=required)
            custom_field_info[csv_column] = info
            if info is None:
                _log(
                    f'[PIPEDRIVE] "{pipedrive_name}" field not found in Pipedrive yet; '
                    f'"{csv_column}" will be blank for every deal.\n'
                )

        address_field = api.resolve_field(person_field_map, "Mailing Address", required=True)

        _log("[PIPEDRIVE] Downloading pipelines and stages...\n")
        pipelines = {str(p["id"]): str(p.get("name") or "") for p in api.all("pipelines", log=_log)}
        stages = {str(s["id"]): str(s.get("name") or "") for s in api.all("stages", log=_log)}

        _log(f"[PIPEDRIVE] Downloading deals (filter {filter_id})...\n")
        # Pipedrive v2 returns the full custom_fields object (plus built-in-style custom
        # fields like Label as top-level keys) regardless of the custom_fields param, and
        # rejects some valid field codes (e.g. "label_ids") when passed to it, so the
        # narrowing param is skipped entirely rather than worked around.
        deals = api.all("deals", filter_id=filter_id, log=_log)
        if not deals:
            raise PipedriveError(
                f"Pipedrive filter {filter_id} returned no deals. Confirm the filter ID and your access to it."
            )
        _log(f"[PIPEDRIVE] {len(deals):,} deals downloaded.\n")

        person_ids = sorted({
            int(_lookup_id(deal["person_id"]))
            for deal in deals
            if deal.get("person_id") is not None
        })

        people = {}
        total_persons = len(person_ids)

        _log(
            f"[PIPEDRIVE] Downloading {total_persons:,} linked contacts...\n"
        )

        for start in range(0, total_persons, 100):
            batch = person_ids[start:start + 100]

            # Suppress the per-batch counter.
            fetched = api.all(
                "persons",
                ids=",".join(str(p) for p in batch),
                log=None,
            )

            for person in fetched:
                people[int(person["id"])] = person

            # Count unique persons actually received across all batches.
            _log(
                f"[PIPEDRIVE] Linked contacts: "
                f"{len(people):,} of {total_persons:,} downloaded...\n"
            )

        _log(
            f"[PIPEDRIVE] Linked contacts download finished: "
            f"{len(people):,} received out of "
            f"{total_persons:,} requested.\n"
        )

        rows = []
        total = len(deals)
        for i, deal in enumerate(deals, start=1):
            row = {
                "Deal - ID": str(deal["id"]),
                "Deal - Title": str(deal.get("title") or ""),
            }
            for csv_column, info in custom_field_info.items():
                row[csv_column] = custom_value(deal, info)

            pipeline_id = _lookup_id(deal.get("pipeline_id"))
            stage_id = _lookup_id(deal.get("stage_id"))
            row["Deal - Pipeline"] = pipelines.get(str(pipeline_id), "") if pipeline_id is not None else ""
            row["Deal - Stage"] = stages.get(str(stage_id), "") if stage_id is not None else ""

            person_id = deal.get("person_id")
            person = people.get(int(_lookup_id(person_id))) if person_id is not None else None
            row.update(build_person_row(person, address_field))

            rows.append(row)
            if total:
                _progress(i / total)

        df = pd.DataFrame(rows)
        for column in CSV_COLUMNS:
            if column not in df.columns:
                df[column] = ""
        df = df[CSV_COLUMNS]

        cache_dir = base_dir / CACHE_DIRNAME
        cache_dir.mkdir(parents=True, exist_ok=True)
        csv_path = cache_dir / CSV_FILENAME
        df.to_csv(csv_path, index=False, encoding="utf-8")

        _log(f"[PIPEDRIVE] Saved {len(df):,} deals to {csv_path}\n")
        return df
    finally:
        api.token = ""


def ensure_fresh_cache(
    base_dir: str | Path,
    log=None,
    progress=None,
) -> dict:
    base_dir = Path(base_dir)
    csv_path = cache_csv_path(base_dir)

    # Fail before cache reuse or downloading if the CSV cannot be written.
    check_cache_write_access(base_dir)

    # A valid cache can be reused without making an API request.
    if csv_path.is_file():
        age = time.time() - csv_path.stat().st_mtime

        if 0 <= age < 8 * 60 * 60:
            if log:
                log(
                    "[PIPEDRIVE] Using cached deals "
                    f"({age / 3600:.1f}h old)."
                )

            return {
                "source": "cache",
                "downloaded": False,
                "age_seconds": age,
            }

    # Missing, expired, or future-dated caches require a download.
    settings = load_settings(base_dir)
    token = str(settings.get("api_token") or "").strip()

    if not token:
        raise PipedriveError(
            "Pipedrive cache is missing or expired. "
            "Enter your API token and click Save API token."
        )

    if not settings.get("base_url") or not settings.get("deals_filter_id"):
        raise PipedriveError(
            "Set base_url and deals_filter_id in "
            "config/pipedrive_settings.json."
        )

    if log:
        log(
            "[PIPEDRIVE] Cache missing or expired. "
            "Downloading fresh deals..."
        )

    download_deals(
        base_dir,
        settings,
        token,
        log=log,
        progress=progress,
    )

    return {
        "source": "api",
        "downloaded": True,
    }
