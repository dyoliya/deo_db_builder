# DEO DB Builder

**DEO DB Builder** is a desktop utility (built with Python) that implements a
**controlled Extract–Transform–Load (ETL) ingestion workflow** for Deceased Owner (DEO) records.

It ingests deceased-owner lead lists from a shared Google Sheet, cleans and standardizes them, matches each record against the company's Pipedrive deals and BUDB (Bottoms Up Database) mineral-ownership records, and persists the result into a versioned, daily SQLite database for long-term retention and reporting.

---

![Version](https://img.shields.io/badge/version-2.0.0-ffab4c?style=for-the-badge&logo=python&logoColor=white)
![Python](https://img.shields.io/badge/python-3.10%2B-273946?style=for-the-badge&logo=python&logoColor=ffab4c)
![Status](https://img.shields.io/badge/status-active-273946?style=for-the-badge&logo=github&logoColor=ffab4c)

---

## 🚧 Problem Statement / Motivation

Deceased-owner lead lists arrive via a shared Google Sheet that multiple team members contribute to, with inconsistent formatting, duplicate entries, and no link back to the company's existing systems of record (Pipedrive deals, BUDB ownership data).

Manually cross-referencing each incoming owner against Pipedrive and BUDB to check "do we already have a deal or ownership record for this person" is slow and error-prone at volume, and without a durable, versioned database there is no reliable way to track what changed between imports or recover from a bad run.

To close this gap, a repeatable ingestion pipeline is required — one that standardizes incoming data, automatically links it to existing Pipedrive deals and BUDB records, and stores the result in a daily, auditable database.

DEO DB Builder addresses this by enforcing a controlled ETL pipeline:
1. Extract deceased-owner records from the shared Google Sheet.
2. Clean and standardize columns, phone numbers, and dates, flagging any suspicious data loss for audit.
3. Match each record to a Pipedrive deal ID and a BUDB ID using cached reference data.
4. Sync the result into a daily-versioned SQLite database, tracking inserts/updates/deletes.
5. Write matched IDs back to the Google Sheet (when that was the source) and generate audit/report files for review.

---

## ✨ Features
- **Google Sheets ingestion**: reads live data from a shared Google Sheet (tab must be named `deo_db`) via a Google service account; each download is also backed up locally under `input/google_sheet_backups/` (one subfolder per spreadsheet ID) for traceability. (Local CSV/XLSX import existed in an earlier version but has been removed from the current UI — `read_input_file()` in `app.py` remains as leftover support code.)
- **Data cleaning**: column name sanitization, phone number and date normalization, and a NULL-conversion audit for anything that looks like unintended data loss.
- **Pipedrive deal matching** against a Pipedrive deals cache, downloaded automatically via the Pipedrive API (owner name + mailing address matching), with a SQLite-backed lookup cache for speed.
- **BUDB ID matching** against a BUDB reference database **you select each session** on the **BUDB** tab (owner name + address matching), with its own signature-based lookup cache that rebuilds only when the selected file changes.
  - **BUDB contact-group auto-resolution**: if a matched owner's BUDB IDs span more than one distinct `contact_group_id`, the app automatically keeps only the IDs belonging to the lowest-numbered group and records the full decision (all groups seen, which IDs were kept/removed) in a `budb_group_selection` audit sheet for review.
- **County enrichment** by combining matched county data from all linked sources.
- **Pre-flight "Check references"**: validates Google Sheet access/permissions and the Pipedrive API token before enabling the build button, so a bad credential or sheet link is caught before any database changes are made.
- **Tracked, crash-safe database builds** (`deo_build_state.py`):
  - Auto-creates/copies forward a new dated DB each day from the most recent prior one (same-day re-runs continue syncing into that day's file).
  - Every build attempt first archives a consistent snapshot of the current database into `database/previous_versions/<build_id>/`, before any changes are made.
  - If a build fails partway through (and no Google writeback has started yet), the app automatically restores the pre-build snapshot so the active database is never left in a half-written state.
  - Syncs rows by `deo_id`, reporting inserted / updated / unchanged / deleted counts.
  - Build status (`building` / `ready_for_log` / `complete` / `rolled_back` / `recovery_required`) is tracked both in `config/build_state.json` and inside the database itself, so the UI's **Build DB** / **Update Action Log** buttons stay in sync with what actually happened.
- **Google Sheets writeback**: after a successful import, writes deal IDs, serial numbers, and upload dates back to the source sheet, with a pending/completed record so an interrupted writeback can be recovered instead of silently lost.
- **Integrated action log**: after a build completes, clicking **Update Action Log** diffs the before/after database snapshots and records every schema change, added row, deleted row, and modified field into a persistent `deo_action_log` table — no manual file copying required.
- **Baseline snapshot / test-run safety net** (`initialize_deo_baseline.py`): a manually-run, standalone script that stamps a known-good database as the official tracked baseline and backs it up to `database/previous_versions/`, so a confirmed-stable version can be restored if a test run needs to be thrown away. See **Baseline Snapshots & Test Runs** below.
- **Audit & reporting output**: a combined NULL-conversion + BUDB group-selection audit workbook, a deal-match audit, and a DEO-associated-deals report are generated automatically after each import.
- **Operator-friendly desktop UI** with file/sheet selection, progress bar, and a live activity log.
---

## 🧠 Logic Flow

### Import / build flow
1. User clicks **Check references** — validates the Google Sheet URL/permissions (if used) and the Pipedrive API token without writing anything. The **Build DB** button stays disabled until this passes.
2. User pastes the Google Sheet URL (its `deo_db` tab is read) and clicks **Build DB**.
3. `prepare_build()` (`deo_build_state.py`) resolves today's database (`database/YYYY-MM-DD-deo.db`):
   - Archives a consistent snapshot of whatever database currently exists into `database/previous_versions/<build_id>/`, **before** any changes are made.
   - If today's file doesn't exist yet, copies it forward from that snapshot (or starts a brand-new DB if none exists).
   - Records the build as `building` in `config/build_state.json` and inside the database itself.
4. For the Google Sheet:
   - Downloads the `deo_db` tab, verifies the sheet hasn't changed mid-download, skips fully blank rows, and saves a timestamped local backup copy to `input/google_sheet_backups/<spreadsheet_id>/`.
   - Cleans and standardizes columns, phone numbers, and dates; records any suspicious NULL conversions for audit.
   - Matches each row to a **Pipedrive deal ID** using the Pipedrive cache (`cache/pipedrive_local_cache.csv`, refreshed automatically via the Pipedrive API).
   - Matches each row to a **BUDB ID** using the BUDB `.db` file selected on the **BUDB** tab; if the match spans multiple `contact_group_id` values, auto-resolves to the lowest group and logs the decision.
   - Combines county information from all matched sources.
5. Validates incoming columns against the existing DB schema — blocks unexpected columns, allows known-missing ones.
6. Syncs rows into the daily table keyed by `deo_id`, assigning `deo_serial_number` and `date_uploaded` as needed, and reports inserted/updated/unchanged/deleted counts.
7. Writes deal IDs, serial numbers, and upload dates back to the sheet, tracking a pending → completed writeback record so a failed writeback can be resumed safely.
8. On success, `finish_build()` marks the build `ready_for_log` and retires the prior active file (its pre-build copy already lives in `database/previous_versions/`). On failure, `fail_build()` automatically restores the pre-build snapshot (unless a Google writeback had already started, in which case it's left as `recovery_required` for manual review) — see **Troubleshooting** below.
9. Saves output artifacts:
   - `audit/` — a combined NULL-conversion + BUDB group-selection audit workbook (only if either kind of issue was found).
   - `existing_deals/` — deal-match audit and the DEO-associated-deals report.

### Action-log flow (`deo_action_log.py`)
This runs **inside `app.py`** — the **Update Action Log** button becomes enabled once a build reaches `ready_for_log`, and clicking it diffs the build's recorded before/after snapshots automatically, with no manual file copying needed.

1. After **Build DB** completes, click **Update Action Log**. The app calls `compare_databases(old_db_path, new_db_path, build_id)` using the paths recorded by the build, then marks the build `complete`. A `deo_action_log_runs` table (keyed by `build_id`) prevents the same build from being logged twice.
2. The comparison:
   - Reads the date from the new DB's filename (expects a `YYYY-MM-DD` pattern, e.g. `2026-03-12-deo.db`).
   - Creates the `deo_action_log` table inside the new DB if it doesn't already exist.
   - If the new DB's log table is still empty, copies forward all historical log rows from the old DB's log table first — this is what carries the full change history along as you roll from one daily DB to the next.
   - Auto-detects the shared data table between old and new (e.g. `deo`) and compares them by `deo_id`:
     - Logs any columns added or removed between the two schemas.
     - Logs any `deo_id` present only in the new DB as `added`, and any present only in the old DB as `deleted` (full row snapshot as JSON).
     - For rows present in both, logs each changed field individually as `modified` (old value → new value). `date_uploaded` is intentionally excluded from field comparison since it changes on every sync.
3. The resulting `deo_action_log` table — now living inside the new DB file — becomes the historical record going forward.
---

## 📝 Requirements
- **Python**: 3.10+ recommended.
- **Services**:
  - Google Cloud service account with Sheets (read/write) and Drive (read-only) scopes — required, since the Google Sheet is the only supported import source.
- **Local reference data** (not included in the repo — must be supplied per environment):
  - A BUDB SQLite database, selected on the **BUDB** tab at the start of each session (any file location — there is no fixed folder or auto-detection). It must contain a `bottoms_up` table with at least one row and, at minimum, the columns: `id`, `contact_group_id`, `Owner`, `Owner (Standardized)`, `Input: Address`, `Input: City`, `Input: State`, `md_address`, `md_city`, `md_state`, `County`, `State`.
  - A Pipedrive deals source at `cache/pipedrive_local_cache.csv`, containing at minimum: `Deal - ID`, `Deal - Title`, `Deal - Contact person`, `Person - Mailing Address`, `Deal - Serial Number`. This is downloaded automatically via the Pipedrive API (see below); there's no supported way to keep it fresh with a manually re-exported file instead.
- **Python packages**: see `requirements.txt` (UI: `customtkinter`; data: `pandas`, `openpyxl`; Google auth: `google-auth`; HTTP: `requests`).

> Install all pinned dependencies from `requirements.txt`.

## ⚠️ Manual Pre-requisites (Current Process)

Pipedrive deals are fetched automatically via the Pipedrive API once a token is entered in the app's **Pipedrive** tab (see **4.3** below) — there's no separate export/rename step needed for this in normal use. The following still require manual setup:

1. **Select the current BUDB `.db` file** on the **BUDB** tab (**Browse** button) at the start of each session — the app does not remember the previous selection between launches, and does not generate, download, or refresh this file on its own.
2. Confirm `config/google_service_account.json` exists and the target Google Sheet is shared with that service account's email — this is required, since the Google Sheet is the only supported import source.
3. **A Pipedrive API token is required** — there is no UI toggle to disable the API and fall back to a manually-exported CSV. `cache/pipedrive_local_cache.csv` is only reused as-is, without any token, while it's less than 8 hours old; once it expires, a valid token is needed to refresh it (enter it on the **Pipedrive** tab and click **Save API token**).


## 🚀 Installation and Setup
1. **Clone repository**
    ```bash
    git clone https://github.com/dyoliya/deo_db_builder.git
    cd deo_db_builder
    ```

2. **Install dependencies**
    ```bash
    pip install -r requirements.txt
    ```

3. **Folder Structure**

    <pre>project/
    │
    ├── app.py                          Main application entry point (GUI)
    ├── deo_transform.py                 Column cleaning, phone/date normalization
    ├── deo_sqlite.py                    SQLite schema + row sync helpers
    ├── deo_deal_id_matcher.py           Row-level Pipedrive deal ID matching
    ├── pipedrive_lookup.py              Pipedrive cache loading + lookup maps
    ├── pipedrive_api.py                 Pipedrive API client + deals download/cache refresh
    ├── deo_budb.py                      Row-level BUDB ID matching + contact-group auto-resolution + cache
    ├── deo_build_state.py               Tracked build lifecycle: prepare/finish/fail-build, pre-build snapshots
    ├── deo_existing_deals_report.py     Associated-deals report generation
    ├── deo_action_log.py                Diffs before/after DB snapshots into an action log (run from app.py)
    ├── initialize_deo_baseline.py       Standalone: stamps a known-good DB as the tracked baseline (see Baseline Snapshots & Test Runs)
    ├── requirements.txt
    ├── version.txt
    ├── config/                         Google service account + Pipedrive settings + build-state tracking (local only)
    │   ├── google_service_account.json
    │   ├── pipedrive_settings.json
    │   └── build_state.json
    ├── cache/                          Local lookup caches + Pipedrive deals CSV (local only)
    ├── database/                       Final output: the daily DEO database(s) (auto-created)
    │   └── previous_versions/          Pre-build snapshots, archived on every build attempt (and by initialize_deo_baseline.py)
    ├── audit/                          Combined NULL-conversion + BUDB group-selection audit exports (auto-created)
    ├── existing_deals/                 Deal-match audit + associated-deals reports (auto-created)
    ├── google_sheet_writeback/         Pending/completed Google writeback records (auto-created)
    └── input/                          Local backup copies of each downloaded Google Sheet (auto-created)
        └── google_sheet_backups/       One subfolder per spreadsheet ID, holding timestamped .xlsx snapshots
    </pre>

    That's everything needed for normal use — the **Update Action Log** button runs the comparison itself using the build's own recorded snapshots, with no file copying needed.

    Everything under `config/`, `cache/`, `database/`, `audit/`, `existing_deals/`, `google_sheet_writeback/`, and `input/` is environment-specific data and is excluded from version control via `.gitignore`.

4. **Set Up Configuration**

    4.1. **Google Service Account** (required for Google Sheets import mode)
    - In Google Cloud Console, create (or reuse) a service account and download its JSON key.
    - Save it as `config/google_service_account.json`.
    - Share the target Google Sheet (tab must be named `deo_db`) with the service account's `client_email` address.
    - No further setup is needed — the app reads this file automatically on each Google Sheets import.

    > Skip this step if a service account already exists for this project — just get `config/google_service_account.json` handed over to you and place it in that path.

    4.2. **BUDB Reference Database**
    - No placement needed — on the **BUDB** tab, click **Browse** and select the current Bottoms Up `.db` file from wherever it's kept.
    - The app validates it on selection: it must have a `bottoms_up` table with at least one row, and include `id`, `contact_group_id`, `Owner`, `Owner (Standardized)`, `Input: Address`, `Input: City`, `Input: State`, `md_address`, `md_city`, `md_state`, `County`, and `State` columns.
    - This selection is **not remembered** between app launches — re-select it (or confirm it's still selected) each session, via **A. First-time preparation** below.

    4.3. **Pipedrive Deals (API download)**
    - On the app's **Pipedrive** tab, paste in a Pipedrive API token (use the **Open API token page** button to find it in Pipedrive under Settings → API) and click **Save API token**. It's saved to `config/pipedrive_settings.json` so it only needs to be entered once.
    - There is no manual download button — **Check references** (and, if needed, **Build DB** itself) automatically downloads fresh deals whenever the cache is missing or older than 8 hours (hardcoded, not configurable).
    - Downloaded deals are written to `cache/pipedrive_local_cache.csv`, so everything downstream (deal matching, the associated-deals report) works the same either way.
    - A valid token is required to refresh the cache. If the cache file is less than 8 hours old, it's reused without needing a token at all — which is the only way a manually-placed CSV can substitute for the API, and only until it ages out.
    - `base_url` and `deals_filter_id` must also be set in `config/pipedrive_settings.json` for a download to succeed.

## 🖥️ User Guide

### A. First-time preparation (repeat each session)
1. Open the project folder.
2. Confirm `config/google_service_account.json` exists (if you'll use Google Sheets import).
3. On the **BUDB** tab, click **Browse** and select the current Bottoms Up `.db` file — this is **not** remembered from the last launch, so it must be re-selected every time the app starts.
4. In the **Pipedrive** tab, confirm an API token is entered and saved (or, if the cache is still less than 8 hours old, confirm `cache/pipedrive_local_cache.csv` exists and is reasonably up to date).

### B. Import deceased-owner records
1. Launch the app (`python app.py`, or the compiled executable).
2. On the **Pipedrive** tab, paste your API token and click **Save API token** (skip if already saved from a previous session).
3. On the **BUDB** tab, click **Browse** and select the current Bottoms Up `.db` file — this isn't remembered between launches, so it's needed every session.
4. Paste the Google Sheet URL into the **Google Sheets URL** field (its `deo_db` tab will be read).
5. Click **Check references** — confirms the sheet is reachable and editable, and the Pipedrive API token is valid. This must succeed before **Build DB** is enabled.
6. Click **Build DB**.
7. Watch the **Activity Log** for progress: cleaning, Pipedrive matching, BUDB matching, county enrichment, schema validation, and row sync.
8. On completion, a dialog confirms the build and prompts you to click **Update Action Log**.
9. Click **Update Action Log** to diff the before/after snapshots and record the change history (see **Action-log flow** above). The button is only enabled once a build is `ready_for_log`, and shows **Action Log updated** once that build is `complete`.

### C. Review outputs
- **The final output is the daily database itself**: `database/YYYY-MM-DD-deo.db` — this is what the whole build updates, and the summary dialog on completion shows its path and table name.
- Check `audit/` for the combined NULL-conversion / BUDB group-selection audit workbook if the log reported suspicious conversions or multi-group BUDB matches.
- Check `existing_deals/` for the deal-match audit and the DEO-associated-deals report generated after every import.
- For Google Sheets imports, check the Activity Log for the writeback confirmation and record path under `google_sheet_writeback/`.

### D. Review the change log (`deo_action_log.py`)
After clicking **Update Action Log** (step B.9 above), review the logged entries in the `deo_action_log` table of the database in `database/` — check that the added/deleted/modified rows make sense for that version.
- **If everything looks correct**: nothing further to do — that build is `complete` and is the official daily version.
- **If something looks wrong**: coordinate with **Jennifer Cariño** on the flagged log entries to correct the input (Google Sheets) data, then:
  1. Restore the pre-build snapshot from `database/previous_versions/` back into `database/`.
  2. Re-run the import in `app.py` against the corrected input to rebuild a new version of the database.
  3. Review the action log again for this rebuilt version before treating it as final.

### E. Important usage notes
- Each calendar day gets its own database file; re-running on the same day continues syncing into that day's file rather than creating a new one.
- Every click of **Build DB** archives a pre-build snapshot into `database/previous_versions/<build_id>/` first, and a failed build is automatically rolled back to that snapshot (unless a Google writeback had already started — see Troubleshooting).
- The **Build DB** button is disabled unless the previous build finished (`complete` or `rolled_back`); an unfinished or `recovery_required` build must be resolved first.
- If a Google Sheets writeback is interrupted, a pending record remains in `google_sheet_writeback/` and must be resolved before the next Google Sheets import can run.
- Incoming files with columns the database doesn't recognize will be rejected with a schema-mismatch error; missing (already-known) columns are allowed.
- `deo_action_log.py` runs automatically from `app.py` via the **Update Action Log** button — see **D** above.

### F. Baseline snapshots & test runs (`initialize_deo_baseline.py`)
`initialize_deo_baseline.py` is a **separate, manually-run script** (not launched from `app.py`). It stamps a known-good `database/*.db` file as the officially tracked baseline and records a safety-net backup of it, so you have a confirmed-stable snapshot to fall back on — most useful right before doing a test/experimental run you might want to discard.

1. **Prepare the stable database**: make sure the `.db` file you want to treat as "stable" is the **only** `.db` file sitting directly inside `database/` (move anything else out first).
2. **Point the script at it**: open `initialize_deo_baseline.py` and edit the hardcoded filename on the `database = folder / "..."` line to match that file's name exactly.
3. **Run it**:
   ```bash
   python initialize_deo_baseline.py
   ```
4. The script:
   - Verifies that file is the only `.db` in `database/`, that it actually has a `deo` table, and that it does **not** already carry build-state tracking (`deo_build_state` table or `config/build_state.json`) — it refuses to run otherwise, so it never silently overwrites existing tracking.
   - Verifies no pending Google writeback files exist in `google_sheet_writeback/`.
   - Copies the database, as-is, into `database/previous_versions/baseline-<id>/` — this is your restore point.
   - Stamps the live database (and `config/build_state.json`) with a fresh `baseline-<id>` build, status `complete`, so the app's **Build DB** button is immediately usable again.
5. You can now do your test run(s) through the app as usual. Each build attempt still creates its own `previous_versions/<build_id>/` snapshot automatically, and a failed build auto-restores on its own.
6. **To throw away a test run and go back to the stable baseline**: close the app, replace the file in `database/` with the copy saved in `database/previous_versions/baseline-<id>/`, delete `config/build_state.json`, and re-run `initialize_deo_baseline.py` (pointed at the restored filename) to re-stamp it as a clean, trackable baseline before resuming normal use.

### G. Common troubleshooting
- **"Pipedrive local cache CSV not found"**: `cache/pipedrive_local_cache.csv` is missing. Click **Check references** with a valid API token entered on the **Pipedrive** tab to download it.
- **"Pipedrive cache is missing or expired. Enter your API token and click Save API token."**: The cache is gone or older than 8 hours and no token is saved. Enter a token on the **Pipedrive** tab and click **Save API token**, then **Check references**.
- **"Cannot write to the Pipedrive cache CSV or its folder."**: Close `cache/pipedrive_local_cache.csv` if it's open in Excel or elsewhere, and confirm you have write permission to the `cache/` folder.
- **Pipedrive download errors (bad token, rate limited, field not found, etc.)**: The error message in the Activity Log / popup names the specific problem (e.g. an unexpected or renamed Pipedrive field) — check `config/pipedrive_settings.json`'s `deal_custom_field_labels` against the actual field names in Pipedrive if a field-not-found error appears.
- **"Select a Bottoms Up .db file first." / "Select an existing .db file."**: Click **Browse** on the **BUDB** tab and select the current Bottoms Up `.db` file — this isn't remembered between app launches.
- **"This file has no bottoms_up table." / "BUDB is missing columns: ..." / "The bottoms_up table contains no records."**: The selected file isn't a valid BUDB export — confirm it has a populated `bottoms_up` table with the required columns listed under **Requirements** above.
- **"A previous Google writeback needs recovery"**: Resolve the pending file in `google_sheet_writeback/` before starting another Google Sheets import.
- **"The Google Sheet changed while downloading"**: Re-run the import after editing of the sheet has stopped.
- **Schema mismatch errors**: Compare the incoming file's columns against the existing database schema reported in the error message.
- **"The previous build is not finished"**: A prior build is stuck at `building`, `ready_for_log`, or `recovery_required`. Finish it (click **Update Action Log**) or resolve its recovery state before clicking **Build DB** again.
- **"Recovery required" / "Build tracking disagrees; recovery required." / "Recorded build database is missing; recovery required."**: `config/build_state.json` and the database's own tracking no longer agree (e.g. a `.db` file was moved or replaced by hand). Inspect `database/` and `config/build_state.json` and the relevant `previous_versions/` snapshot before making further changes — don't delete anything until the mismatch is understood.
- **"Google writeback may have occurred; recovery is required."**: A build failed *after* the Google Sheets writeback had already started, so the sheet may be partially updated. The database is preserved as-is for manual review — check the sheet and the pending record in `google_sheet_writeback/` before deciding next steps.
- **(From `initialize_deo_baseline.py`) "Expected only `<name>` directly inside database/."**: Remove any other `.db` files from `database/`, or edit the script's hardcoded filename to match the one stable file you placed there.
- **(From `initialize_deo_baseline.py`) "Tracking file already exists" / "This database already has a build-state table."**: The target database is already tracked — inspect its current state with the app instead of re-initializing it.
- **(From `initialize_deo_baseline.py`) "Pending Google writeback files exist."**: Resolve the pending file(s) in `google_sheet_writeback/` before initializing a baseline.

---

## 👩‍💻 Credits
- **2026-10-09**: Project created by **Julia** ([@dyoliya](https://github.com/dyoliya))
- 2026–present: Maintained by **Julia** for **Community Minerals II, LLC**
