# DEO DB Builder

**DEO DB Builder** is a desktop utility (built with Python) that implements a
**controlled Extract–Transform–Load (ETL) ingestion workflow** for Deceased Owner (DEO) records.

It ingests deceased-owner lead lists from a local file or a shared Google Sheet, cleans and standardizes them, matches each record against the company's Pipedrive deals and BUDB (Bottoms Up Database) mineral-ownership records, and persists the result into a versioned, daily SQLite database for long-term retention and reporting.

---

![Version](https://img.shields.io/badge/version-1.0.0-ffab4c?style=for-the-badge&logo=python&logoColor=white)
![Python](https://img.shields.io/badge/python-3.10%2B-273946?style=for-the-badge&logo=python&logoColor=ffab4c)
![Status](https://img.shields.io/badge/status-active-273946?style=for-the-badge&logo=github&logoColor=ffab4c)

---

## 🚧 Problem Statement / Motivation

Deceased-owner lead lists arrive from different sources — local spreadsheets and a shared Google Sheet that multiple team members contribute to — with inconsistent formatting, duplicate entries, and no link back to the company's existing systems of record (Pipedrive deals, BUDB ownership data).

Manually cross-referencing each incoming owner against Pipedrive and BUDB to check "do we already have a deal or ownership record for this person" is slow and error-prone at volume, and without a durable, versioned database there is no reliable way to track what changed between imports or recover from a bad run.

To close this gap, a repeatable ingestion pipeline is required — one that standardizes incoming data, automatically links it to existing Pipedrive deals and BUDB records, and stores the result in a daily, auditable database.

DEO DB Builder addresses this by enforcing a controlled ETL pipeline:
1. Extract deceased-owner records from a local file or Google Sheet.
2. Clean and standardize columns, phone numbers, and dates, flagging any suspicious data loss for audit.
3. Match each record to a Pipedrive deal ID and a BUDB ID using cached reference data.
4. Sync the result into a daily-versioned SQLite database, tracking inserts/updates/deletes.
5. Write matched IDs back to the Google Sheet (when that was the source) and generate audit/report files for review.

---

## ✨ Features
- **Dual input sources**: local CSV/XLSX files, or a live Google Sheet (tab `deo_db`) via a Google service account.
- **Data cleaning**: column name sanitization, phone number and date normalization, and a NULL-conversion audit for anything that looks like unintended data loss.
- **Pipedrive deal matching** against a local Pipedrive export cache (owner name + mailing address matching), with a SQLite-backed lookup cache for speed.
- **BUDB ID matching** against the company's BUDB reference database (owner name + address matching), with its own signature-based lookup cache that rebuilds only when the BUDB file changes.
- **County enrichment** by combining matched county data from all linked sources.
- **Daily-versioned SQLite database**:
  - Auto-creates/copies forward a new dated DB each day from the most recent prior one.
  - Archives older daily databases into `database/previous_versions/`.
  - Syncs rows by `deo_id`, reporting inserted / updated / unchanged / deleted counts.
- **Google Sheets writeback**: after a successful import, writes deal IDs, serial numbers, and upload dates back to the source sheet, with a pending/completed record so an interrupted writeback can be recovered instead of silently lost.
- **Audit & reporting output**: NULL-conversion audit, deal-match audit, and a DEO-associated-deals report are generated automatically after each import.
- **Operator-friendly desktop UI** with file/sheet selection, progress bar, and a live activity log.
- **Change log between daily DB versions** (`deo_action_log.py`, run separately from `app.py`): diffs an "old" and a "new" DB snapshot and records every schema change, added row, deleted row, and modified field into a persistent `deo_action_log` table, carrying that history forward into each new DB version.
---

## 🧠 Logic Flow

### Import flow
1. User selects an input source — **Local file** (CSV/XLSX) or **Google Sheets** (pastes the sheet URL; reads tab `deo_db`).
2. App resolves today's database (`database/YYYY-MM-DD-deo.db`):
   - If it already exists, continue syncing into it.
   - If not, copy forward the most recent prior daily DB (or start a brand-new one if none exists), and move older dated DBs into `database/previous_versions/`.
3. For each input file (or the sheet):
   - Reads rows. For Google Sheets, verifies the sheet hasn't changed mid-download and skips fully blank rows.
   - Cleans and standardizes columns, phone numbers, and dates; records any suspicious NULL conversions for audit.
   - Matches each row to a **Pipedrive deal ID** using the local Pipedrive cache (`cache/pipedrive_local_cache.csv`).
   - Matches each row to a **BUDB ID** using the BUDB reference database (`budb/*.db`).
   - Combines county information from all matched sources.
4. Validates incoming columns against the existing DB schema — blocks unexpected columns, allows known-missing ones.
5. Syncs rows into the daily table keyed by `deo_id`, assigning `deo_serial_number` and `date_uploaded` as needed, and reports inserted/updated/unchanged/deleted counts.
6. If the source was Google Sheets: writes deal IDs, serial numbers, and upload dates back to the sheet, tracking a pending → completed writeback record so a failed writeback can be resumed safely.
7. Saves output artifacts:
   - `audit/` — NULL-conversion audit (only if suspicious conversions were found).
   - `existing_deals/` — deal-match audit and the DEO-associated-deals report.

### Change-log flow (`deo_action_log.py`)
This is a **separate, manually-run script** — it is not launched by `app.py` and does not run automatically as part of an import.

1. After `app.py` finishes updating `database/<today>-deo.db`, manually copy that new daily DB into `new/`, and copy the previous version you're comparing against into `old/` (each folder must contain exactly one `.db` file).
2. Run `python deo_action_log.py`.
3. The script:
   - Reads the date from the `new/` DB's filename (expects a `YYYY-MM-DD` pattern, e.g. `2026-03-12-deo.db`).
   - Creates the `deo_action_log` table inside the **new** DB if it doesn't already exist.
   - If the new DB's log table is still empty, copies forward all historical log rows from the old DB's log table first — this is what carries the full change history along as you roll from one daily DB to the next.
   - Auto-detects the shared data table between old and new (e.g. `deo`) and compares them by `deo_id`:
     - Logs any columns added or removed between the two schemas.
     - Logs any `deo_id` present only in the new DB as `added`, and any present only in the old DB as `deleted` (full row snapshot as JSON).
     - For rows present in both, logs each changed field individually as `modified` (old value → new value). `date_uploaded` is intentionally excluded from field comparison since it changes on every sync.
4. The resulting `deo_action_log` table — now living inside the new DB file — becomes the historical record going forward; that same file is what you'd place in `old/` the next time you run this comparison.
---

## 📝 Requirements
- **Python**: 3.10+ recommended.
- **Services**:
  - Google Cloud service account with Sheets (read/write) and Drive (read-only) scopes — required only for the Google Sheets import mode.
- **Local reference data** (not included in the repo — must be supplied per environment):
  - A BUDB SQLite database placed in `budb/`, containing at minimum: `id`, `contact_group_id`, `Owner`, `Owner (Standardized)`, `Input: Address`, `Input: City`, `Input: State`, `md_address`, `md_city`, `md_state`.
  - A Pipedrive export CSV placed at `cache/pipedrive_local_cache.csv`, containing at minimum: `Deal - ID`, `Deal - Title`, `Deal - Contact person`, `Person - Mailing Address`, `Deal - Serial Number`.
- **Python packages**: see `requirements.txt` (UI: `customtkinter`; data: `pandas`, `openpyxl`; Google auth: `google-auth`; HTTP: `requests`).

> Install all pinned dependencies from `requirements.txt`.

## ⚠️ Manual Pre-requisites (Current Process)

As of this initial version, there is **no automated fetch** for Pipedrive deals or BUDB data — the app only reads whatever is already sitting on disk. Before running `app.py` to build or update a DEO database, the following must be prepared manually each time fresh data is needed:

1. **Export all Pipedrive deals** and save/rename the file as `cache/pipedrive_local_cache.csv` (required columns listed under **Requirements** above). There is no in-app button or background job that pulls this from Pipedrive's API — it is a manual export step every time you want up-to-date deal matching.
2. **Place the current BUDB `.db` file** in `budb/` (exactly one file expected). The app does not generate, download, or refresh this file on its own.
3. **(Google Sheets import only)** Confirm `config/google_service_account.json` exists and the target Google Sheet is shared with that service account's email.


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
    ├── deo_budb.py                      Row-level BUDB ID matching + cache
    ├── deo_existing_deals_report.py     Associated-deals report generation
    ├── deo_action_log.py                Standalone utility: diffs old/ vs new/ DB snapshots into an action log
    ├── requirements.txt
    ├── version.txt
    ├── config/                         Google service account credentials (local only)
    │   └── google_service_account.json
    ├── budb/                           BUDB reference database (local only)
    ├── cache/                          Local lookup caches + Pipedrive export CSV (local only)
    ├── database/                       Daily DEO databases (auto-created)
    │   └── previous_versions/          Archived daily databases
    ├── audit/                          NULL-conversion audit exports (auto-created)
    ├── existing_deals/                 Deal-match audit + associated-deals reports (auto-created)
    ├── google_sheet_writeback/         Pending/completed Google writeback records (auto-created)
    ├── input/                          Local input files (optional, for manual drops)
    ├── old/                            Manually placed: previous DB snapshot, for deo_action_log.py
    └── new/                            Manually placed: latest DB snapshot, for deo_action_log.py
    </pre>

    Everything under `config/`, `budb/`, `cache/`, `database/`, `audit/`, `existing_deals/`, `google_sheet_writeback/`, and `input/` is environment-specific data and is excluded from version control via `.gitignore`.

4. **Set Up Configuration**

    4.1. **Google Service Account** (required for Google Sheets import mode)
    - In Google Cloud Console, create (or reuse) a service account and download its JSON key.
    - Save it as `config/google_service_account.json`.
    - Share the target Google Sheet (tab must be named `deo_db`) with the service account's `client_email` address.
    - No further setup is needed — the app reads this file automatically on each Google Sheets import.

    4.2. **BUDB Reference Database**
    - Place exactly one `.db` file inside `budb/`. The app expects exactly one file there and will raise an error if none or more than one is found.

    4.3. **Pipedrive Local Cache**
    - Export deals from Pipedrive and save the file as `cache/pipedrive_local_cache.csv`, with the required columns listed under **Requirements** above.
    - The app rebuilds its internal lookup cache automatically whenever this CSV changes.

## 🖥️ User Guide

### A. First-time preparation
1. Open the project folder.
2. Confirm `config/google_service_account.json` exists (if you'll use Google Sheets import).
3. Confirm `budb/` contains exactly one BUDB `.db` file.
4. Confirm `cache/pipedrive_local_cache.csv` exists and is reasonably up to date.

### B. Import deceased-owner records
1. Launch the app (`python app.py`, or the compiled executable).
2. Choose an **Input source**: `Local file` or `Google Sheets`.
   - For local files, click **Browse** and select one or more CSV/XLSX files.
   - For Google Sheets, paste the spreadsheet URL (its `deo_db` tab will be read).
3. Click **Import & Sync to DB**.
4. Watch the **Activity Log** for progress: cleaning, Pipedrive matching, BUDB matching, county enrichment, schema validation, and row sync.
5. On completion, a summary dialog shows the daily database path and table used.

### C. Review outputs
- Check `audit/` for a NULL-conversion audit if the log reported suspicious conversions.
- Check `existing_deals/` for the deal-match audit and the DEO-associated-deals report generated after every import.
- For Google Sheets imports, check the Activity Log for the writeback confirmation and record path under `google_sheet_writeback/`.

### D. Update the change log (`deo_action_log.py`)
Run this after an import whenever you want the historical change log updated for that new DB version:
1. After `app.py` finishes and `database/<today>-deo.db` is updated, copy that file into `new/`.
2. Copy the previous version you're comparing against (e.g. yesterday's DB, or whatever DB currently holds the existing `deo_action_log` history) into `old/`.
   - Only one `.db` file should be in each of `old/` and `new/` at a time — clear out the previous run's files first.
3. Run:
   ```bash
   python deo_action_log.py
   ```
4. The script logs schema changes, added rows, deleted rows, and modified fields into the `deo_action_log` table inside the **new** DB, and carries forward the old DB's log history automatically if the new DB doesn't have any yet.
5. **Review the logged entries** in the `deo_action_log` table of the DB in `new/` — check that the added/deleted/modified rows make sense for that version.
   - **If everything looks correct**: replace the DB in `database/` with the reviewed `.db` from `new/` — this makes it the official daily version.
   - **If something looks wrong**: coordinate with **Jennifer Cariño** on the flagged log entries to correct the input (Google Sheets) data, then:
     1. Revert `database/` by moving the previous good DB back from `database/previous_versions/` into `database/`.
     2. Re-run the import in `app.py` against the corrected input to rebuild a new version of the database.
     3. Repeat the change-log comparison (steps 1–5) against this rebuilt version before replacing `database/` again.

### E. Important usage notes
- Each calendar day gets its own database file; re-running on the same day continues syncing into that day's file rather than creating a new one.
- If a Google Sheets writeback is interrupted, a pending record remains in `google_sheet_writeback/` and must be resolved before the next Google Sheets import can run.
- Incoming files with columns the database doesn't recognize will be rejected with a schema-mismatch error; missing (already-known) columns are allowed.
- `deo_action_log.py` is a separate, standalone script (not launched from `app.py`) — see **D** above for the manual steps.

### F. Common troubleshooting
- **"Pipedrive local cache CSV not found"**: Ensure `cache/pipedrive_local_cache.csv` exists with the required columns.
- **"BUDB folder not found" / "No .db file found in BUDB folder"**: Place exactly one BUDB `.db` file in `budb/`.
- **"A previous Google writeback needs recovery"**: Resolve the pending file in `google_sheet_writeback/` before starting another Google Sheets import.
- **"The Google Sheet changed while downloading"**: Re-run the import after editing of the sheet has stopped.
- **Schema mismatch errors**: Compare the incoming file's columns against the existing database schema reported in the error message.
- **"No .db file found in folder" / "Expected only 1 .db file in folder"** (from `deo_action_log.py`): Make sure `old/` and `new/` each contain exactly one `.db` file before running the script.
- **"No common user table found between old and new databases"**: The two DBs don't share a data table with the same name — confirm you placed the correct snapshots in `old/` and `new/`.

---

## 👩‍💻 Credits
- **2026-10-09**: Project created by **Julia** ([@dyoliya](https://github.com/dyoliya))
- 2026–present: Maintained by **Julia** for **Community Minerals II, LLC**
