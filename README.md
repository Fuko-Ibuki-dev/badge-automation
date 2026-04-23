# Badge Automation Pipeline

Automates the end-to-end processing of skills and role badge data — from raw CSV attachments to a formatted master tracker and ready export files.

---

## Repository Contents

| File | Type | Purpose |
|------|------|---------|
| `badge_pipeline.py` | Python | **Main script** — full pipeline (import → role badges → export) |
| `transform_badges.py` | Python | Standalone: import CSVs to Masterlist only |
| `award_role_badges.py` | Python | Standalone: check and award role badges only |
| `Badge_Data_Formatter.js` | Office Script | Excel automation: formats raw badge data into a staging layout |
| `Append_To_Staging.js` | Office Script | Excel automation: appends formatted rows to the staging sheet |

> **For day-to-day use, run `badge_pipeline.py` only.** The two standalone scripts and Office Scripts are kept for reference and individual use cases.

---

## What `badge_pipeline.py` Does

The pipeline runs in three sequential parts:

```
PART 1 — Import CSVs
  incoming/*.csv  →  deduplicate  →  append to Masterlist sheet

PART 2 — Award Role Badges
  Read Masterlist  →  find eligible users  →  append role badge rows

PART 3 — Export to To_DTO folder
  Fill Badges_Excel_Template.xlsx  +  write combined CSV
  → saved to To/<timestamp>/
```

A single `pipeline_log.xlsx` records every action across all three parts.

---

## Prerequisites

### Python version
Python 3.10 or later (uses `X | Y` union type hints).

### Install dependencies
```bash
pip install pandas openpyxl
```

### Required files (not in this repo — set up locally)
| File | Description |
|------|-------------|
| `master_list.xlsx` | Master badge tracker with a `Masterlist` sheet |
| `Badges_Excel_Template.xlsx` | Upload template (headers on row 3) |
| CSV files in `incoming/` | Raw badge data exported from the training system |

---

## Folder Structure

Set up this folder structure on your machine (paths are configured at the top of `badge_pipeline.py`):

```
badge_automation/
├── incoming/                   ← drop new CSV files here before running
├── processed/                  ← CSVs are moved here automatically after processing
├── To/
│   └── 20250422_143012/        ← timestamped folder created each run
│       ├── Badges_All_<stamp>.xlsx
│       └── Combined_All_<stamp>.csv
├── master_list.xlsx
├── Badges_Excel_Template.xlsx
├── pipeline_log.xlsx           ← created automatically on first run
└── badge_pipeline.py
```

---

## Configuration

All settings are at the **top of `badge_pipeline.py`** — no other file needs editing.

```python
# ── Folders ──────────────────────────────────────────────────────────────
INPUT_FOLDER        = r"C:\...\incoming"
PROCESSED_FOLDER    = r"C:\...\processed"
TO_BASE_FOLDER  = r"C:\...\To"

# ── Files ─────────────────────────────────────────────────────────────────
MASTER_FILE         = r"C:\...\master_list.xlsx"
LOG_FILE            = r"C:\...\pipeline_log.xlsx"
BADGE_TEMPLATE_FILE = r"C:\...\Badges_Excel_Template.xlsx"

# ── Master sheet settings ──────────────────────────────────────────────────
MASTERLIST_SHEET    = "Masterlist"
HEADER_ROW          = 3    # row number of the header row in the Masterlist sheet

# ── Processing mode ────────────────────────────────────────────────────────
PROCESS_MODE = "both"      # "skills" | "roles" | "both"

# ── Export split mode ──────────────────────────────────────────────────────
EXPORT_MODE  = "combined"  # "combined" | "split"
```

### `PROCESS_MODE` options

| Value | What runs | Use when |
|-------|-----------|----------|
| `"skills"` | Part 1 + Part 3 (skills badges only) | Weekly CSV import |
| `"roles"` | Part 2 + Part 3 (role badges only) | Ad-hoc role badge check |
| `"both"` | Parts 1 + 2 + 3 | Full end-to-end run |

### `EXPORT_MODE` options

| Value | Excel output | CSV output |
|-------|-------------|------------|
| `"combined"` | One file with all new rows | One file with all rows |
| `"split"` | One file **per unique badge name** | One combined file with all rows |

---

## Running the Pipeline

1. Drop new CSV files into the `incoming/` folder.
2. Open `badge_pipeline.py` and confirm the CONFIG paths are correct.
3. Set `PROCESS_MODE` and `EXPORT_MODE` as needed.
4. Run:

```bash
python badge_pipeline.py
```

5. Check the console output and `pipeline_log.xlsx` for results.
6. Collect the generated files from `To/<timestamp>/`.

---

## Mode Quick Reference

```
Weekly skills badge import + export:
  PROCESS_MODE = "skills"
  EXPORT_MODE  = "combined"   (or "split" for one Excel per badge)

Ad-hoc role badge check + export:
  PROCESS_MODE = "roles"
  EXPORT_MODE  = "combined"

Full run (import + role badges + export):
  PROCESS_MODE = "both"
  EXPORT_MODE  = "combined"
```

---

## Role Badge Requirements

Role badges are awarded automatically when a user holds **all** required skill badges. The mapping is defined in `ROLE_BADGE_REQUIREMENTS` at the top of `badge_pipeline.py`:

To add a new role badge, add an entry to `ROLE_BADGE_REQUIREMENTS` — no other changes needed.

---

## Input CSV Format

The script accepts column names flexibly (case-insensitive, common aliases recognised). The expected columns are:

| Column | Example |
|--------|---------|
| Identifier (Email_Address) | `john.doe@company.com` |
| Preferred Name(Name to appear on badge) | `John Doe` |
| Name of Skills Badge | `Example Skills Badge` |
| Skills Badge Level | `Basic` |
| Course / Programme Title | `Effective Design` |
| Date of Course Completion | `22-Aug-25` |
| Training Provider | `Example Company` |

Accepted date formats: `22-Aug-25`, `22/08/2025`, `2025-08-22`, Excel serial numbers, and other common variants.

---

## Output Files

### Masterlist sheet (in `master_list.xlsx`)
New rows are appended after the last existing row. No other sheets or existing rows are modified.

| Column | Description |
|--------|-------------|
| Name | Learner's preferred name |
| Email | Upper-cased email address |
| Training Provider | As provided in the CSV |
| Skills Area | Badge name without level |
| Skills Area and Level | Full badge name |
| Badge Level | `Basic` / `Advanced` |
| Date of Award | `D-MMM-YY` format, e.g. `22-Aug-25` |
| Month | Numeric month, e.g. `8` |
| Year | 4-digit year, e.g. `2025` |
| Programme | Course / programme title |

### DTO Excel (`Badges_DTO_<label>_<stamp>.xlsx`)
A filled copy of `Badges_Excel_Template.xlsx` with data starting at row 4.

### Combined CSV (`Combined_<label>_<stamp>.csv`)
All new rows in the original incoming CSV column format — ready to re-import or share.

### Pipeline log (`pipeline_log.xlsx`)
One row per action. Columns: Timestamp, Part, Subject, Rows, Status, Detail.

---

## Office Scripts (Excel)

These run inside Excel via the **Automate** tab and do not require Python.

### `Badge_Data_Formatter.js`
Reads raw badge export data from the active sheet and writes a formatted staging layout to a new `Transformed` sheet. Run this first when a new raw export arrives.

**Source columns read:** Email (A), Preferred Name (D), Skills Badge (E), Programme (G), Completion Date (H), Training Provider (I)

### `Append_To_Staging.js`
Appends the rows from the `Transformed` sheet into the main staging sheet.

**To run either script:** Open the Excel file → Automate tab → select the script → Run.

---

## Duplicate Handling

A row is considered a duplicate if it matches an existing Masterlist row on all three of:
- Email address
- Skills Area and Level
- Date of Award

Duplicates are skipped silently and counted in the log.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `Sheet "Masterlist" not found` | Wrong sheet name or file | Check `MASTERLIST_SHEET` in config |
| `Cannot find column 'Email'` | CSV has an unexpected column name | Add the alias to `COLUMN_ALIASES` |
| Dedup warnings on first run | Header row not found at `HEADER_ROW` | Check `HEADER_ROW` matches actual row in Excel |
| Date of Award is blank | Unrecognised date format | Add the format to `_parse_date()` |
| Role badge not awarded | Skill area name mismatch | Check `Skills Area` column values match `ROLE_BADGE_REQUIREMENTS` keys exactly |

---

## Dependencies

| Package | Version | Use |
|---------|---------|-----|
| `pandas` | ≥ 1.5 | CSV reading, DataFrame operations |
| `openpyxl` | ≥ 3.1 | Excel read/write without destroying formatting |

Both are installable via `pip install pandas openpyxl`.
