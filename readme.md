# Excel Schema Diff

Reconcile member records between an internal system of record and an external
consultant's data feed — even when the two sides use different column names,
date formats, and casing conventions.

The reconciliation logic is entirely **mapping-driven**: no external column
name is hardcoded in the script. Onboarding a new plan with a differently
formatted consultant feed is a spreadsheet edit, not a code change.

## Table of contents

- [How it works](#how-it-works)
- [Project structure](#project-structure)
- [Data files](#data-files)
- [Usage](#usage)
- [Output](#output)
- [Normalization rules](#normalization-rules)
- [Onboarding a new plan](#onboarding-a-new-plan)
- [Requirements](#requirements)

## How it works

```
internal_system_data.csv ─┐
                           ├─▶ reconcile.py ─▶ reconciliation_report.xlsx
external_consultant_data.xlsx ─┘        ▲
                                         │
                        header_mapping_table.xlsx
                        (canonical field ⇄ external header)
```

1. `reconcile.py` reads `header_mapping_table.xlsx` to learn which internal
   ("canonical") field name corresponds to which external column header.
2. It joins the two datasets on `External ID` / `Client ID`.
3. Every mapped field is **normalized** before comparison (see
   [Normalization rules](#normalization-rules)) so formatting-only
   differences aren't flagged as real changes — only genuine value drift,
   missing records, and new members are.
4. Results are written to a three-tab Excel workbook.

## Project structure

```
excel-schema-diff/
├── data/
│   ├── internal_system_data.csv       # internal system of record
│   ├── external_consultant_data.xlsx  # consultant feed (sample input)
│   ├── header_mapping_table.xlsx      # canonical field ⇄ external header
│   └── reconciliation_report.xlsx     # generated output (see Usage)
├── scripts/
│   └── reconcile.py                   # the reconciliation script
├── requirements.txt
└── readme.md
```

## Data files

| File | Description |
|---|---|
| `internal_system_data.csv` | Internal system of record, using canonical field names (`Internal ID`, `External ID`, `First Name`, … `Updated By`). `Internal ID` is formatted as `{Plan Number}{6-digit sequence}` (e.g. `101000001`). All sample rows belong to Plan `101` — Delta Alpha Construction. |
| `external_consultant_data.xlsx` | The same underlying members as sent by an external consultant, but with different header names (e.g. `Client ID`, `DOB`, `Mem_Status`) and formatting (`MM/DD/YYYY` dates, a merged address field, UPPERCASE status). Intentionally includes a few records missing from the internal system, a few new members not yet onboarded internally, and several fields with drifted values, to simulate real-world data quality issues. `External ID` / `Client ID` is the join key between the two files. |
| `header_mapping_table.xlsx` | Maps each canonical field to its external header, with notes on formatting differences (`Header Mapping` tab). Also includes a `Plan Reference` tab (`Plan Number` → `Plan Name`) so the mapping can be extended to future plans. |
| `reconciliation_report.xlsx` | Output of `scripts/reconcile.py` — see [Output](#output). |

## Usage

```bash
pip install -r requirements.txt

cd scripts
python3 reconcile.py \
    --internal ../data/internal_system_data.csv \
    --external ../data/external_consultant_data.xlsx \
    --mapping ../data/header_mapping_table.xlsx \
    --output ../data/reconciliation_report.xlsx
```

All arguments default to the paths above, so `python3 reconcile.py` with no
flags works out of the box from the `scripts/` directory.

| Flag | Default | Description |
|---|---|---|
| `--internal` | `../data/internal_system_data.csv` | Internal system-of-record CSV |
| `--external` | `../data/external_consultant_data.xlsx` | External consultant feed |
| `--mapping` | `../data/header_mapping_table.xlsx` | Canonical ⇄ external header mapping |
| `--output` | `../data/reconciliation_report.xlsx` | Where to write the report |

## Output

`reconciliation_report.xlsx` has three tabs:

1. **Member Comparison** — one row per member, with `{Field} (Internal)`,
   `{Field} (External)`, and `{Field} Changed` (True/False) columns for every
   mapped field, plus an overall `Any Field Changed` flag.
2. **Differences** — flat list of every discrepancy found: value mismatches,
   records missing from the external feed, and new members not yet onboarded
   internally.
3. **Summary** — record counts (total internal/external members, matched,
   missing, new, and field-level mismatches).

## Normalization rules

| Field type | Example fields | Normalization |
|---|---|---|
| Date | Date of Birth, Join Date, Last Updated, Created Date | Parsed from `YYYY-MM-DD`, `MM/DD/YYYY`, `YYYY/MM/DD`, or `MM-DD-YYYY` and compared as ISO dates |
| Phone | Phone Number | All non-digit characters stripped |
| Postal code | Postal Code | Whitespace stripped, uppercased |
| Everything else | Name, Status, Address, … | Whitespace collapsed, trimmed, lowercased |

## Onboarding a new plan

Because the mapping is data-driven, adding a plan whose consultant feed uses
different headers only requires editing `header_mapping_table.xlsx`:

1. Add a row to the `Plan Reference` tab (`Plan Number` → `Plan Name`).
2. Update the `Header Mapping` tab if the new plan's consultant uses different
   external column names for any canonical field.
3. Re-run `reconcile.py` — no code changes needed.

## Requirements

- Python 3
- `openpyxl` (`pip install -r requirements.txt`)
