# Excel Schema Diff

Sample data and tooling for reconciling member records between an internal system
and an external consultant's data feed, where the same fields arrive under
different header names and formatting conventions.

## Data files (`data/`)

- **`internal_system_data.csv`** — internal system of record, using the canonical
  field names (Internal ID, External ID, First Name, ... Updated By). Internal ID
  is formatted as `{Plan Number}{6-digit sequence}` (e.g. `101000001`). All sample
  rows belong to Plan `101` — Delta Alpha Construction.
- **`external_consultant_data.xlsx`** — the same underlying members as sent by an
  external consultant, but with different header names (e.g. `Client ID`, `DOB`,
  `Mem_Status`) and formatting (`MM/DD/YYYY` dates, merged address field,
  UPPERCASE status). Intentionally includes a few records missing from the
  internal system, a few new members not yet onboarded internally, and several
  fields with drifted values, to simulate real-world data quality issues.
  `External ID` / `Client ID` is the join key between the two files.
- **`header_mapping_table.xlsx`** — maps each canonical field to its external
  header, with notes on formatting differences. Includes a `Plan Reference` tab
  (`Plan Number` → `Plan Name`) so the mapping can be extended to future plans.
- **`reconciliation_report.xlsx`** — output of `scripts/reconcile.py` (see below).

## Reconciliation script (`scripts/reconcile.py`)

Compares the internal and external datasets field-by-field, driven entirely by
`header_mapping_table.xlsx` — no field names are hardcoded, so editing the
mapping table (e.g. onboarding a new plan with different consultant headers)
changes the reconciliation without touching the code.

It normalizes values before comparing so formatting-only differences (date
format, name casing, phone punctuation, postal code spacing) aren't flagged as
real changes, while genuine value drift, missing records, and new members are.

### Usage

```bash
cd scripts
python3 reconcile.py \
    --internal ../data/internal_system_data.csv \
    --external ../data/external_consultant_data.xlsx \
    --mapping ../data/header_mapping_table.xlsx \
    --output ../data/reconciliation_report.xlsx
```

All arguments default to the paths above, so `python3 reconcile.py` with no
flags works out of the box.

Requires `openpyxl` (`pip install openpyxl`).

### Output

`reconciliation_report.xlsx` has three tabs:

1. **Member Comparison** — one row per member, with `{Field} (Internal)`,
   `{Field} (External)`, and `{Field} Changed` (True/False) columns for every
   mapped field, plus an overall `Any Field Changed` flag.
2. **Differences** — flat list of every discrepancy found: value mismatches,
   records missing from the external feed, and new members not yet onboarded
   internally.
3. **Summary** — record counts (total internal/external members, matched,
   missing, new, and field-level mismatches).
