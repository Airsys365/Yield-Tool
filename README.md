# Yield & Scrap Cost Tool

Desktop tool for weekly yield and scrap cost reporting in electronics manufacturing. Reads a weekly production export, joins it with an operation cost table, and appends yield / scrap metrics to a persistent `results.xlsx` workbook shared via SharePoint or OneDrive.

## Install & run

```bash
pip install -r requirements.txt
python app.py
```

## Workflow

1. **Pick the files** — *DB source* (the weekly production export), *Costs* (operation cost table), and *Results* (the workbook to append to; created on first run).
2. **Set the year** and click **Check files**. The preview shows which weeks are new and a form appears for *Completed Detectors* counts (units shipped per week).
3. **Fill in detector counts** for the new weeks (leave blank to fill in later).
4. Click **Add new weeks to results.xlsx**.

Existing weeks are never recalculated — scrap costs are locked at the prices active when each week was first processed. Updating the Costs file only affects future weeks.

## Input formats

| File | Required columns |
|---|---|
| DB source | `Item`, `OperationNo`, `TaskNo`, `TaskDescription`, `InsertDate`, `Good`, `Bad` |
| Costs     | `Product`, `Item`, `OperationNo`, `AccumulatedCost` (optional `Name`, `TaskDescription`, `TaskNo`) |

Column names are matched by keyword (case-insensitive). European decimal commas (`3,42`) are accepted in the Costs file.

## Output

`results.xlsx` contains the following sheets:

- **ProductSummary** — yield, scrap qty, scrap cost, scrap per detector, per product / week.
- **ComponentSummary** — same metrics broken down by component (item).
- **WeeklyDetail** — one row per operation per week.
- **Charts** — weekly trends for scrap cost, scrap per detector, and per-product yield.
- **Detectors** — applied detector counts.

## Files in this repo

| File | Purpose |
|---|---|
| `app.py` | CustomTkinter GUI |
| `core.py` | Data loading, filtering, calculation, merging |
| `excel_writer.py` | Excel output (sheets, styles, charts) |
| `make_test_data.py` | Generates `test_raw.xlsx` for development |
| `generate_report.py` | Old single-file CLI version (kept for reference) |

See `CLAUDE.md` for architectural notes.
