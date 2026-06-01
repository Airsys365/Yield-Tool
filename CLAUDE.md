# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Project Purpose

Desktop tool for weekly yield and scrap cost reporting in electronics manufacturing. Reads raw production data directly from a DB export file, joins with the operation cost table, calculates yield/scrap metrics, and appends results to a persistent Excel workbook. Designed for one person to run weekly; results are shared via SharePoint/OneDrive.

---

## Running the App

```bash
pip install -r requirements.txt
python app.py
```

Generate test data for development:
```bash
python make_test_data.py        # creates test_raw.xlsx with Raw + Costs + Detectors sheets
```

Run pipeline headlessly (for testing logic without GUI):
```bash
python -c "
import core, excel_writer
raw      = core.load_raw('test_raw.xlsx')        # DB source (Raw sheet or csv)
costs    = core.load_costs('test_raw.xlsx')      # Costs sheet, or a separate costs.xlsx
existing = core.load_results('results.xlsx')
info     = core.analyse(raw, costs, existing)
new      = core.process_new_weeks(raw, costs, info['weeks_to_add'])
merged   = core.merge_results(existing, new)
# Detector counts: GUI form in the app; tests can read a Detectors sheet
dets     = core.load_detectors('test_raw.xlsx', sorted(set(costs['Product'])))
merged, n = core.apply_detectors(merged, dets)
excel_writer.save_results('results.xlsx', merged)
"
```

---

## Architecture

Three files with strict separation:

| File | Role |
|---|---|
| `core.py` | All data logic — loading, filtering, calculation, merging |
| `excel_writer.py` | All Excel output — sheets, styles, charts |
| `app.py` | CustomTkinter GUI only — calls core and excel_writer, no business logic |

`generate_report.py` is the old single-file version (pre-refactor). It still works as a CLI but is superseded by the three-file architecture.

---

## Key Design Decisions

**Two independent input files** — the DB source (weekly production export) and the Costs table are picked separately in the GUI and read independently. There is no intermediate "raw" file step; `core.load_raw()` reads the DB export directly and `core.load_costs()` reads the cost table from its own path.

**Append-only results.xlsx** — `core.load_results()` reads existing calculated data, `core.process_new_weeks()` only calculates weeks not already present, `core.merge_results()` concatenates old + new. Old weeks are never recalculated even if the Costs file changes. This preserves historical scrap costs at the prices that were active when each week was processed.

**costs.xlsx is the source of truth for item/product mapping** — there is no hardcoded item list. The `Product` column in costs.xlsx determines which product each item belongs to. Items in the raw data that are absent from costs.xlsx are silently ignored (logged in GUI). This means adding a new product or item requires only editing costs.xlsx.

**Year filter on the DB source** — the GUI has a Year field (defaults to current year). When "Check files" runs, only DB-source rows whose `InsertDate` falls in that ISO year are considered. Lets the user point at a large historical export and process one year at a time.

**Week key is ISO format** — `YYYY-Www` (e.g. `2026-W19`), derived via `strftime("%G-W%V")`. Used as the deduplication key between raw data and results.

**Prices are locked at processing time** — scrap cost = `Bad × AccumulatedCost` where `AccumulatedCost` is the total accumulated unit cost up to that operation (not the marginal cost of the operation itself).

**Detector-only weeks are allowed** — if the user enters a completed-detector count for a product/week that has no production rows, `core.apply_detectors()` creates a detector-only stub in `ProductSummary` (yield / scrap fields stay null). This keeps the weekly detector series continuous even when no scrap was reported.

---

## Input File Formats

**DB source** (the weekly production export) — required columns: `Item`, `OperationNo`, `TaskNo`, `TaskDescription`, `InsertDate`, `Good`, `Bad`. Column names are detected flexibly by keyword matching; positional fallback handles 7- or 8-column variants (some exports have an empty column between InsertDate and Good). Read by `core.load_raw()`.

**Costs** — required columns: `Product`, `Item`, `OperationNo`, `AccumulatedCost`. Optional: `Name`, `TaskDescription`, `TaskNo`. European decimal comma (`3,42`) is handled automatically. `core.load_costs()` reads the sheet named `Costs`, falling back to the first sheet — so a standalone `costs.xlsx` works too.

**Detectors (GUI form)** — completed detector counts (units shipped per week) are entered through the GUI. After "Check files", a scrollable entry block appears with one row for each new product/week pair, plus any gap weeks between the last processed week and today where production rows are missing. Counts left blank are skipped and can be filled in on a later run.

**Detectors sheet (test / headless only)** — `core.load_detectors(path, products)` can still read a `Detectors` sheet from an xlsx in matrix layout (first column `Week`, remaining columns are product headers — extra words tolerated, e.g. `Shallow SH2A295`) or in long `Product / Week / Completed Detectors` layout. Used by `make_test_data.py` and the headless pipeline; not used by the GUI.

**results.xlsx** — written and read by this tool. Sheets: `ProductSummary`, `ComponentSummary`, `WeeklyDetail`, `Charts`, `Detectors`. Never store this in git — it lives in SharePoint/OneDrive.

---

## Completed Detectors

`CompletedDetectors` (units shipped per week, used for the scrap/unit KPI) is not in the DB production export. The GUI form collects it after "Check files": one entry per new product/week pair, plus any gap weeks. `core.apply_detectors()` writes the counts into `ProductSummary` and computes `scrap_per_detector = total_scrap_cost / completed_detectors`. For weeks with a count but no production rows, it creates a detector-only stub row (see Key Design Decisions). The `scrap_per_detector` field stays `null` for any product/week the user left blank — the next run will prompt for it again as a gap week.

---

## Notes

- `config.json` is auto-created next to the script and stores last-used file paths and the selected year across sessions.
- The GUI runs file I/O in background threads to keep the UI responsive — all `core.*` calls in `app.py` are inside `threading.Thread(target=_work, daemon=True)`.
- The "Add new weeks" button stays gray/disabled until "Check files" finds something to process; it turns blue when active.
