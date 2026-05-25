# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Project Purpose

Desktop tool for weekly yield and scrap cost reporting in electronics manufacturing. Reads raw production data exported from a database, joins with operation cost table, calculates yield/scrap metrics, and appends results to a persistent Excel workbook. Designed for one person to run weekly; results are shared via SharePoint/OneDrive.

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
raw      = core.load_raw('test_raw.xlsx')
costs    = core.load_costs('test_raw.xlsx')      # Costs sheet in the same file
existing = core.load_results('results.xlsx')
info     = core.analyse(raw, costs, existing)
new      = core.process_new_weeks(raw, costs, info['weeks_to_add'])
merged   = core.merge_results(existing, new)
merged, n = core.apply_detectors(merged, core.load_detectors('test_raw.xlsx'))
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

**Append-only results.xlsx** — `core.load_results()` reads existing calculated data, `core.process_new_weeks()` only calculates weeks not already present, `core.merge_results()` concatenates old + new. Old weeks are never recalculated even if `costs.xlsx` changes. This preserves historical scrap costs at the prices that were active when each week was processed.

**costs.xlsx is the source of truth for item/product mapping** — there is no hardcoded item list. The `Product` column in costs.xlsx determines which product each item belongs to. Items in the raw data that are absent from costs.xlsx are silently ignored (logged in GUI). This means adding a new product or item requires only editing costs.xlsx.

**Week key is ISO format** — `YYYY-Www` (e.g. `2026-W19`), derived via `strftime("%G-W%V")`. Used as the deduplication key between raw data and results.

**Prices are locked at processing time** — scrap cost = `Bad × AccumulatedCost` where `AccumulatedCost` is the total accumulated unit cost up to that operation (not the marginal cost of the operation itself).

---

## Input File Formats

The raw file is a single source of truth holding three sheets: the raw export, `Costs`, and `Detectors`.

**Raw production export** (first sheet) — required columns: `Item`, `OperationNo`, `TaskNo`, `TaskDescription`, `InsertDate`, `Good`, `Bad`. Column names are detected flexibly by keyword matching; positional fallback handles 7- or 8-column variants (some exports have an empty column between InsertDate and Good).

**Costs sheet** — required columns: `Product`, `Item`, `OperationNo`, `TaskNo`, `AccumulatedCost`. Optional: `Name`, `TaskDescription`. European decimal comma (`3,42`) is handled automatically. `core.load_costs()` reads the sheet named `Costs`, falling back to the first sheet.

**Detectors sheet** (optional) — matrix layout: first column = `Week`, every other column header names a product, cells hold the count. Headers may carry extra text (e.g. `Shallow SH2A295`) — the product name inside the header is matched, so `core.load_detectors(path, products)` takes the known product list. `Week` is the calendar week (`W18`, `2026-W18`, or `18`; zero-padding tolerated). A long `Product / Week / Completed Detectors` layout is also accepted as a fallback. Read by `core.load_detectors()`, applied by `core.apply_detectors()`.

**results.xlsx** — written and read by this tool. Sheets: `ProductSummary`, `ComponentSummary`, `WeeklyDetail`, `Charts`, `Detectors`. Never store this in git — it lives in SharePoint/OneDrive.

---

## Completed Detectors

`CompletedDetectors` (units shipped per week, used for scrap/unit KPI) is not in the raw production data. It is supplied via the `Detectors` sheet in the **raw file** (single source of truth): `core.load_detectors()` reads it and `core.apply_detectors()` fills `completed_detectors` / `scrap_per_detector` on every run. Keeping it in the raw file means regenerating `results.xlsx` from scratch (new path) never loses a year of counts. The `scrap_per_detector` field stays `null` for any product/week missing from the Detectors sheet. There is no longer a detector-entry form in the GUI.

---

## Notes

- `config.json` is auto-created next to the script and stores last-used file paths across sessions.
- The GUI runs file I/O in background threads to keep the UI responsive — all `core.*` calls in `app.py` are inside `threading.Thread(target=_work, daemon=True)`.
