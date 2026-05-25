#!/usr/bin/env python3
"""
Yield / Scrap Cost Report Generator
====================================
DEPRECATED: superseded by the app.py + core.py + excel_writer.py architecture.
Still works as a standalone CLI but no longer maintained.

Usage:
  python generate_report.py raw_data.xlsx                         # generate report
  python generate_report.py raw_data.xlsx --template              # generate cost template only
  python generate_report.py raw_data.xlsx costs.xlsx              # with costs
  python generate_report.py raw_data.xlsx costs.xlsx detectors.xlsx  # full run

Input files:
  raw_data.xlsx   – flat production table (Item, OperationNo, TaskNo, TaskDescription, InsertDate, Good, Bad)
  costs.xlsx      – operation cost table (Item, OperationNo, TaskNo, TaskDescription, AccumulatedCost)
  detectors.xlsx  – completed detectors (Product, Week, CompletedDetectors)   e.g. "2026-W19"

Output:
  yield_report.xlsx  – multi-sheet Excel workbook
"""

import sys
import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers as xl_numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00

warnings.filterwarnings("ignore")

# ── PRODUCT / ITEM CONFIGURATION ───────────────────────────────────────────────────
PRODUCT_ITEM_MAP: dict[str, str] = {
    # Shallow
    "IP2A295": "Shallow",
    "IH2A295": "Shallow",
    "IP2A296": "Shallow",
    "IH2A296": "Shallow",
    "SH2A295": "Shallow",
    "IM2A296": "Shallow",
    # Deep
    "IP2A280": "Deep",
    "IM2A280": "Deep",
    "IH2A280": "Deep",
    "SH2A280": "Deep",
}

# ── STYLES ────────────────────────────────────────────────────────────────────────────
HDR_DARK  = PatternFill("solid", fgColor="1F3864")
HDR_MID   = PatternFill("solid", fgColor="2E75B6")
HDR_LIGHT = PatternFill("solid", fgColor="BDD7EE")
ALT_ROW   = PatternFill("solid", fgColor="EBF3FB")
WARN_FILL = PatternFill("solid", fgColor="FFE699")
ERR_FILL  = PatternFill("solid", fgColor="FFC7CE")
OK_FILL   = PatternFill("solid", fgColor="C6EFCE")

FONT_WHITE = Font(bold=True, color="FFFFFF", size=10)
FONT_DARK  = Font(bold=True, color="1F3864", size=10)
FONT_WARN  = Font(color="7F6000", size=10)
FONT_ERR   = Font(color="9C0006", size=10)
FONT_OK    = Font(color="276221", size=10)
FONT_BODY  = Font(size=10)

_side = Side(style="thin", color="B8CCE4")
BORDER = Border(left=_side, right=_side, top=_side, bottom=_side)

FMT_PCT   = "0.00%"
FMT_MONEY = '#,##0.00'
FMT_INT   = '#,##0'
FMT_DATE  = "DD.MM.YYYY"


# ── HELPERS ─────────────────────────────────────────────────────────────────────────

def _parse_european_float(val):
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = s.replace(" ", "").replace(" ", "")
    if re.match(r"^-?\d+,\d+$", s):
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def _col_width(ws):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        mx = max((len(str(c.value or "")) for c in col), default=4)
        ws.column_dimensions[letter].width = min(mx + 4, 40)


def _header(ws, row: int, ncols: int, fill=None, font=None):
    fill = fill or HDR_DARK
    font = font or FONT_WHITE
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _body_cell(ws, row, col, value, fmt=None, fill=None, font=None, align="left"):
    c = ws.cell(row, col, value if not (isinstance(value, float) and np.isnan(value)) else None)
    c.border = BORDER
    c.font = font or FONT_BODY
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


# ── DATA LOADING ──────────────────────────────────────────────────────────────────

def _read_file(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(path, sep=None, engine="python")
    return pd.read_excel(path)


def load_raw(path: str) -> pd.DataFrame:
    df = _read_file(path)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for col in df.columns:
        low = col.lower()
        if "item" in low and "col_item" not in col_map:
            col_map["Item"] = col
        elif "operationno" in low.replace(" ", "").replace("_", "") and "OperationNo" not in col_map:
            col_map["OperationNo"] = col
        elif "taskno" in low.replace(" ", "").replace("_", "") and "TaskNo" not in col_map:
            col_map["TaskNo"] = col
        elif "taskdesc" in low.replace(" ", "").replace("_", "") and "TaskDescription" not in col_map:
            col_map["TaskDescription"] = col
        elif "insertdate" in low.replace(" ", "").replace("_", "") and "InsertDate" not in col_map:
            col_map["InsertDate"] = col
        elif "good" in low and "Good" not in col_map:
            col_map["Good"] = col
        elif "bad" in low and "Bad" not in col_map:
            col_map["Bad"] = col

    required = ["Item", "OperationNo", "TaskNo", "TaskDescription", "InsertDate", "Good", "Bad"]
    missing = [r for r in required if r not in col_map]
    if missing:
        cols = list(df.columns)
        if len(cols) >= 7:
            positional = ["Item", "OperationNo", "TaskNo", "TaskDescription", "InsertDate"]
            if len(cols) == 8:
                positional += ["_skip", "Good", "Bad"]
            else:
                positional += ["Good", "Bad"]
            col_map = {k: cols[i] for i, k in enumerate(positional) if k != "_skip"}
        else:
            raise ValueError(f"Cannot map required columns. Missing: {missing}. Found: {list(df.columns)}")

    df = df.rename(columns={v: k for k, v in col_map.items()})
    df = df[required].copy()

    df["Item"]            = df["Item"].astype(str).str.strip()
    df["OperationNo"]     = pd.to_numeric(df["OperationNo"], errors="coerce")
    df["TaskNo"]          = pd.to_numeric(df["TaskNo"], errors="coerce")
    df["TaskDescription"] = df["TaskDescription"].astype(str).str.strip()
    df["InsertDate"]      = pd.to_datetime(df["InsertDate"], dayfirst=True, errors="coerce")
    df["Good"]            = pd.to_numeric(df["Good"], errors="coerce").fillna(0).astype(int)
    df["Bad"]             = pd.to_numeric(df["Bad"], errors="coerce").fillna(0).astype(int)

    invalid_dates = df["InsertDate"].isna().sum()
    if invalid_dates:
        print(f"  ⚠  {invalid_dates} rows with unparseable InsertDate (will be dropped)")
        df = df.dropna(subset=["InsertDate"])

    return df


def load_costs(path: str) -> pd.DataFrame:
    df = _read_file(path)
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Item", "OperationNo", "TaskNo", "AccumulatedCost"]
    missing = [r for r in required if r not in df.columns]
    if missing:
        raise ValueError(f"Cost file missing columns: {missing}. Found: {list(df.columns)}")
    df["Item"]            = df["Item"].astype(str).str.strip()
    df["OperationNo"]     = pd.to_numeric(df["OperationNo"], errors="coerce")
    df["TaskNo"]          = pd.to_numeric(df["TaskNo"], errors="coerce")
    df["AccumulatedCost"] = df["AccumulatedCost"].apply(_parse_european_float)
    return df[required].dropna(subset=["OperationNo", "TaskNo"])


def load_detectors(path: str) -> pd.DataFrame:
    df = _read_file(path)
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Product", "Week", "CompletedDetectors"]
    missing = [r for r in required if r not in df.columns]
    if missing:
        raise ValueError(f"Detectors file missing columns: {missing}. Found: {list(df.columns)}")
    df["Product"]            = df["Product"].astype(str).str.strip()
    df["Week"]               = df["Week"].astype(str).str.strip()
    df["CompletedDetectors"] = pd.to_numeric(df["CompletedDetectors"], errors="coerce")
    return df[required]


# ── PROCESSING ─────────────────────────────────────────────────────────────────────

def assign_product(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Product"] = df["Item"].map(PRODUCT_ITEM_MAP)
    unmapped = df[df["Product"].isna()]["Item"].unique()
    if len(unmapped):
        print(f"  ⚠  Items not in PRODUCT_ITEM_MAP (rows skipped): {list(unmapped)}")
    return df.dropna(subset=["Product"])


def derive_week(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Week"] = df["InsertDate"].dt.strftime("%G-W%V")
    return df


def aggregate_weekly(df: pd.DataFrame) -> pd.DataFrame:
    key = ["Product", "Week", "Item", "OperationNo", "TaskNo", "TaskDescription"]
    return df.groupby(key, as_index=False, dropna=False).agg(
        weekly_good=("Good", "sum"),
        weekly_bad=("Bad", "sum"),
        row_count=("Good", "count"),
        date_min=("InsertDate", "min"),
        date_max=("InsertDate", "max"),
    )


def join_costs(df: pd.DataFrame, costs: pd.DataFrame | None) -> pd.DataFrame:
    df = df.copy()
    if costs is None or costs.empty:
        df["AccumulatedCost"] = np.nan
        return df
    merged = df.merge(
        costs[["Item", "OperationNo", "TaskNo", "AccumulatedCost"]],
        on=["Item", "OperationNo", "TaskNo"],
        how="left",
    )
    missing = merged[merged["AccumulatedCost"].isna()][
        ["Item", "OperationNo", "TaskNo", "TaskDescription"]
    ].drop_duplicates()
    if len(missing):
        print(f"  ⚠  {len(missing)} operations with no cost mapping:")
        print(missing.to_string(index=False, max_rows=20))
    return merged


def calculate_ops(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    total = df["weekly_good"] + df["weekly_bad"]
    df["operation_yield"]      = np.where(total > 0, df["weekly_good"] / total, np.nan)
    df["operation_scrap_cost"] = df["weekly_bad"] * df["AccumulatedCost"]
    return df


def component_summary(ops: pd.DataFrame) -> pd.DataFrame:
    def _composite(yields):
        v = yields.dropna()
        return v.prod() if len(v) else np.nan

    return ops.groupby(["Product", "Week", "Item"], as_index=False).agg(
        component_composite_yield=("operation_yield", _composite),
        component_total_scrap_cost=("operation_scrap_cost", "sum"),
    )


def product_summary(comp: pd.DataFrame, detectors: pd.DataFrame | None) -> pd.DataFrame:
    prod = comp.groupby(["Product", "Week"], as_index=False).agg(
        total_scrap_cost=("component_total_scrap_cost", "sum"),
    )
    if detectors is not None and not detectors.empty:
        prod = prod.merge(detectors, on=["Product", "Week"], how="left")
    else:
        prod["CompletedDetectors"] = np.nan

    prod["scrap_per_completed_detector"] = np.where(
        prod["CompletedDetectors"] > 0,
        prod["total_scrap_cost"] / prod["CompletedDetectors"],
        np.nan,
    )
    return prod.sort_values(["Product", "Week"])


def run_pipeline(raw_path, costs_path=None, detectors_path=None):
    print(f"\n{'='*60}")
    print(f"  Loading raw data: {raw_path}")
    raw = load_raw(raw_path)
    print(f"  Rows loaded: {len(raw)}")

    costs = None
    if costs_path:
        print(f"  Loading costs: {costs_path}")
        costs = load_costs(costs_path)
        print(f"  Cost rows: {len(costs)}")

    detectors = None
    if detectors_path:
        print(f"  Loading completed detectors: {detectors_path}")
        detectors = load_detectors(detectors_path)

    raw  = assign_product(raw)
    raw  = derive_week(raw)
    ops  = aggregate_weekly(raw)
    ops  = join_costs(ops, costs)
    ops  = calculate_ops(ops)
    comp = component_summary(ops)
    prod = product_summary(comp, detectors)

    weeks = sorted(ops["Week"].dropna().unique())
    print(f"  Weeks: {weeks}")
    print(f"  Products: {sorted(ops['Product'].dropna().unique())}")
    print(f"  Items: {sorted(ops['Item'].unique())}")
    print(f"{'='*60}\n")
    return raw, ops, comp, prod


# ── TEMPLATE GENERATOR ───────────────────────────────────────────────────────────

def generate_cost_template(raw_path: str, out_path: str = "costs_template.xlsx"):
    raw = load_raw(raw_path)
    raw = assign_product(raw)
    ops = raw.groupby(
        ["Product", "Item", "OperationNo", "TaskNo", "TaskDescription"],
        as_index=False
    ).size().drop(columns=["size"])
    ops = ops.sort_values(["Product", "Item", "OperationNo", "TaskNo"])
    ops["AccumulatedCost"] = ""

    wb = Workbook()
    ws = wb.active
    ws.title = "Costs"

    cols    = list(ops.columns)
    labels  = ["Product", "Item", "OperationNo", "TaskNo", "Task Description", "Accumulated Cost (€)"]
    ws.row_dimensions[1].height = 36

    for ci, label in enumerate(labels, 1):
        c = ws.cell(1, ci, label)
        c.fill  = HDR_DARK
        c.font  = FONT_WHITE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    for ri, row in enumerate(ops.itertuples(index=False), 2):
        fill = ALT_ROW if ri % 2 == 0 else None
        for ci, val in enumerate(row, 1):
            c = ws.cell(ri, ci, val if val != "" else None)
            c.border = BORDER
            c.font   = FONT_BODY
            if fill:
                c.fill = fill
            if cols[ci-1] == "AccumulatedCost":
                c.number_format = FMT_MONEY
                c.fill = PatternFill("solid", fgColor="FFFFC0")

    _col_width(ws)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    ws2 = wb.create_sheet("Instructions")
    instructions = [
        ("AccumulatedCost", "The TOTAL accumulated cost of the unit UP TO AND INCLUDING this operation (in €)."),
        ("", "E.g. if Punching costs 2.01 and Via filling adds 0.09, Via filling AccumulatedCost = 2.10"),
        ("", "Fill in every yellow cell. Missing values will show as warnings in the report."),
        ("", "Save this file as costs.xlsx and pass it to generate_report.py"),
    ]
    ws2["A1"] = "Column"
    ws2["B1"] = "Description"
    ws2["A1"].font = FONT_DARK
    ws2["B1"].font = FONT_DARK
    for ri, (k, v) in enumerate(instructions, 2):
        ws2.cell(ri, 1, k).font = FONT_BODY
        ws2.cell(ri, 2, v).font = FONT_BODY
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 80

    wb.save(out_path)
    print(f"✅ Cost template saved to: {out_path}")
    print(f"   Fill in the 'AccumulatedCost' column (yellow cells) and re-run with costs file.\n")


# ── EXCEL WRITER ───────────────────────────────────────────────────────────────────

def _freeze(ws, cell="A2"):
    ws.freeze_panes = cell
    ws.sheet_view.showGridLines = False


def write_product_summary(wb: Workbook, prod: pd.DataFrame):
    ws = wb.create_sheet("1. Product Summary")
    _freeze(ws)
    ws.row_dimensions[1].height = 36

    labels = ["Product", "Week", "Completed Detectors", "Total Scrap Cost (€)",
              "Scrap / Detector (€)", "Data Status"]
    for ci, lbl in enumerate(labels, 1):
        c = ws.cell(1, ci, lbl)
        c.fill, c.font = HDR_DARK, FONT_WHITE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    for ri, row in enumerate(prod.itertuples(index=False), 2):
        has_cost = not (isinstance(row.total_scrap_cost, float) and np.isnan(row.total_scrap_cost))
        has_det  = not (isinstance(row.CompletedDetectors, float) and np.isnan(row.CompletedDetectors))
        has_spd  = not (isinstance(row.scrap_per_completed_detector, float)
                        and np.isnan(row.scrap_per_completed_detector))
        bg = ALT_ROW if ri % 2 == 0 else None

        _body_cell(ws, ri, 1, row.Product, fill=bg, align="center")
        _body_cell(ws, ri, 2, row.Week,    fill=bg, align="center")
        _body_cell(ws, ri, 3, row.CompletedDetectors if has_det else None, FMT_INT, fill=bg, align="right")
        _body_cell(ws, ri, 4, row.total_scrap_cost if has_cost else None,  FMT_MONEY, fill=bg, align="right")
        _body_cell(ws, ri, 5, row.scrap_per_completed_detector if has_spd else None, FMT_MONEY, fill=bg, align="right")

        if not has_cost:
            status, sf, ff = "Missing costs", WARN_FILL, FONT_WARN
        elif not has_det:
            status, sf, ff = "Enter detectors", WARN_FILL, FONT_WARN
        else:
            status, sf, ff = "OK ✓", OK_FILL, FONT_OK
        c = ws.cell(ri, 6, status)
        c.fill, c.font, c.border = sf, ff, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

    _col_width(ws)


def write_component_summary(wb: Workbook, comp: pd.DataFrame):
    ws = wb.create_sheet("2. Component Summary")
    _freeze(ws)
    ws.row_dimensions[1].height = 36

    labels = ["Product", "Week", "Item", "Composite Yield", "Total Scrap Cost (€)"]
    for ci, lbl in enumerate(labels, 1):
        c = ws.cell(1, ci, lbl)
        c.fill, c.font = HDR_MID, FONT_WHITE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    comp_sorted = comp.sort_values(["Product", "Week", "Item"])
    for ri, row in enumerate(comp_sorted.itertuples(index=False), 2):
        bg = ALT_ROW if ri % 2 == 0 else None
        _body_cell(ws, ri, 1, row.Product, fill=bg, align="center")
        _body_cell(ws, ri, 2, row.Week,    fill=bg, align="center")
        _body_cell(ws, ri, 3, row.Item,    fill=bg)
        yld = row.component_composite_yield
        c4 = _body_cell(ws, ri, 4, yld if pd.notna(yld) else None, FMT_PCT, fill=bg, align="right")
        if pd.notna(yld) and yld < 0.95:
            c4.fill = WARN_FILL
        _body_cell(ws, ri, 5, row.component_total_scrap_cost, FMT_MONEY, fill=bg, align="right")

    _col_width(ws)


def write_semi_raw(wb: Workbook, ops: pd.DataFrame):
    ws = wb.create_sheet("3. Weekly Detail Table")
    _freeze(ws, "A2")
    ws.row_dimensions[1].height = 48

    labels = [
        "Product", "Week", "Item",
        "Op No", "Task No", "Task Description",
        "Good", "Bad", "Op Yield",
        "Accum. Cost (€)", "Scrap Cost (€)",
        "Source\nRows", "Date From", "Date To",
    ]
    cols = [
        "Product", "Week", "Item",
        "OperationNo", "TaskNo", "TaskDescription",
        "weekly_good", "weekly_bad", "operation_yield",
        "AccumulatedCost", "operation_scrap_cost",
        "row_count", "date_min", "date_max",
    ]

    for ci, lbl in enumerate(labels, 1):
        c = ws.cell(1, ci, lbl)
        c.fill, c.font = HDR_DARK, FONT_WHITE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    ops_sorted = ops.sort_values(["Product", "Week", "Item", "OperationNo", "TaskNo"])

    for ri, row in enumerate(ops_sorted[cols].itertuples(index=False), 2):
        bg = ALT_ROW if ri % 2 == 0 else None
        for ci, (col, val) in enumerate(zip(cols, row), 1):
            fmt   = None
            align = "left"
            fill  = bg
            font  = FONT_BODY

            if col in ("Product", "Week"):
                align = "center"
            elif col in ("weekly_good", "weekly_bad", "row_count", "OperationNo", "TaskNo"):
                fmt, align = FMT_INT, "right"
            elif col == "operation_yield":
                fmt, align = FMT_PCT, "right"
                if pd.notna(val) and val < 0.95:
                    fill = WARN_FILL
            elif col in ("AccumulatedCost", "operation_scrap_cost"):
                fmt, align = FMT_MONEY, "right"
                if pd.isna(val):
                    fill, font, val = WARN_FILL, FONT_WARN, "⚠ missing"
            elif col in ("date_min", "date_max"):
                fmt = FMT_DATE

            safe_val = val if not (isinstance(val, float) and np.isnan(val)) else None
            if isinstance(val, str):
                safe_val = val
            _body_cell(ws, ri, ci, safe_val, fmt=fmt, fill=fill, font=font, align=align)

    widths = [10, 10, 16, 7, 8, 30, 8, 8, 9, 14, 14, 7, 12, 12]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def write_raw_sheet(wb: Workbook, raw: pd.DataFrame):
    ws = wb.create_sheet("RAW (source)")
    _freeze(ws)
    ws.row_dimensions[1].height = 30

    cols = ["Product", "Week", "Item", "OperationNo", "TaskNo",
            "TaskDescription", "InsertDate", "Good", "Bad"]
    available = [c for c in cols if c in raw.columns]

    for ci, col in enumerate(available, 1):
        c = ws.cell(1, ci, col)
        c.fill, c.font = HDR_LIGHT, FONT_DARK
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    for ri, row in enumerate(raw[available].itertuples(index=False), 2):
        bg = ALT_ROW if ri % 2 == 0 else None
        for ci, val in enumerate(row, 1):
            col = available[ci - 1]
            fmt   = FMT_DATE if col == "InsertDate" else (FMT_INT if col in ("Good", "Bad") else None)
            safe  = val if not (isinstance(val, float) and np.isnan(val)) else None
            _body_cell(ws, ri, ci, safe, fmt=fmt, fill=bg)

    _col_width(ws)


def write_detectors_input(wb: Workbook, prod: pd.DataFrame):
    ws = wb.create_sheet("↳ Enter Detectors Here")
    _freeze(ws)
    ws.row_dimensions[1].height = 36
    ws["A1"] = "↓ Fill in CompletedDetectors column (yellow cells) and re-run script"
    ws["A1"].font = Font(bold=True, color="7F6000", size=11)
    ws.merge_cells("A1:D1")
    ws.row_dimensions[2].height = 30

    labels = ["Product", "Week (YYYY-Www)", "Completed Detectors", "Notes"]
    for ci, lbl in enumerate(labels, 1):
        c = ws.cell(2, ci, lbl)
        c.fill, c.font = HDR_MID, FONT_WHITE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER

    week_rows = prod[["Product", "Week", "CompletedDetectors"]].drop_duplicates().sort_values(["Product", "Week"])
    for ri, row in enumerate(week_rows.itertuples(index=False), 3):
        bg = ALT_ROW if ri % 2 == 0 else None
        _body_cell(ws, ri, 1, row.Product, fill=bg, align="center")
        _body_cell(ws, ri, 2, row.Week,    fill=bg, align="center")
        has_det = not (isinstance(row.CompletedDetectors, float) and np.isnan(row.CompletedDetectors))
        c = ws.cell(ri, 3, row.CompletedDetectors if has_det else None)
        c.number_format = FMT_INT
        c.border = BORDER
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.fill = OK_FILL if has_det else PatternFill("solid", fgColor="FFFFC0")
        c.font = FONT_OK if has_det else FONT_BODY
        _body_cell(ws, ri, 4, "" , fill=bg)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 40
    ws.sheet_view.showGridLines = False


def write_report(raw, ops, comp, prod, out_path="yield_report.xlsx"):
    wb = Workbook()
    wb.remove(wb.active)

    write_product_summary(wb, prod)
    write_component_summary(wb, comp)
    write_semi_raw(wb, ops)
    write_raw_sheet(wb, raw)
    write_detectors_input(wb, prod)

    wb.save(out_path)
    print(f"✅ Report saved → {out_path}")
    missing_costs = ops["AccumulatedCost"].isna().sum()
    missing_det   = prod["CompletedDetectors"].isna().sum()
    if missing_costs:
        print(f"   ⚠  {missing_costs} operations still missing cost → fill costs.xlsx and re-run")
    if missing_det:
        print(f"   ⚠  {missing_det} week/product rows missing Completed Detectors → fill '↳ Enter Detectors Here' tab")
    if not missing_costs and not missing_det:
        print("   All data complete ✓")


# ── MAIN ────────────────────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]

    if not args:
        print(__doc__)
        sys.exit(0)

    raw_path       = args[0]
    costs_path     = args[1] if len(args) > 1 else None
    detectors_path = args[2] if len(args) > 2 else None
    out_path       = next((a[len("--output="):] for a in flags if a.startswith("--output=")),
                          "yield_report.xlsx")

    if "--template" in flags:
        generate_cost_template(raw_path, "costs_template.xlsx")
        return

    raw, ops, comp, prod = run_pipeline(raw_path, costs_path, detectors_path)
    write_report(raw, ops, comp, prod, out_path)


if __name__ == "__main__":
    main()
