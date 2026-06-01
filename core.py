"""
core.py — data processing
"""
from __future__ import annotations

import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ── File loading ────────────────────────────────────────────────────────────────────

def load_costs(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        xl = pd.ExcelFile(path)
        sheet = next((s for s in xl.sheet_names if s.lower() == "costs"), xl.sheet_names[0])
        df = xl.parse(sheet)
    else:
        df = pd.read_csv(path, sep=None, engine="python")
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Product", "Item", "OperationNo", "AccumulatedCost"]
    _check_columns(df, required, path)
    df["Product"]         = df["Product"].astype(str).str.strip()
    df["Item"]            = df["Item"].astype(str).str.strip()
    df["OperationNo"]     = pd.to_numeric(df["OperationNo"], errors="coerce")
    df["AccumulatedCost"] = df["AccumulatedCost"].apply(_parse_float)
    if "Name" not in df.columns:
        df["Name"] = ""
    df["Name"] = df["Name"].astype(str).str.strip().replace("nan", "")
    if "TaskDescription" not in df.columns:
        df["TaskDescription"] = ""
    df["TaskDescription"] = df["TaskDescription"].astype(str).str.strip()
    return df.dropna(subset=["OperationNo", "AccumulatedCost"])


def load_raw(path: str) -> pd.DataFrame:
    df = _read_file(path)
    df.columns = [str(c).strip() for c in df.columns]
    col_map = _detect_raw_columns(df)
    df = df.rename(columns={v: k for k, v in col_map.items()})
    required = ["Item", "OperationNo", "TaskNo", "TaskDescription",
                "InsertDate", "Good", "Bad"]
    df = df[[c for c in required if c in df.columns]].copy()
    df["Item"]            = df["Item"].astype(str).str.strip()
    df["OperationNo"]     = pd.to_numeric(df["OperationNo"], errors="coerce")
    df["TaskNo"]          = pd.to_numeric(df["TaskNo"], errors="coerce")
    df["TaskDescription"] = df["TaskDescription"].astype(str).str.strip()
    df["InsertDate"]      = pd.to_datetime(df["InsertDate"], dayfirst=True, errors="coerce")
    df["Good"]            = pd.to_numeric(df["Good"], errors="coerce").fillna(0).astype(int)
    df["Bad"]             = pd.to_numeric(df["Bad"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["InsertDate", "OperationNo", "TaskNo"])
    return df


def load_results(path: str) -> dict[str, pd.DataFrame]:
    empty = {
        "detail":    _empty_detail(),
        "component": _empty_component(),
        "product":   _empty_product(),
    }
    p = Path(path)
    if not p.exists():
        return empty
    DETAIL_COL_MAP = {
        "Op No":            "OperationNo",
        "Task Description": "TaskDescription",
        "Good":             "weekly_good",
        "Bad":              "weekly_bad",
        "Op Yield":         "operation_yield",
        "Accum.\nCost (€)": "AccumulatedCost",
        "Scrap\nCost (€)":  "operation_scrap_cost",
        "Rows":             "row_count",
        "Date From":        "date_min",
        "Date To":          "date_max",
    }
    PRODUCT_COL_MAP = {
        "Completed\nDetectors":  "completed_detectors",
        "Total Scrap\nCost (€)": "total_scrap_cost",
        "Scrap /\nDetector (€)": "scrap_per_detector",
        "Composite\nYield":      "composite_yield",
        "Scrap\nQty":            "total_bad_qty",
        "Status":                "_status",
    }
    COMPONENT_COL_MAP = {
        "Composite\nYield":      "composite_yield",
        "Total Scrap\nCost (€)": "total_scrap_cost",
    }

    try:
        xl = pd.ExcelFile(path)

        def _read(sheet, empty_df, col_map=None):
            if sheet not in xl.sheet_names:
                return empty_df
            df = xl.parse(sheet)
            if col_map:
                df = df.rename(columns=col_map)
            return df

        results = {
            "detail":    _read("WeeklyDetail",     _empty_detail(),    DETAIL_COL_MAP),
            "component": _read("ComponentSummary", _empty_component(), COMPONENT_COL_MAP),
            "product":   _read("ProductSummary",   _empty_product(),   PRODUCT_COL_MAP),
        }

        prod = results["product"]
        if "Week" in prod.columns and len(prod):
            week_re = re.compile(r"^\d{4}-W\d{2}$")
            mask = prod["Week"].astype(str).str.match(week_re)
            results["product"] = prod[mask].reset_index(drop=True)

        for key, cols in {
            "product":   ["composite_yield", "total_bad_qty", "total_scrap_cost",
                          "completed_detectors", "scrap_per_detector"],
            "component": ["composite_yield", "total_scrap_cost"],
        }.items():
            df_ = results[key]
            for col in cols:
                if col in df_.columns:
                    df_[col] = pd.to_numeric(df_[col], errors="coerce")

        det_sheet = next(
            (name for name in ("Detectors", "↳ Детекторы") if name in xl.sheet_names),
            None,
        )
        if det_sheet:
            det = xl.parse(det_sheet, header=1)
            det.columns = [str(c).strip() for c in det.columns]
            if all(c in det.columns for c in ["Product", "Week", "Completed Detectors"]):
                for _, row in det.iterrows():
                    try:
                        count = int(row["Completed Detectors"])
                        if count > 0:
                            results = set_completed_detectors(
                                results,
                                str(row["Product"]).strip(),
                                str(row["Week"]).strip(),
                                count,
                            )
                    except (ValueError, TypeError):
                        pass
        return results
    except Exception as e:
        raise ValueError(f"Failed to read {path}: {e}")


# ── Analysis / preview ────────────────────────────────────────────────────────────

def analyse(raw: pd.DataFrame, costs: pd.DataFrame,
            existing: dict[str, pd.DataFrame]) -> dict:
    known_items = set(costs["Item"].unique())
    product_map = dict(zip(costs["Item"], costs["Product"]))

    raw_filtered = raw[raw["Item"].isin(known_items)].copy()
    raw_filtered["Week"] = raw_filtered["InsertDate"].dt.strftime("%G-W%V")

    weeks_in_raw     = sorted(raw_filtered["Week"].dropna().unique())
    weeks_in_results = sorted(existing["detail"]["Week"].dropna().unique()
                              if "Week" in existing["detail"].columns else [])
    ignored_items    = sorted(set(raw["Item"].unique()) - known_items)

    existing_prod_weeks: set[tuple] = set()
    prod_df = existing.get("product", pd.DataFrame())
    if not prod_df.empty and "Product" in prod_df.columns and "Week" in prod_df.columns:
        existing_prod_weeks = {
            (str(r["Product"]), str(r["Week"])) for _, r in prod_df.iterrows()
        }

    weeks_to_add = []
    new_prod_weeks: set[tuple[str, str]] = set()
    for w in weeks_in_raw:
        week_products = {product_map[i]
                         for i in raw_filtered[raw_filtered["Week"] == w]["Item"].unique()
                         if i in product_map}
        if any((p, w) not in existing_prod_weeks for p in week_products):
            weeks_to_add.append(w)
        for p in week_products:
            new_prod_weeks.add((p, w))

    prod_week_pairs = existing_prod_weeks | new_prod_weeks

    return {
        "known_items":      sorted(known_items),
        "product_map":      product_map,
        "weeks_in_raw":     weeks_in_raw,
        "weeks_in_results": weeks_in_results,
        "weeks_to_add":     weeks_to_add,
        "ignored_items":    ignored_items,
        "prod_week_pairs":  prod_week_pairs,
    }


# ── Calculation ───────────────────────────────────────────────────────────────────

def process_new_weeks(raw: pd.DataFrame, costs: pd.DataFrame,
                      weeks_to_add: list[str]) -> dict[str, pd.DataFrame]:
    if not weeks_to_add:
        return {"detail": _empty_detail(), "component": _empty_component(),
                "product": _empty_product()}

    product_map = dict(zip(costs["Item"], costs["Product"]))
    known_items = set(costs["Item"])

    df = raw[raw["Item"].isin(known_items)].copy()
    df["Product"] = df["Item"].map(product_map)
    df["Week"]    = df["InsertDate"].dt.strftime("%G-W%V")
    df = df[df["Week"].isin(weeks_to_add)]

    key = ["Product", "Week", "Item", "OperationNo", "TaskDescription"]
    ops = df.groupby(key, as_index=False, dropna=False).agg(
        weekly_good=("Good", "sum"),
        weekly_bad=("Bad", "sum"),
        row_count=("Good", "count"),
        date_min=("InsertDate", "min"),
        date_max=("InsertDate", "max"),
    )

    ops = ops.merge(
        costs[["Item", "OperationNo", "AccumulatedCost", "Name"]],
        on=["Item", "OperationNo"], how="inner"
    )

    total = ops["weekly_good"] + ops["weekly_bad"]
    ops["operation_yield"]      = np.where(total > 0, ops["weekly_good"] / total, np.nan)
    ops["operation_scrap_cost"] = ops["weekly_bad"] * ops["AccumulatedCost"]

    def _composite(s):
        v = s.dropna()
        return v.prod() if len(v) else np.nan

    comp = ops.groupby(["Product", "Week", "Item", "Name"], as_index=False).agg(
        composite_yield=("operation_yield", _composite),
        total_scrap_cost=("operation_scrap_cost", "sum"),
    )

    bad_by_prod = ops.groupby(["Product", "Week"], as_index=False)["weekly_bad"].sum().rename(
        columns={"weekly_bad": "total_bad_qty"}
    )
    prod = comp.groupby(["Product", "Week"], as_index=False).agg(
        total_scrap_cost=("total_scrap_cost", "sum"),
        composite_yield=("composite_yield", "mean"),
    )
    prod = prod.merge(bad_by_prod, on=["Product", "Week"], how="left")
    prod["completed_detectors"] = np.nan
    prod["scrap_per_detector"]  = np.nan

    missing = ops[ops["AccumulatedCost"].isna()][
        ["Item", "OperationNo", "TaskDescription"]
    ].drop_duplicates()

    return {
        "detail":        ops,
        "component":     comp,
        "product":       prod,
        "missing_costs": missing,
    }


def merge_results(existing: dict[str, pd.DataFrame],
                  new: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    _dedup_keys = {
        "detail":    ["Product", "Week", "Item", "OperationNo"],
        "component": ["Product", "Week", "Item"],
        "product":   ["Product", "Week"],
    }

    def _concat(key):
        old = existing[key]
        nw  = new.get(key, pd.DataFrame())
        if nw.empty:
            return old
        combined = pd.concat([old, nw], ignore_index=True)
        keys = [k for k in _dedup_keys.get(key, []) if k in combined.columns]
        if keys:
            combined = combined.drop_duplicates(subset=keys, keep="first")
        return combined

    return {
        "detail":    _concat("detail"),
        "component": _concat("component"),
        "product":   _concat("product"),
    }


def set_completed_detectors(results: dict[str, pd.DataFrame],
                             product: str, week: str, count: int
                             ) -> dict[str, pd.DataFrame]:
    prod = results["product"].copy()
    for col in ("total_scrap_cost", "completed_detectors", "scrap_per_detector"):
        if col in prod.columns:
            prod[col] = pd.to_numeric(prod[col], errors="coerce")
    mask = (prod["Product"] == product) & (prod["Week"] == week)
    if mask.sum() == 0:
        raise ValueError(f"No data for {product} / {week}")
    prod.loc[mask, "completed_detectors"] = count
    prod.loc[mask, "scrap_per_detector"]  = (
        prod.loc[mask, "total_scrap_cost"] / count if count > 0 else np.nan
    )
    results["product"] = prod
    return results


def load_detectors(path: str,
                   products: list[str] | None = None) -> list[tuple[str, str, int]]:
    p = Path(path)
    if p.suffix.lower() not in (".xlsx", ".xls"):
        return []
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return []
    sheet = next((s for s in xl.sheet_names if s.lower() in ("detectors", "detector")), None)
    if not sheet:
        return []

    df = xl.parse(sheet)
    df.columns = [str(c).strip() for c in df.columns]
    if df.shape[1] < 2 or df.empty:
        return []

    lower = {c.lower(): c for c in df.columns}
    pcol = lower.get("product")
    wcol = lower.get("week")
    ccol = next((lower[k] for k in
                 ("completed detectors", "completeddetectors",
                  "detectors", "count", "completed") if k in lower), None)

    out: list[tuple[str, str, int]] = []

    if pcol and wcol and ccol:
        for _, r in df.iterrows():
            product = str(r[pcol]).strip()
            week    = str(r[wcol]).strip()
            try:
                count = int(float(r[ccol]))
            except (ValueError, TypeError):
                continue
            if product and week and product.lower() != "nan" and count > 0:
                out.append((product, week, count))
        return out

    week_col  = df.columns[0]
    prod_cols = list(df.columns[1:])

    def _match_product(header: str) -> str | None:
        if not products:
            return str(header).strip()
        tokens = re.split(r"[\s,/]+", str(header).strip())
        for tok in tokens:
            for prod in products:
                if tok.lower() == prod.lower():
                    return prod
        for prod in products:
            if prod.lower() in str(header).lower():
                return prod
        return None

    col_product = {c: _match_product(c) for c in prod_cols}

    for _, r in df.iterrows():
        week = str(r[week_col]).strip()
        if not week or week.lower() == "nan":
            continue
        for c in prod_cols:
            product = col_product[c]
            if not product:
                continue
            try:
                count = int(float(r[c]))
            except (ValueError, TypeError):
                continue
            if count > 0:
                out.append((product, week, count))
    return out


def import_raw_from_source(source_path: str) -> tuple[pd.DataFrame, str]:
    """Read source file and return rows for the previous ISO week."""
    from datetime import date, timedelta
    today = date.today()
    prev  = today - timedelta(days=7)
    yr, wk, _ = prev.isocalendar()
    target = f"{yr}-W{wk:02d}"
    df = load_raw(source_path)
    df["_week"] = df["InsertDate"].dt.strftime("%G-W%V")
    out = df[df["_week"] == target].drop(columns=["_week"]).reset_index(drop=True)
    return out, target


def apply_detectors(results: dict[str, pd.DataFrame],
                    rows: list[tuple[str, str, int]]) -> tuple[dict[str, pd.DataFrame], int]:
    prod = results.get("product")
    if prod is None or prod.empty or "Week" not in prod.columns:
        return results, 0
    weeks = set(prod["Week"].astype(str))

    applied = 0
    for product, week, count in rows:
        wk = week
        if wk not in weeks:
            m = re.search(r"(\d+)\s*$", wk)
            suffix = f"W{int(m.group(1)):02d}" if m else wk
            cands = [w for w in weeks if w == suffix or w.endswith("-" + suffix)]
            if len(cands) == 1:
                wk = cands[0]
        try:
            results = set_completed_detectors(results, product, wk, count)
            applied += 1
        except ValueError:
            continue
    return results, applied


# ── Helpers ─────────────────────────────────────────────────────────────────────────

def _read_file(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(path, sep=None, engine="python")
    return pd.read_excel(path)


def _parse_float(val) -> float:
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "")
    if re.match(r"^-?\d+,\d+$", s):
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def _check_columns(df: pd.DataFrame, required: list[str], path: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing columns in {Path(path).name}: {missing}\n"
            f"Found: {list(df.columns)}"
        )


def _detect_raw_columns(df: pd.DataFrame) -> dict[str, str]:
    col_map: dict[str, str] = {}
    for col in df.columns:
        low = col.lower().replace(" ", "").replace("_", "")
        if "item" in low and "Item" not in col_map:
            col_map["Item"] = col
        elif "operationno" in low and "OperationNo" not in col_map:
            col_map["OperationNo"] = col
        elif "taskno" in low and "TaskNo" not in col_map:
            col_map["TaskNo"] = col
        elif "taskdesc" in low and "TaskDescription" not in col_map:
            col_map["TaskDescription"] = col
        elif "insertdate" in low and "InsertDate" not in col_map:
            col_map["InsertDate"] = col
        elif col.lower() == "good" and "Good" not in col_map:
            col_map["Good"] = col
        elif col.lower() == "bad" and "Bad" not in col_map:
            col_map["Bad"] = col

    needed = ["Item", "OperationNo", "TaskNo", "TaskDescription",
              "InsertDate", "Good", "Bad"]
    missing = [k for k in needed if k not in col_map]
    if missing:
        cols = list(df.columns)
        if len(cols) >= 7:
            positions = ["Item", "OperationNo", "TaskNo", "TaskDescription", "InsertDate"]
            if len(cols) == 8:
                positions += ["_skip", "Good", "Bad"]
            else:
                positions += ["Good", "Bad"]
            for i, name in enumerate(positions):
                if name != "_skip" and name not in col_map:
                    col_map[name] = cols[i]
    return col_map


def _empty_detail() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "Product", "Week", "Item", "Name", "OperationNo", "TaskDescription",
        "weekly_good", "weekly_bad", "operation_yield", "AccumulatedCost",
        "operation_scrap_cost", "row_count", "date_min", "date_max",
    ])


def _empty_component() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "Product", "Week", "Item", "Name", "composite_yield", "total_scrap_cost",
    ])


def _empty_product() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "Product", "Week", "total_scrap_cost", "composite_yield",
        "total_bad_qty", "completed_detectors", "scrap_per_detector",
    ])
