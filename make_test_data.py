"""Generate test raw data and costs file from the uploaded Excel files."""
import pandas as pd
import numpy as np

# ── Reconstruct raw data from the user's sample ─────────────────────────────────────
raw_rows = [
    # Item        OpNo  TaskNo  TaskDesc                          InsertDate    Good   Bad
    ("IP2A295",   120,  2270,   "CONDUCTOR PRINTING A-SIDE",     "8.5.2026",   6600,  0),
    ("SH2A295",   130,  1810,   "VISUAL CONTROL",                "27.4.2026",  196,   0),
    ("SH2A295",   130,  1810,   "VISUAL CONTROL",                "6.5.2026",   55,    0),
    ("SH2A295",   130,  1810,   "VISUAL CONTROL",                "8.5.2026",   37,    0),
    ("SH2A295",   130,  1810,   "VISUAL CONTROL",                "12.5.2026",  673,   1),
    ("IP2A295",   130,  2270,   "CONDUCTOR PRINTING A-SIDE",     "8.5.2026",   6600,  0),
    ("IP2A295",   140,  2270,   "CONDUCTOR PRINTING A-SIDE",     "8.5.2026",   6600,  0),
    ("IP2A295",   150,  2270,   "CONDUCTOR PRINTING A-SIDE",     "8.5.2026",   6600,  0),
    ("SH2A295",   160,  1650,   "PACKAGE",                       "6.5.2026",   55,    0),
    ("SH2A295",   160,  1650,   "PACKAGE",                       "12.5.2026",  984,   4),
    ("SH2A295",   160,  1650,   "PACKAGE",                       "13.5.2026",  436,   2),
    ("IP2A295",   160,  2270,   "CONDUCTOR PRINTING A-SIDE",     "11.5.2026",  6600,  0),
    ("IP2A295",   170,  2260,   "CONDUCTOR PRINTING B-SIDE",     "11.5.2026",  6600,  0),
    ("IP2A295",   180,  2260,   "CONDUCTOR PRINTING B-SIDE",     "11.5.2026",  6600,  0),
    # Additional synthetic rows for IH2A295 and SH2A295
    ("IH2A295",    60,  1320,   "PLASTIC GLUING",                "5.5.2026",   2222,  179),
    ("IH2A295",    60,  1320,   "PLASTIC GLUING",                "12.5.2026",  1800,  20),
    ("IH2A295",   100,  1330,   "WIRE BONDING",                  "5.5.2026",   2000,  0),
    ("IH2A295",   100,  1330,   "WIRE BONDING",                  "12.5.2026",  1780,  2),
    ("SH2A295",    20,  1800,   "PLASTIC DRYING",                "28.4.2026",  5000,  0),
    ("SH2A295",    25,  1810,   "PLASTIC GLUING",                "28.4.2026",  4800,  20),
    ("SH2A295",   110,  1820,   "TESTING",                       "5.5.2026",   960,   15),
    ("SH2A295",   110,  1820,   "TESTING",                       "12.5.2026",  1400,  28),
]

raw = pd.DataFrame(raw_rows, columns=["Item", "OperationNo", "TaskNo", "TaskDescription",
                                       "InsertDate", "Good", "Bad"])
raw["InsertDate"] = pd.to_datetime(raw["InsertDate"], dayfirst=True)

# ── Costs extracted from the Shallow Excel (costs.xlsx) ───────────────────────────────────
cost_rows = [
    # Product   Item       Name                  OpNo  TaskNo  TaskDesc                        AccumCost
    ("Shallow","IP2A295", "IP board",             120,  2270,   "CONDUCTOR PRINTING A-SIDE",    2.20),
    ("Shallow","IP2A295", "IP board",             130,  2270,   "CONDUCTOR PRINTING A-SIDE",    2.28),
    ("Shallow","IP2A295", "IP board",             140,  2270,   "CONDUCTOR PRINTING A-SIDE",    2.42),
    ("Shallow","IP2A295", "IP board",             150,  2270,   "CONDUCTOR PRINTING A-SIDE",    2.48),
    ("Shallow","IP2A295", "IP board",             160,  2270,   "CONDUCTOR PRINTING A-SIDE",    2.60),
    ("Shallow","IP2A295", "IP board",             170,  2260,   "CONDUCTOR PRINTING B-SIDE",    2.70),
    ("Shallow","IP2A295", "IP board",             180,  2260,   "CONDUCTOR PRINTING B-SIDE",    2.77),
    ("Shallow","IH2A295", "Base assembly",         60,  1320,   "PLASTIC GLUING",               3.61),
    ("Shallow","IH2A295", "Base assembly",        100,  1330,   "WIRE BONDING",                 3.77),
    ("Shallow","SH2A295", "Shallow detector",      20,  1800,   "PLASTIC DRYING",               7.37),
    ("Shallow","SH2A295", "Shallow detector",      25,  1810,   "PLASTIC GLUING",               8.37),
    ("Shallow","SH2A295", "Shallow detector",     110,  1820,   "TESTING",                     11.88),
    ("Shallow","SH2A295", "Shallow detector",     130,  1810,   "VISUAL CONTROL",              12.10),
    ("Shallow","SH2A295", "Shallow detector",     160,  1650,   "PACKAGE",                     12.35),
]

costs = pd.DataFrame(cost_rows, columns=["Product", "Item", "Name", "OperationNo", "TaskNo",
                                          "TaskDescription", "AccumulatedCost"])

# ── Completed detectors — matrix layout: week rows, product columns ───────────────────────
detectors = pd.DataFrame({
    "Week":            ["2026-W18", "2026-W19", "2026-W20"],
    "Shallow SH2A295": [1537, 748, 1162],
})

# ── Write a single workbook: raw data + Costs + Detectors sheets ───────────────────────
with pd.ExcelWriter("test_raw.xlsx") as xl:
    raw.to_excel(xl, sheet_name="Raw", index=False)
    costs.to_excel(xl, sheet_name="Costs", index=False)
    detectors.to_excel(xl, sheet_name="Detectors", index=False)
print(f"Saved test_raw.xlsx — Raw ({len(raw)}), Costs ({len(costs)}), "
      f"Detectors ({len(detectors)})")
