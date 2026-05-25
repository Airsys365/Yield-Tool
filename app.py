"""
app.py — Yield & Scrap Cost Tool GUI
Run: python app.py
"""
from __future__ import annotations

import json
import threading
import time
import traceback
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

import core
import excel_writer

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

CONFIG_FILE     = Path(__file__).parent / "config.json"
RESULTS_DEFAULT = Path(__file__).parent / "results.xlsx"

ACCENT       = "#2E75B6"
ACCENT_HOVER = "#1F5496"
ACCENT_FAINT = "#E8F0F8"
BORDER       = "#D0D7DE"
MUTED        = "#6E7781"
OK_COLOR     = "#1A7F37"
ERR_COLOR    = "#C0392B"


HELP_TEXT = """Yield & Scrap Cost Tool

WHAT IT DOES
Reads a weekly production export and computes yield and scrap cost per
week, then appends the results to results.xlsx. Existing weeks are
never recalculated — historical scrap costs stay locked at the prices
that were active when the week was first processed.

THE RAW FILE HOLDS EVERYTHING (single source of truth)
The raw .xlsx file contains three sheets:

1. Raw data (first sheet) — the database export.
   Required columns: Item, OperationNo, TaskNo, TaskDescription,
   InsertDate, Good, Bad. Column names are matched by keyword, case
   does not matter.

2. Costs — the cost table.
   Required columns: Product, Item, OperationNo, TaskNo,
   AccumulatedCost. Optional: TaskDescription. European decimal
   commas (3,42) are accepted.

3. Detectors — completed detector counts. Matrix layout: the first
   column is the Week, every other column header names a product, and
   the cells hold the counts:

       Week      | Gamma SH2A280 | Shallow SH2A295
       2026-W18  |               | 1230
       2026-W20  | 1015          |

   Week is the calendar week — W18 or 2026-W18. A column header may
   carry an extra item code (e.g. "Shallow SH2A295"); the product name
   inside it is what gets matched. This sheet is optional; without it
   the "Scrap per Detector" column simply stays empty.

Items present in raw but missing from Costs are dropped (listed under
"Ignored" in the preview).

Keeping detectors in the raw file means you never re-type a year of
counts: if the results path changes or the file is regenerated from
scratch, the detector numbers are read straight back from the raw
file.

WORKFLOW
1. Pick the raw file and the results.xlsx path. If results.xlsx does
   not exist yet it will be created on the first run.

2. Click "Check files". The preview shows which products and items
   were found, which weeks already live in results.xlsx, and which
   weeks are about to be added.

3. Click "Add new weeks". New weeks are calculated and detector
   counts from the raw file's Detectors sheet are applied. If there
   are no new weeks the button becomes "Refresh detector counts" and
   re-applies the detector numbers to the existing weeks.

OUTPUT SHEETS (results.xlsx)
- ProductSummary    overall yield, scrap qty, scrap cost, scrap per
                   detector per product/week, plus a cumulative
                   summary across all weeks.
- ComponentSummary  same metrics broken down by component (item).
- WeeklyDetail      row per operation per week. Yellow highlight:
                   yield below 95% or missing price.
- Charts            weekly trend charts: scrap cost, scrap per
                   detector, and one yield chart per product.
- Detectors         the detector counts that were applied.

WEEK FORMAT
ISO week, YYYY-Www (for example 2026-W18). On chart X-axes only "W18"
is shown for readability.

LOCKED PRICES
Once a week is written to results.xlsx its prices and scrap costs are
frozen. If you later edit AccumulatedCost in the Costs sheet those
changes apply only to weeks added afterwards. To recalculate a
historical week, delete its rows from WeeklyDetail, ComponentSummary
and ProductSummary, then run the tool again.

TROUBLESHOOTING
The Log panel on the right shows the full traceback when something
fails. Common causes: missing Costs sheet, mismatched column names,
or a product renamed in Costs without updating historical rows in
results.xlsx.
"""


def _load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        scale, win_w, win_h = self._compute_scaling()
        self._scale = scale
        ctk.set_widget_scaling(scale)
        ctk.set_window_scaling(scale)

        self.title("Yield & Scrap Cost Tool")
        self._apply_geometry(win_w, win_h)
        self.minsize(820, 480)
        self.resizable(True, True)

        self._cfg      = _load_config()
        self._analysis: dict | None = None

        self._build_ui()
        self._restore_paths()

    def _compute_scaling(self) -> tuple[float, int, int]:
        DESIGN_W, DESIGN_H = 980, 560
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        usable_w = sw * 0.90
        usable_h = sh * 0.90
        scale = min(usable_w / DESIGN_W, usable_h / DESIGN_H, 1.0)
        scale = max(0.65, min(1.0, scale))
        win_w = min(DESIGN_W, int(usable_w))
        win_h = min(DESIGN_H, int(usable_h))
        return scale, win_w, win_h

    def _apply_geometry(self, w: int, h: int):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, sh // 3 - h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── UI ────────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        # Two panes: left (controls + preview) | sep | right (log)
        self.grid_columnconfigure(0, weight=2, minsize=int(420 * self._scale))
        self.grid_columnconfigure(1, weight=0, minsize=1)
        self.grid_columnconfigure(2, weight=3)

        self._build_header()

        tk.Frame(self, bg=BORDER, width=1).grid(
            row=1, column=1, sticky="ns", pady=(0, 12))

        self._build_left()
        self._build_right()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=3, sticky="ew", padx=16, pady=(12, 8))
        hdr.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(hdr, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(left, text="Yield & Scrap Cost Tool",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkLabel(
            left,
            text="   Appends new weeks to results.xlsx — existing weeks are never recalculated",
            text_color=MUTED,
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            hdr, text="Help", width=72, height=28,
            command=self._show_help,
            fg_color="transparent", border_width=1,
            text_color=ACCENT, border_color=ACCENT,
            hover_color=ACCENT_FAINT,
        ).grid(row=0, column=1, sticky="e")

    def _build_left(self):
        left = ctk.CTkFrame(self)
        left.grid(row=1, column=0, sticky="nsew", padx=(12, 4), pady=(0, 12))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(4, weight=1)  # preview expands

        # Files block
        ff = ctk.CTkFrame(left, fg_color="transparent")
        ff.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        ff.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(ff, text="Files",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=4, pady=(4, 4))

        self._raw_var     = ctk.StringVar()
        self._results_var = ctk.StringVar(value=str(RESULTS_DEFAULT))

        self._file_row(ff, "Raw data:", self._raw_var, self._pick_raw,
                       "DB export  (Costs + Detectors sheets)", row=1)
        self._file_row(ff, "Results:",  self._results_var, self._pick_results,
                       "results.xlsx", row=2)

        # Buttons
        ctk.CTkButton(
            left, text="Check files", command=self._run_analysis, height=32,
            fg_color="transparent", border_width=1,
            text_color=ACCENT, border_color=ACCENT,
            hover_color=ACCENT_FAINT,
        ).grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 4))

        self._run_btn = ctk.CTkButton(
            left, text="Add new weeks to results.xlsx",
            command=self._run_update, height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            state="disabled",
        )
        self._run_btn.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))

        # Preview (expands)
        pf = ctk.CTkFrame(left, fg_color="transparent")
        pf.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0, 8))
        pf.grid_columnconfigure(0, weight=1)
        pf.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(pf, text="Preview",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=4, pady=(4, 4))

        self._preview = ctk.CTkTextbox(
            pf, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", wrap="word",
            border_width=1, border_color=BORDER,
        )
        self._preview.grid(row=1, column=0, sticky="nsew")

    def _build_right(self):
        lf = ctk.CTkFrame(self)
        lf.grid(row=1, column=2, sticky="nsew", padx=(4, 12), pady=(0, 12))
        lf.grid_columnconfigure(0, weight=1)
        lf.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(lf, text="Log",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4))

        self._log = ctk.CTkTextbox(
            lf, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", wrap="word",
            border_width=1, border_color=BORDER,
        )
        self._log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        t = self._log._textbox  # type: ignore[attr-defined]
        t.tag_config("ts",   foreground=MUTED)
        t.tag_config("info", foreground="#1F2328")
        t.tag_config("ok",   foreground=OK_COLOR, font=("Consolas", 11, "bold"))
        t.tag_config("err",  foreground=ERR_COLOR)
        t.tag_config("warn", foreground="#9A6700")

    def _file_row(self, parent, label, var, pick_cmd, hint, row):
        ctk.CTkLabel(parent, text=label, width=70, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(4, 4), pady=3)
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.grid(row=row, column=0, sticky="ew", padx=(78, 4), pady=3)
        wrap.grid_columnconfigure(0, weight=1)
        e = ctk.CTkEntry(wrap, textvariable=var, placeholder_text=hint)
        e.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        e.bind("<FocusOut>", lambda _: self._clear_analysis())
        ctk.CTkButton(wrap, text="…", width=32, command=pick_cmd,
                      fg_color="transparent", border_width=1,
                      text_color=ACCENT, border_color=ACCENT,
                      hover_color=ACCENT_FAINT).grid(row=0, column=1)

    # ── Help dialog ──────────────────────────────────────────────────────────────────

    def _show_help(self):
        win = ctk.CTkToplevel(self)
        win.title("Help — Yield & Scrap Cost Tool")
        win.transient(self)
        win.geometry("680x580")
        win.minsize(520, 400)

        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(0, weight=1)

        box = ctk.CTkTextbox(
            win, font=ctk.CTkFont(family="Segoe UI", size=12),
            wrap="word", border_width=1, border_color=BORDER,
        )
        box.grid(row=0, column=0, sticky="nsew", padx=14, pady=(14, 8))
        box.insert("end", HELP_TEXT)
        box.configure(state="disabled")

        ctk.CTkButton(
            win, text="Close", width=96, command=win.destroy,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
        ).grid(row=1, column=0, sticky="e", padx=14, pady=(0, 12))

        win.bind("<Escape>", lambda _e: win.destroy())
        win.after(50, win.focus_set)
        win.grab_set()

    # ── File dialogs ─────────────────────────────────────────────────────────────

    def _pick_raw(self):
        p = filedialog.askopenfilename(
            title="Raw data file",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
        if p:
            self._raw_var.set(p)
            self._clear_analysis()

    def _pick_results(self):
        p = filedialog.asksaveasfilename(
            title="Results file",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if p:
            self._results_var.set(p)
            self._clear_analysis()

    # ── Analysis ────────────────────────────────────────────────────────────────────

    def _run_analysis(self):
        self._clear_analysis()
        raw_path     = self._raw_var.get().strip()
        results_path = self._results_var.get().strip()

        if not raw_path:
            messagebox.showwarning("No file", "Select a raw data file")
            return

        self._set_preview("Reading files...")

        def _work():
            try:
                raw      = core.load_raw(raw_path)
                costs    = core.load_costs(raw_path)
                existing = core.load_results(results_path)
                info     = core.analyse(raw, costs, existing)
                products = sorted(set(info["product_map"].values()))
                info["detector_rows"] = core.load_detectors(raw_path, products)
                self.after(0, lambda: self._show_analysis(info))
            except Exception as e:
                msg = f"Error:\n{e}"
                self.after(0, lambda: self._set_preview(msg))

        threading.Thread(target=_work, daemon=True).start()

    def _show_analysis(self, info: dict):
        self._analysis = info
        lines = []
        lines.append(f"  Items: {len(info['known_items'])}   Products: "
                     f"{', '.join(sorted(set(info['product_map'].values())))}")
        if info["ignored_items"]:
            lines.append(f"  Ignored: {', '.join(info['ignored_items'][:5])}"
                         f"{'...' if len(info['ignored_items']) > 5 else ''}")
        lines.append(f"  In raw data:  {', '.join(info['weeks_in_raw']) or '—'}")
        lines.append(f"  In results:   {', '.join(info['weeks_in_results']) or '—'}")

        det_rows = info.get("detector_rows", [])
        lines.append(f"  Detectors sheet: {len(det_rows)} row(s)"
                     if det_rows else "  Detectors sheet: none")

        results_path = self._results_var.get().strip()
        results_exist = Path(results_path).exists() if results_path else False

        if info["weeks_to_add"]:
            lines.append(f"  To add:       {', '.join(info['weeks_to_add'])}")
            btn_text, btn_on = "Add new weeks to results.xlsx", True
        elif results_exist:
            lines.append("  No new weeks — use the button below to refresh detector counts")
            btn_text, btn_on = "Refresh detector counts", True
        else:
            lines.append("  No new weeks to add")
            btn_text, btn_on = "Add new weeks to results.xlsx", False

        self._run_btn.configure(state="normal" if btn_on else "disabled", text=btn_text)
        self._set_preview("\n".join(lines))
        self._save_current_paths()

    def _clear_analysis(self, *_):
        self._analysis = None
        self._run_btn.configure(state="disabled", text="Add new weeks to results.xlsx")
        self._set_preview("")

    # ── Add weeks + apply detectors ───────────────────────────────────────────────

    def _run_update(self):
        if not self._analysis:
            return

        weeks_to_add = self._analysis.get("weeks_to_add", [])
        raw_path     = self._raw_var.get().strip()
        results_path = self._results_var.get().strip()

        if not weeks_to_add and not Path(results_path).exists():
            messagebox.showwarning("No file",
                                   "results.xlsx not found.\nAdd weeks first.")
            return

        self._run_btn.configure(state="disabled", text="Working...")
        self._log_clear()

        def _work():
            try:
                self._log_add("Reading files...", "info")
                existing = core.load_results(results_path)

                if weeks_to_add:
                    raw   = core.load_raw(raw_path)
                    costs = core.load_costs(raw_path)
                    self._log_add(f"Processing weeks: {', '.join(weeks_to_add)}", "info")
                    new    = core.process_new_weeks(raw, costs, weeks_to_add)
                    merged = core.merge_results(existing, new)
                    self._log_add(f"Weeks added: {len(weeks_to_add)}", "info")
                    self._log_add(f"Rows in WeeklyDetail: {len(merged['detail'])}", "info")
                else:
                    merged = existing
                    self._log_add("No new weeks — refreshing detector counts only", "info")

                products = sorted(merged["product"]["Product"].dropna().astype(str).unique())
                det_rows = core.load_detectors(raw_path, products)
                merged, det_applied = core.apply_detectors(merged, det_rows)
                if det_rows:
                    self._log_add(
                        f"Detectors: {det_applied} of {len(det_rows)} row(s) applied", "info")
                else:
                    self._log_add("No Detectors sheet in raw file", "warn")

                self._log_add("Saving results.xlsx...", "info")
                excel_writer.save_results(results_path, merged)

                missing_det = merged["product"]["completed_detectors"].isna().sum()
                if missing_det:
                    self._log_add(
                        f"{missing_det} product/week rows still without detectors "
                        f"— add them to the Detectors sheet", "warn")

                self._log_add("Done.", "ok")
                self.after(0, lambda: self._run_btn.configure(
                    state="disabled", text="Add new weeks to results.xlsx"))
                self.after(0, self._clear_analysis)

            except Exception as e:
                tb = traceback.format_exc()
                self._log_add(f"Error: {e}", "err")
                self._log_add(tb, "err")
                self.after(0, lambda: self._run_btn.configure(
                    state="normal", text="Add new weeks to results.xlsx"))

        threading.Thread(target=_work, daemon=True).start()

    # ── Helpers ─────────────────────────────────────────────────────────────────────────

    def _set_preview(self, text: str):
        self._preview.configure(state="normal")
        self._preview.delete("1.0", "end")
        self._preview.insert("end", text)
        self._preview.configure(state="disabled")

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _log_add(self, text: str, kind: str = "info"):
        def _do():
            self._log.configure(state="normal")
            t = self._log._textbox  # type: ignore[attr-defined]
            ts = time.strftime("[%H:%M:%S]  ")
            t.insert("end", ts, ("ts",))
            t.insert("end", text + "\n", (kind,))
            t.see("end")
            self._log.configure(state="disabled")
        self.after(0, _do)

    def _restore_paths(self):
        if "raw"     in self._cfg: self._raw_var.set(self._cfg["raw"])
        if "results" in self._cfg: self._results_var.set(self._cfg["results"])

    def _save_current_paths(self):
        self._cfg.update({"raw": self._raw_var.get(),
                          "results": self._results_var.get()})
        _save_config(self._cfg)


if __name__ == "__main__":
    app = App()
    app.mainloop()
