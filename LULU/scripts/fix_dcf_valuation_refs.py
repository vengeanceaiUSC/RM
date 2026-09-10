#!/usr/bin/env python3
"""Fix broken DCF valuation refs after Scenarios column B/C delete.

Stale Scenarios!$B$* cells are empty (old Notes col); base case lives in col D.
Also repairs Gordon growth block row/column refs shifted by the growth-driver insert.

Run:  cd LULU && python3 scripts/fix_dcf_valuation_refs.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

SCN_BASE = "Scenarios!D"
WACC = "WACC!B36"

# Gordon growth block (final layout after growth-driver + FY25 D&A inserts)
VAL = {
    "sum_pv": ("B43", "=SUM(C39:G39)"),
    "terminal_g": ("B44", f"={SCN_BASE}10"),
    "terminal_ebitda": ("B45", "=G11+G17"),
    "terminal_value": (
        "B46",
        f"=G36*(1+{SCN_BASE}10)/({WACC}-{SCN_BASE}10)",
    ),
    "exit_multiple": ("B47", "=B46/B45"),
    "tv_memo": (
        "B48",
        f"=G36/B45*(1+{SCN_BASE}10)/({WACC}-{SCN_BASE}10)",
    ),
    "pv_terminal": ("B49", f"=B46/(1+{WACC})^G37"),
    "enterprise_value": ("B50", "=B43+B49"),
    "equity_value": ("B55", "=B50+B51+B52"),
    "implied_price": ("B57", "=B55/B56"),
    "upside": ("B59", "=B57/B58-1"),
    "tv_weight": ("B60", "=B49/B50"),
    "exit_multiple_input": ("B82", "=B47"),
    "exit_tv_gordon": ("B83", "=B47"),
    "exit_tv_multiple": ("B84", "=B45*B82"),
    "exit_pv_tv": ("B85", f"=B83/(1+{WACC})^G37"),
    "exit_ev": ("B86", "=B43+B85"),
    "exit_price": ("B87", "=(B85+B51+B52)/B56"),
}


def _rewrite_scenarios_base(formula: str) -> str:
    return formula.replace("Scenarios!$B$", SCN_BASE)


def fix(path: Path = TARGET) -> dict:
    tmp = path.with_suffix(".valfix.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"scenarios_b_to_d": 0, "valuation_cells": 0, "dcf_rows": 0}

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not isinstance(val, str) or not val.startswith("="):
                    continue
                new_val = _rewrite_scenarios_base(val)
                if new_val != val:
                    cell.value = new_val
                    stats["scenarios_b_to_d"] += 1

    dcf = wb["DCF"]

    # Discount factor + PV of explicit FCF
    for col in ("C", "D", "E", "F", "G"):
        dcf[f"{col}38"] = f"=1/(1+{WACC})^{col}37"
        dcf[f"{col}39"] = f"={col}36*{col}38"
    stats["dcf_rows"] += 2

    for _, (addr, formula) in VAL.items():
        dcf[addr] = formula
        stats["valuation_cells"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"DCF valuation ref fix: {fix()}")
