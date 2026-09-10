#!/usr/bin/env python3
"""Set $750M buybacks and replace ugly WC float literals with FY25 formulas.

Run:  cd LULU && python3 scripts/fix_buybacks_and_decimals.py
"""
from __future__ import annotations

import math
import re
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGETS = [
    ROOT / "model18_wsp_formulas.xlsx",
    ROOT / "model18_humanized.xlsx",
    ROOT / "model18_humanized (1).xlsx",
]

BUYBACK_K = 750_000
DAY_FMT = "0.0"
PCT_FMT = "0.0%"

# Scenarios WC assumption rows (cols F,G,H)
SCN_WC_FORMULAS: dict[int, tuple[str, str]] = {
    15: ("=(DCF!$E$21/DCF!$E$5)*365", DAY_FMT),
    16: ("=(DCF!$E$23/DCF!$E$8)*365", DAY_FMT),
    18: ("=(DCF!$E$25/DCF!$E$8)*365", DAY_FMT),
    19: ("=DCF!$E$27/DCF!$E$5", PCT_FMT),
    20: ("=DCF!$E$29/DCF!$E$5", PCT_FMT),
}

# DCF FY25A WC driver rows (col E)
DCF_FY25_FORMULAS: dict[int, tuple[str, str]] = {
    7: ("=(E5-E8)/E5", PCT_FMT),
    12: ("=E11/E5", PCT_FMT),
    20: ("=(E21/E5)*365", DAY_FMT),
    22: ("=(E23/E8)*365", DAY_FMT),
    24: ("=(E25/E8)*365", DAY_FMT),
    26: ("=E27/E5", PCT_FMT),
    28: ("=E29/E5", PCT_FMT),
}


def _is_ugly_float(val) -> bool:
    if not isinstance(val, float) or isinstance(val, bool):
        return False
    if not math.isfinite(val):
        return False
    s = f"{val:.12f}".rstrip("0")
    return "." in s and len(s.split(".")[1]) > 4


def fix_workbook(path: Path) -> dict[str, int]:
    wb = openpyxl.load_workbook(path)
    stats = {"buyback": 0, "wc_formulas": 0, "rounded": 0}

    scn = wb["Scenarios"]
    for col in ("B", "C"):
        if scn[f"{col}21"].value != BUYBACK_K:
            scn[f"{col}21"] = BUYBACK_K
            stats["buyback"] += 1

    dcf = wb["DCF"]
    want_budget = "=Scenarios!$C$21"
    if dcf["E65"].value != want_budget:
        dcf["E65"] = want_budget
        dcf["E65"].number_format = "#,##0;(#,##0)"
        stats["buyback"] += 1

    for row, (formula, fmt) in SCN_WC_FORMULAS.items():
        for col in ("F", "G", "H"):
            cell = scn[f"{col}{row}"]
            if cell.value != formula:
                cell.value = formula
                stats["wc_formulas"] += 1
            cell.number_format = fmt

    for row, (formula, fmt) in DCF_FY25_FORMULAS.items():
        cell = dcf[f"E{row}"]
        if cell.value != formula:
            cell.value = formula
            stats["wc_formulas"] += 1
        cell.number_format = fmt

    # FY25 accretive EPS — link to NI ÷ ending shares, not a pasted float
    eps_formula = "=E77/E75"
    if dcf["E78"].value != eps_formula:
        dcf["E78"] = eps_formula
        stats["wc_formulas"] += 1
    dcf["E78"].number_format = "$#,##0.00"

    # Round any remaining ugly floats on label rows (margin %, tax, etc.)
    ugly_pat = re.compile(r"0\.\d{5,}")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            label = str(row[0].value or "").lower()
            for cell in row[1:]:
                if not _is_ugly_float(cell.value):
                    continue
                if any(k in label for k in ("dso", "dio", "dpo", "days")):
                    cell.number_format = DAY_FMT
                elif any(
                    k in label
                    for k in ("margin", "rate", "growth", "tax", "weight", "mix", "erp", "%")
                ):
                    cell.number_format = PCT_FMT
                    cell.value = round(cell.value, 3)
                stats["rounded"] += 1

    wb.save(path)
    return stats


if __name__ == "__main__":
    for path in TARGETS:
        if not path.exists():
            print(f"Skip missing: {path.name}")
            continue
        stats = fix_workbook(path)
        print(f"{path.name}: {stats}")
