#!/usr/bin/env python3
"""Repair stale cross-sheet refs that zero out Scenarios / DCF / Comps.

After the growth-driver row insert, FY25 net revenue moved B5→B6 but 58
Scenarios formulas still anchor on empty DCF!$B$5. Sensitivity grid and a
few Comps / exit-multiple links also point at wrong rows.

Rewrites formula strings only — no equation logic deleted.

Run:  cd LULU && python3 scripts/fix_stale_formula_refs.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

FY25_REV = "DCF!$B$6"

# Global string replacements (order matters for overlapping patterns)
GLOBAL_REWRITES: list[tuple[str, str]] = [
    ("DCF!$B$5", FY25_REV),
    ("DCF!$D$5", FY25_REV),
]

# DCF sensitivity grid: unaltered F35:J35 / J35 / E50+E51 / E55
# → final C36:G36 / G36 / B51+B52 / B56 after growth-driver insert
SENS_REWRITES: list[tuple[str, str]] = [
    ("C35:G35", "C36:G36"),
    ("(G35*", "(G36*"),
    ("+G35*", "+G36*"),
    ("B50+B51", "B51+B52"),
    ("/B55", "/B56"),
]


def _rewrite_formula(formula: str, rules: list[tuple[str, str]]) -> str:
    out = formula
    for old, new in rules:
        out = out.replace(old, new)
    return out


def fix(path: Path = TARGET) -> dict:
    tmp = path.with_suffix(".reffix.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"global": 0, "sensitivity": 0, "specific": 0}

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not isinstance(val, str) or not val.startswith("="):
                    continue
                new_val = _rewrite_formula(val, GLOBAL_REWRITES)
                if ws.title == "DCF" and cell.row in range(100, 105):
                    newer = _rewrite_formula(new_val, SENS_REWRITES)
                    if newer != new_val:
                        stats["sensitivity"] += 1
                    new_val = newer
                if new_val != val:
                    cell.value = new_val
                    stats["global"] += 1

    dcf = wb["DCF"]
    comps = wb["Comps"]

    # Capex % row should link assumption, not compute capex dollars
    for col in ("C", "D", "E", "F", "G"):
        if str(dcf[f"{col}18"].value or "").startswith("=-"):
            dcf[f"{col}18"] = "=Scenarios!$D$13"
            stats["specific"] += 1

    # Repurchase schedule UFCF link
    if dcf["G69"].value == "=G35":
        dcf["G69"] = "=G36"
        stats["specific"] += 1

    # Exit multiple method — selected multiple lives on B82 (matches unaltered E82=E46)
    dcf["B82"] = "=B47"
    stats["specific"] += 1

    # Comps cross-checks (unaltered E44=terminal EBITDA, E46=Gordon implied multiple)
    if comps["B6"].value == "=DCF!B44":
        comps["B6"] = "=DCF!B45"
        stats["specific"] += 1
    if comps["B44"].value == "=DCF!B46":
        comps["B44"] = "=DCF!B47"
        stats["specific"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"Stale formula ref fix: {fix()}")
