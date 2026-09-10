#!/usr/bin/env python3
"""Repair DCF WACC × terminal-g sensitivity grid.

Bug: cols F–H used ^$E$36 (empty → TV undiscounted, inflated prices);
cols I–J kept stale hardcoded ^5 formulas → cliff at 2.5% / 3.0% g.

Fix: all F99:J103 use $A{row}, {col}$98, and ^$J$36 (forecast horizon = 5).

Run:  cd LULU && python3 scripts/fix_sensitivity_grid.py
"""
from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "model18_wsp_formulas.xlsx"
GCOLS = ("F", "G", "H", "I", "J")


def fix_sensitivity_grid(path: Path = TARGET) -> int:
    wb = openpyxl.load_workbook(path)
    dcf = wb["DCF"]
    n = 0
    for row in range(99, 104):
        for col in GCOLS:
            new = (
                f"=(NPV($A{row},$F$35:$J$35)+($J$35*(1+{col}$98)/($A{row}-{col}$98))"
                f"/(1+$A{row})^$J$36+$E$50+$E$51)/$E$55"
            )
            cell = dcf[f"{col}{row}"]
            if cell.value != new:
                cell.value = new
                n += 1
    wb.save(path)
    return n


if __name__ == "__main__":
    changes = fix_sensitivity_grid()
    print(f"Fixed {changes} sensitivity cells → {TARGET}")
