#!/usr/bin/env python3
"""Link DCF repurchase schedule NI / shares / EPS to Scenarios pitch bridge.

The deck cites DCF row 78 (Share Repurchase & EPS Accretion Schedule) on
slides 12–14. Forecast cols F–J must match Scenarios base-case pitch bridge
(rows 143–147 NI, 183–187 shares, 188–192 EPS).

Run:  cd LULU && python3 scripts/fix_dcf_repurchase_eps.py
"""
from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "model18_wsp_formulas.xlsx"

FCOLS = ["F", "G", "H", "I", "J"]  # FY26E–FY30E
BASE = "C"  # Scenarios base-case column


def main() -> None:
    wb = openpyxl.load_workbook(TARGET)
    dcf = wb["DCF"]
    n = 0
    for i, col in enumerate(FCOLS):
        ni_row = 143 + i
        sh_row = 183 + i
        eps_row = 188 + i
        links = (
            (f"{col}77", f"=Scenarios!${BASE}${ni_row}"),
            (f"{col}75", f"=Scenarios!${BASE}${sh_row}"),
            (f"{col}78", f"=Scenarios!${BASE}${eps_row}"),
        )
        for addr, formula in links:
            if dcf[addr].value != formula:
                dcf[addr] = formula
                n += 1
    wb.save(TARGET)
    print(f"Linked {n} DCF repurchase cells to Scenarios pitch bridge → {TARGET}")


if __name__ == "__main__":
    main()
