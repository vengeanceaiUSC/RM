#!/usr/bin/env python3
"""Label Revenue Drivers unit scalars with explicit unit meaning.

B3 = ×1,000,000 (digital sessions input is in millions)
C3 = ÷1,000 (formula output in $000 on this tab)

Run:  cd LULU && python3 scripts/label_rd_unit_conversions.py
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
TARGETS = [
    ROOT / "model18_wsp_formulas.xlsx",
    ROOT / "model18_humanized.xlsx",
    ROOT / "model18_humanized (1).xlsx",
]

B3_LABEL = "×1,000,000 (sessions in millions → count)"
C3_LABEL = "÷1,000 ($ → $000 on this tab)"


def fix_workbook(path: Path) -> None:
    wb = openpyxl.load_workbook(path)
    rd = wb["Revenue Drivers"]
    rd["A3"] = "Unit conversions (drivers)"
    rd["B2"] = B3_LABEL
    rd["C2"] = C3_LABEL
    rd["B3"] = 1_000_000
    rd["C3"] = 1_000
    for addr in ("B2", "C2"):
        rd[addr].font = Font(italic=True, size=9)
    rd["H3"] = "see B2/C2"
    rd["H3"].font = Font(italic=True, size=8)
    wb.save(path)


if __name__ == "__main__":
    for path in TARGETS:
        if path.exists():
            fix_workbook(path)
            print(f"Updated {path.name}")
