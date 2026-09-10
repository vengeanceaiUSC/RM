#!/usr/bin/env python3
"""Add subtle analyst-built cues without changing model numbers or formulas.

- Scratch tab with quick sanity checks (formulas only, no new hardcodes)
- Yellow highlight on 2–3 base-case inputs under review
- Orphan rough calc off to the right on Scenarios (common working area)

Run:  cd LULU && python3 scripts/add_analyst_touch.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

YELLOW = PatternFill(fill_type="solid", fgColor="FFFF99")
FONT_NAME = "Garamond"

# Base-case Scenarios inputs analysts often highlight while debating
HIGHLIGHT_CELLS = [
    ("Scenarios", "F4"),   # FY2026E revenue growth
    ("Scenarios", "F10"),  # Terminal growth
    ("Scenarios", "F13"),  # Capex %
]


def add_touch(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".touch.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"scratch": 0, "highlights": 0, "side_calc": 0}

    # Scratch tab
    if "Scratch" not in wb.sheetnames:
        ws = wb.create_sheet("Scratch")
        ws["A1"] = "Quick checks (not in model)"
        ws["A1"].font = Font(name=FONT_NAME, bold=True, size=10)
        ws["A3"] = "Mkt cap ($k)"
        ws["B3"] = "=WACC!D10"
        ws["A4"] = "FY26 rev growth check"
        ws["B4"] = "=Scenarios!F4"
        ws["A5"] = "Terminal g"
        ws["B5"] = "=Scenarios!F10"
        ws["A6"] = "Implied px (DCF)"
        ws["B6"] = "=DCF!D56"
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 14
        for r in range(3, 7):
            ws.cell(r, 1).font = Font(name=FONT_NAME, size=10)
            ws.cell(r, 2).font = Font(name=FONT_NAME, size=10, color="006100")
        stats["scratch"] = 1

    # Yellow highlights on key inputs
    for sheet, coord in HIGHLIGHT_CELLS:
        cell = wb[sheet][coord]
        cell.fill = YELLOW
        stats["highlights"] += 1

    # Side working column on Scenarios (does not feed the model)
    scn = wb["Scenarios"]
    if scn["I4"].value is None:
        scn["I3"] = "spot check"
        scn["I3"].font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
        scn["I4"] = "=F25*F4"
        scn["I4"].font = Font(name=FONT_NAME, size=9, color="808080")
        scn.column_dimensions["I"].width = 11
        stats["side_calc"] = 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    result = add_touch()
    print(f"Analyst touch applied: {result}")
