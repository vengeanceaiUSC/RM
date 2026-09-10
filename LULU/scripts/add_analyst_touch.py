#!/usr/bin/env python3
"""Populate Scratch tab with messy analyst working notes — not a sterile formula block.

Real scratch tabs hold pasted 10-K commentary, misaligned EPS blocks, and rough WACC
sanity checks. No model cross-links; no perfect alignment.

Run:  cd LULU && python3 scripts/add_analyst_touch.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

SCRIPTS = Path(__file__).resolve().parent
import sys

sys.path.insert(0, str(SCRIPTS))
from styles import FONT_NAME  # noqa: E402

ROOT = SCRIPTS.parent
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# Raw 10-K management commentary excerpt (FY2025, Item 7 — trimmed)
TENK_DUMP = (
    "MD&A — FY2025: Net revenue increased 7% to $11.1 billion. Comparable sales "
    "increased 1% on a constant dollar basis. Americas comparable sales decreased 3%; "
    "China Mainland increased 20%; Rest of World increased 9%. Gross profit as a percent "
    "of net revenue was 59.2% compared to 58.3% last year. We ended the year with 811 "
    "company-operated stores globally. Looking ahead, we expect net revenue to decrease "
    "5% to 7% in Q2 FY2026 vs prior year..."
)

# Misaligned EPS block (pasted from press release — columns don't line up)
EPS_BLOCK = [
    ("Q2 FY25", "2.92", "$"),
    ("Q3 FY25", "2.87", ""),
    ("Q4 FY25", "5.99", "  "),
    ("FY25", "14.64", " diluted"),
    ("FY26 guide", "~$14.50", "??? check slide 8"),
]

# Rough WACC sanity checks (disconnected, unformatted)
WACC_CHECKS = [
    ("rf", "4.8%", "FRED DGS10 9/9/26"),
    ("ERP", "5.5%", "Damodaran Jan-26 + 177bps"),
    ("beta", "1.35", "Yahoo 5Y mo — relever w/ lease D/E"),
    ("WACC?", "8.1%", "back of envelope — tie to WACC tab"),
]


def _write_messy_scratch(wb) -> None:
    if "Scratch" in wb.sheetnames:
        del wb["Scratch"]
    ws = wb.create_sheet("Scratch")

    ws["A1"] = "scratch — dont touch"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=11)

    ws["A3"] = "10-K dump (FY25 MD&A)"
    ws["A3"].font = Font(name=FONT_NAME, bold=True, size=9)
    ws["A4"] = TENK_DUMP
    ws["A4"].font = Font(name=FONT_NAME, size=9)
    ws["A4"].alignment = openpyxl.styles.Alignment(wrap_text=True)
    ws.row_dimensions[4].height = 72

    ws["D8"] = "EPS paste from release"
    ws["D8"].font = Font(name=FONT_NAME, bold=True, size=9)
    for i, (q, eps, note) in enumerate(EPS_BLOCK):
        r = 9 + i
        ws.cell(r, 4, q)
        ws.cell(r, 6, eps)  # skip col E — misaligned on purpose
        ws.cell(r, 8, note)
        for c in (4, 6, 8):
            ws.cell(r, c).font = Font(name=FONT_NAME, size=9)

    ws["A18"] = "wacc sanity"
    ws["A18"].font = Font(name=FONT_NAME, bold=True, size=9)
    for i, (label, val, note) in enumerate(WACC_CHECKS):
        r = 19 + i
        ws.cell(r, 1, label)
        ws.cell(r, 3, val)
        ws.cell(r, 5, note)
        ws.row_dimensions[r].height = 14 if i % 2 == 0 else 19
        for c in (1, 3, 5):
            ws.cell(r, c).font = Font(name=FONT_NAME, size=9)

    ws["G22"] = "todo: fix capex fade note on Scenarios (typo: 'fades donw')"
    ws["G22"].font = Font(name=FONT_NAME, size=8, italic=True)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 2
    ws.column_dimensions["F"].width = 8
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 10
    ws.sheet_view.showGridLines = True


def add_touch(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".touch.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)

    # Remove prior yellow highlights on Scenarios inputs (user applies colors manually)
    for coord in ("F4", "F10", "F13"):
        cell = wb["Scenarios"][coord]
        cell.fill = openpyxl.styles.PatternFill()

    _write_messy_scratch(wb)

    wb.save(tmp)
    tmp.replace(path)
    return {"scratch_rebuilt": 1, "highlights_cleared": 3}


if __name__ == "__main__":
    result = add_touch()
    print(f"Analyst touch applied: {result}")
