#!/usr/bin/env python3
"""Add 8-word driver notes to Scenarios col B (colon format).

Run:  cd LULU && python3 scripts/add_scenario_driver_notes.py
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "model18_wsp_formulas.xlsx"

# row -> note (8-word behavior description; colon style)
NOTES: dict[int, str] = {
    # Key assumptions (rows 4–22)
    4: "Year-one sales change: bear falls most.",
    5: "Years two-five average growth: bull fastest.",
    6: "Starting operating margin before linear ramp.",
    7: "One-time year-one EBIT boost: zero after.",
    8: "Year-five target margin after linear ramp.",
    9: "Flat discount rate: higher bear, lower bull.",
    10: "Perpetuity growth rate: flat across all years.",
    11: "Flat tax percentage applied every forecast year.",
    12: "Depreciation scales as fixed percent of revenue.",
    13: "Investment scales as fixed percent of revenue.",
    14: "Merchandise margin held flat every forecast year.",
    15: "Collection days held flat near six yearly.",
    16: "Starting inventory days before annual decline begins.",
    17: "Inventory days fall one day each year.",
    18: "Payment days held flat near twenty-five yearly.",
    19: "Other current assets flat percent of revenue.",
    20: "Accrued payables flat percent of revenue.",
    21: "Bear base retire fixed dollars yearly.",
    22: "Bull spends set percent of free cash.",
    # 5-year forecast paths (year 1 rows)
    25: "Grows or shrinks using scenario-specific growth rates.",
    30: "Stays flat at scenario margin every year.",
    35: "Moves with revenue when gross margin stays flat.",
    40: "Linearly ramps from run-rate toward terminal margin.",
    45: "Revenue times margin: year one adds tariff refund.",
    50: "EBIT taxed at flat cash rate each year.",
    55: "Scales upward with revenue at fixed percentage annually.",
    60: "Scales upward with revenue at fixed percentage annually.",
    65: "Held flat at six days every forecast year.",
    70: "Tracks revenue changes while collection days stay flat.",
    75: "Falls one day per year across five years.",
    80: "Declines with DIO days: grows with rising COGS.",
    85: "Held flat at twenty-five days every year.",
    90: "Grows with COGS while payment days remain flat.",
    95: "Grows at flat percent of revenue yearly.",
    100: "Grows at flat percent of revenue every year.",
    105: "Rises as receivables, inventory, prepaids expand.",
    110: "Rises as payables and accruals grow with sales.",
    115: "Operating assets minus liabilities: level shifts yearly.",
    120: "Prior minus current NWC: positive releases cash.",
    125: "NOPAT plus depreciation minus capex plus working capital.",
}

# Replace em-dash / en-dash separators in col A labels with colons
LABEL_FIXES = (
    (" — flat vs FY25", ": flat vs FY25"),
    (" — flat w/ FY25", ": flat w/ FY25"),
    ("ΔNWC (prior yr − current yr) – year ", "ΔNWC (prior yr minus current yr): year "),
    (" – year ", ": year "),
)


def main() -> None:
    wb = openpyxl.load_workbook(TARGET)
    scn = wb["Scenarios"]
    scn["B3"] = "How it moves"
    scn["B3"].font = Font(bold=True, size=9)
    scn["B3"].alignment = Alignment(horizontal="left", wrap_text=True)

    note_font = Font(size=9, italic=True, color="404040")
    for row, text in NOTES.items():
        cell = scn[f"B{row}"]
        cell.value = text
        cell.font = note_font
        cell.alignment = Alignment(horizontal="left", wrap_text=True, vertical="top")

    for row in range(1, scn.max_row + 1):
        cell = scn.cell(row, 1)
        if not cell.value or not isinstance(cell.value, str):
            continue
        label = cell.value
        for old, new in LABEL_FIXES:
            label = label.replace(old, new)
        cell.value = label

    wb.save(TARGET)
    print(f"Added {len(NOTES)} driver notes → {TARGET}")


if __name__ == "__main__":
    main()
