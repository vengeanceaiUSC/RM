#!/usr/bin/env python3
"""Add 8-word driver notes to Revenue Drivers col I (How it moves).

Run:  cd LULU && python3 scripts/add_revenue_driver_notes.py
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "model18_wsp_formulas.xlsx"

# row -> note (~8 words; colon style)
NOTES: dict[int, str] = {
    # I. Store fleet
    6: "Prior-year ending count carried into this year.",
    7: "New stores added in region each forecast year.",
    8: "Stores closed yearly as leases roll off.",
    9: "Beginning plus openings minus closures each year.",
    10: "Prior-year ending count carried into this year.",
    11: "New stores added in region each forecast year.",
    12: "Stores closed yearly as leases roll off.",
    13: "Beginning plus openings minus closures each year.",
    14: "Prior-year ending count carried into this year.",
    15: "New stores added in region each forecast year.",
    16: "Stores closed yearly as leases roll off.",
    17: "Beginning plus openings minus closures each year.",
    18: "Sum of Americas, China, and Rest of World.",
    19: "Average box size grows modestly through FY30.",
    20: "Ending stores times average square feet per store.",
    # II. Store productivity
    24: "Sales per foot recovers slowly after FY26 trough.",
    25: "Declines then flatlines as Americas traffic stabilizes.",
    26: "Strong growth moderates as China store base scales.",
    27: "Mid-single-digit comps fading to low-single digits.",
    28: "Last year's store revenue rolled forward annually.",
    29: "Prior revenue grown by weighted geographic comp rates.",
    30: "Net new doors times sq ft, productivity, ramp.",
    31: "Comparable store sales plus new store contribution.",
    # III. E-commerce
    38: "Digital traffic grows modestly across forecast years.",
    39: "Conversion rate recovers gradually from FY26 pressure.",
    40: "Order size steps up slowly with promo normalization.",
    41: "Sessions times conversion times average order value.",
    # IV. Other / geo / mix
    45: "Wholesale, license, outlets: low-single-digit growth.",
    46: "Americas revenue scales with consolidated growth rate.",
    47: "China revenue grows faster than consolidated average.",
    48: "Rest of World scales with total revenue growth.",
    49: "Women's share fades slightly as men's mix rises.",
    50: "Men's mix rises gradually through the forecast.",
    51: "Accessories share held steady near thirteen percent.",
    # V. Reconciliation
    55: "Store channel plus e-comm plus other channels.",
    56: "Pulls Scenarios base revenue path used by DCF.",
    57: "Driver-built revenue minus Scenarios base-case revenue.",
    58: "Variance as percent of Scenarios consolidated revenue.",
}


def main() -> None:
    wb = openpyxl.load_workbook(TARGET)
    rd = wb["Revenue Drivers"]
    rd["I5"] = "How it moves"
    rd["I5"].font = Font(bold=True, size=9)
    rd["I5"].alignment = Alignment(horizontal="left", wrap_text=True)

    note_font = Font(size=9, italic=True, color="404040")
    for row, text in NOTES.items():
        cell = rd[f"I{row}"]
        cell.value = text
        cell.font = note_font
        cell.alignment = Alignment(horizontal="left", wrap_text=True, vertical="top")

    rd.column_dimensions["I"].width = 36
    wb.save(TARGET)
    print(f"Added {len(NOTES)} Revenue Drivers notes → {TARGET}")


if __name__ == "__main__":
    main()
