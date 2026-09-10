#!/usr/bin/env python3
"""Apply human display formats to long-float hardcodes without changing stored values.

Grid shows 6.3 / 128.8 / 5.1% while formula bar keeps full precision (safe for DCF).

Run:  cd LULU && python3 scripts/fix_decimal_display.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# (sheet, row, cols, format)
DAY_ROWS = [
    ("Scenarios", 15, (5, 6, 7)),   # DSO
    ("Scenarios", 16, (5, 6, 7)),   # DIO
    ("Scenarios", 18, (5, 6, 7)),   # DPO
]
PCT_ROWS = [
    ("Scenarios", 19, (5, 6, 7)),   # Prepaid
    ("Scenarios", 20, (5, 6, 7)),   # Accrued
    ("DCF", 7, (4,)),               # FY25 gross margin
]


def fix(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".decfmt.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"day_fmt": 0, "pct_fmt": 0}

    for sheet, row, cols in DAY_ROWS:
        ws = wb[sheet]
        for c in cols:
            cell = ws.cell(row, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0"
                stats["day_fmt"] += 1

    for sheet, row, cols in PCT_ROWS:
        ws = wb[sheet]
        for c in cols:
            cell = ws.cell(row, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"
                stats["pct_fmt"] += 1

    # Broader pass: any hardcode with >4 decimal places visible
    for sn in ("WACC", "Scenarios", "NOPAT Bridge", "DCF", "Comps", "Revenue Drivers"):
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, float):
                    continue
                if cell.value == int(cell.value) and abs(cell.value) >= 100:
                    continue
                s = f"{cell.value:.12f}".rstrip("0")
                if "." not in s or len(s.split(".")[1]) <= 3:
                    continue
                label = str(ws.cell(cell.row, 1).value or "").lower()
                if any(k in label for k in ("dso", "dio", "dpo", "days")):
                    cell.number_format = "0.0"
                    stats["day_fmt"] += 1
                elif abs(cell.value) < 1 and any(
                    k in label for k in ("%", "margin", "rate", "growth", "weight", "mix")
                ):
                    cell.number_format = "0.0%"
                    stats["pct_fmt"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"Decimal display formats: {fix()}")
