#!/usr/bin/env python3
"""Fill missing Source (col J) text on Revenue Drivers hardcode rows.

Every hardcoded assumption must cite its origin. Formula-only rows are left blank.

Run:  cd LULU && python3 scripts/add_revenue_driver_sources.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

SCRIPTS = Path(__file__).resolve().parent
import sys

sys.path.insert(0, str(SCRIPTS))
from styles import BLACK, FONT_NAME  # noqa: E402

ROOT = SCRIPTS.parent
TARGET = ROOT / "unbeiesgbar_final.xlsx"

SOURCE_COL = 10
VALUE_COLS = range(2, 9)

# Row → specific citation (matches data.py ASSUMPTION_SRC keys where applicable)
ROW_SOURCES: dict[int, str] = {
    6: "LULU FY2025 10-K, Item 2 — Americas 462 company-operated stores",
    9: "LULU FY2025 10-K, store roll-forward — Americas ending 476 stores",
    10: "LULU FY2025 10-K, Item 2 — China Mainland 151 stores",
    13: "LULU FY2025 10-K, store roll-forward — China ending 172 stores",
    14: "LULU FY2025 10-K, Item 2 — Rest of World 154 stores",
    17: "LULU FY2025 10-K, store roll-forward — RoW ending 163 stores",
    18: "LULU FY2025 10-K, Item 2 — Total company-operated 811 stores",
    28: "LULU FY2025 10-K, company-operated stores net revenue ($5,049.7M)",
    29: "LULU FY2025 10-K, comparable store sales base (prior-year store rev)",
    30: "Revenue Drivers tab — net-new-store calc; FY25 anchor $0",
    31: "LULU FY2025 10-K, company-operated stores net revenue ($5,049.7M)",
    41: "LULU FY2025 10-K, e-commerce net revenue ($4,918.7M)",
    55: "LULU FY2025 10-K, consolidated net revenue ($11,102.6M)",
}


def _is_hardcode(val) -> bool:
    return isinstance(val, (int, float)) and not isinstance(val, bool)


def add_sources(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".rd_src.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    ws = wb["Revenue Drivers"]
    stats = {"filled": 0, "skipped_existing": 0}

    for row, citation in ROW_SOURCES.items():
        src_cell = ws.cell(row, SOURCE_COL)
        if src_cell.value or src_cell.hyperlink:
            stats["skipped_existing"] += 1
            continue
        has_hc = any(_is_hardcode(ws.cell(row, c).value) for c in VALUE_COLS)
        if not has_hc:
            continue
        src_cell.value = citation
        src_cell.font = Font(name=FONT_NAME, color=BLACK, italic=True, size=9)
        stats["filled"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    result = add_sources()
    print(f"Revenue Drivers sources: {result}")
