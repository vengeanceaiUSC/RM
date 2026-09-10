#!/usr/bin/env python3
"""Clear robotic Source labels on pure calculation rows (no external data).

Leaves Source hyperlinks and hardcode provenance intact. Only clears text-only
labels like 'Revenue Drivers: ending-stores formula' on formula rows.

Run:  cd LULU && python3 scripts/clear_formula_sources.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# Text-only source labels to clear when the value cell is a formula.
CLEAR_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"^Revenue Drivers:\s*.+\s*formula$", re.I),
    re.compile(r"^Revenue Drivers:\s*.+\s*calc$", re.I),
    re.compile(r"^NOPAT Bridge tab:", re.I),
    re.compile(r"^WACC tab:\s*price × shares$", re.I),
    re.compile(r"^WACC tab:\s*CAPM build$", re.I),
]

SOURCE_COL: dict[str, int] = {
    "WACC": 3,
    "Scenarios": 3,
    "NOPAT Bridge": 3,
    "Comps": 3,
    "DCF": 3,
    "Revenue Drivers": 10,
}

VALUE_COLS: dict[str, tuple[int, ...]] = {
    "WACC": (4,),
    "Scenarios": (5, 6, 7),
    "NOPAT Bridge": (4, 5, 6, 7, 8, 9),
    "DCF": tuple(range(4, 11)),
    "Comps": tuple(range(4, 11)),
    "Revenue Drivers": tuple(range(2, 9)),
}


def _is_formula(val) -> bool:
    return isinstance(val, str) and val.startswith("=")


def _row_is_formula_only(ws, row: int, vcols: tuple[int, ...]) -> bool:
    has_formula = False
    has_hardcode = False
    for c in vcols:
        v = ws.cell(row, c).value
        if v is None:
            continue
        if _is_formula(v):
            has_formula = True
        elif isinstance(v, (int, float)):
            has_hardcode = True
        elif isinstance(v, str) and not v.startswith("="):
            has_hardcode = True
    return has_formula and not has_hardcode


def clear(path: Path = TARGET) -> list[str]:
    tmp = path.with_suffix(".src_clear.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    changes: list[str] = []

    for sheet, scol in SOURCE_COL.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        vcols = VALUE_COLS.get(sheet, (4,))
        for r in range(1, ws.max_row + 1):
            src = ws.cell(r, scol)
            if src.hyperlink:
                continue
            if not isinstance(src.value, str) or not src.value.strip():
                continue
            text = src.value.strip()
            should_clear = False
            if sheet == "Revenue Drivers" and text.startswith("Revenue Drivers:"):
                # Internal roll-forward / arithmetic — leave blank like a human model
                should_clear = True
            elif any(p.search(text) for p in CLEAR_PATTERNS):
                should_clear = _row_is_formula_only(ws, r, vcols)
            if not should_clear:
                continue
            changes.append(f"{sheet}!{src.coordinate}: {text!r} → (blank)")
            src.value = None

    wb.save(tmp)
    tmp.replace(path)
    return changes


if __name__ == "__main__":
    changed = clear()
    print(f"Cleared {len(changed)} robotic formula sources")
    for c in changed:
        print(f"  {c}")
