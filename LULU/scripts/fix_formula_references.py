#!/usr/bin/env python3
"""Replace hardcoded FY25 anchors in formula text with cell links (same values).

- Scenarios revenue: 11102600 → DCF!$D$5 (Net revenue FY25A)
- Comps EBITDA: 2210615+496228 → DCF!D11+496228 (EBIT link; D&A constant unchanged)

Does not change stored hardcodes or valuation inputs.

Run:  cd LULU && python3 scripts/fix_formula_references.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

REPLACEMENTS: list[tuple[str, str]] = [
    ("=11102600*(1+E4)", "=DCF!$D$5*(1+E4)"),
    ("=11102600*(1+F4)", "=DCF!$D$5*(1+F4)"),
    ("=11102600*(1+G4)", "=DCF!$D$5*(1+G4)"),
    ("=2210615+496228", "=DCF!D11+496228"),
]


def fix(path: Path = TARGET) -> list[str]:
    tmp = path.with_suffix(".formula_refs.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    changes: list[str] = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                for old, new in REPLACEMENTS:
                    if v == old:
                        cell.value = new
                        changes.append(f"{ws.title}!{cell.coordinate}: {old} → {new}")
                        break

    wb.save(tmp)
    tmp.replace(path)
    return changes


if __name__ == "__main__":
    changed = fix()
    print(f"Updated {len(changed)} formula references")
    for c in changed:
        print(f"  {c}")
