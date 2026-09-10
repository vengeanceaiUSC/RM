#!/usr/bin/env python3
"""Set Garamond on every cell in the LULU Excel model (GIS font rule).

Preserves color, bold, italic, size, and underline on each run.

Run:  cd LULU && python3 scripts/apply_garamond_excel.py
"""
from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "model18_wsp_formulas.xlsx"
FONT_NAME = "Garamond"


def _garamond(cell: Cell) -> int:
    if cell.font is None:
        return 0
    if cell.font.name == FONT_NAME:
        return 0
    f = copy(cell.font)
    f.name = FONT_NAME
    cell.font = f
    return 1


def apply(path: Path = MODEL) -> dict[str, int]:
    wb = load_workbook(path)
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                changed += _garamond(cell)
    wb.save(path)
    return {"cells_updated": changed, "file": path.name}


def main() -> None:
    stats = apply()
    print(f"Garamond applied → {stats['file']} ({stats['cells_updated']} cells updated)")


if __name__ == "__main__":
    main()
