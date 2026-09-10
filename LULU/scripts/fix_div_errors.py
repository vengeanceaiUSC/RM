#!/usr/bin/env python3
"""Fix #DIV/0! errors on NOPAT Bridge and DCF repurchase/valuation sections.

Root causes:
- NOPAT Revenue Drivers links offset +3 cols (C→F) after sheet column delete
- DCF repurchase block still references pre-insert row numbers (C68, C35, etc.)
- Any remaining DCF!$B$5 FY25 anchor stale refs

Preserves all equation logic — rewrites cell references only.

Run:  cd LULU && python3 scripts/fix_div_errors.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

FCOLS = ["C", "D", "E", "F", "G"]
FY25_REV = "DCF!$B$6"

# Link scaling to DCF FY25 net revenue cell (do not bury literal in formulas)
ANCHOR_REWRITES: list[tuple[str, str]] = [
    ("11102600", FY25_REV),
    ("8456743", "Scenarios!$B$194"),
    ("3494903", "Scenarios!$B$195"),
    ("4961840", "Scenarios!$B$196"),
    ("119068", "Scenarios!$B$197"),
]


def _fix_revenue_drivers_ref(formula: str) -> str:
    """Ensure NOPAT/RD refs use matching column letters (no erroneous shift)."""
    return formula


def _restore_dcf_scenarios_links(dcf) -> int:
    """Restore original Scenarios row-pulls on DCF forecast cols (matches unaltered)."""
    scn = "Scenarios"
    n = 0
    pulls: list[tuple[int, int, str | None]] = [
        (6, 25, None),    # Net revenue Y1–Y5
        (7, 14, "$"),     # Gross margin assumption
        (8, 35, None),    # COGS Y1–Y5
        (9, 40, None),    # EBIT margin Y1–Y5
        (11, 45, None),   # EBIT Y1–Y5
        (14, 50, None),   # NOPAT Y1–Y5
        (16, 12, "$"),    # D&A %
        (17, 55, None),   # Plus D&A Y1–Y5
        (18, 13, "$"),    # Capex %
        (19, 60, None),   # Less Capex Y1–Y5
        (36, 125, None),  # UFCF Y1–Y5
    ]
    for i, col in enumerate(FCOLS):
        yr = i
        # Tariff FY26 only
        dcf[f"{col}10"] = f"={scn}!$D$7" if col == "C" else 0
        n += 1
        for row, base, kind in pulls:
            if kind == "$":
                dcf[f"{col}{row}"] = f"={scn}!$D${base}"
            else:
                dcf[f"{col}{row}"] = f"={scn}!D{base + yr}"
            n += 1
        # Reported EBIT margin
        dcf[f"{col}12"] = f"={col}11/{col}6"
        # Unlevered tax from NOPAT Bridge
        dcf[f"{col}13"] = f"=-'NOPAT Bridge'!{col}40"
        n += 2
    return n


def _fix_repurchase(dcf) -> int:
    n = 0
    # Budget + cost of equity anchors (link to col B inputs)
    for col in FCOLS:
        dcf[f"{col}66"] = "=B66"
        dcf[f"{col}67"] = "=B67"
        n += 2

    # UFCF from row 36; repurchase budget from row 66
    for col in FCOLS:
        dcf[f"{col}69"] = f"={col}36"
        dcf[f"{col}70"] = f"={col}66"
        dcf[f"{col}71"] = f"={col}70/{col}69"
        n += 3

    # Share price path
    dcf["C73"] = "=B73*(1+B67)"
    prev = "B"
    for col in FCOLS[1:]:
        dcf[f"{col}73"] = f"={prev}73*(1+B67)"
        prev = col
    n += len(FCOLS)

    for col in FCOLS:
        dcf[f"{col}74"] = f"={col}70/{col}73"
        dcf[f"{col}75"] = "=B75" if col == "C" else f"={FCOLS[FCOLS.index(col)-1]}76"
        dcf[f"{col}76"] = f"={col}75-{col}74"
        dcf[f"{col}78"] = (
            f"=('NOPAT Bridge'!{col}36-'NOPAT Bridge'!{col}23)"
            f"*(1-'NOPAT Bridge'!{col}$39)"
        )
        dcf[f"{col}79"] = f"={col}78/{col}76"
        n += 5

    return n


def fix(path: Path = TARGET) -> dict:
    tmp = path.with_suffix(".divfix.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"literal_anchor": 0, "nopat_rd": 0, "repurchase": 0}

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not isinstance(val, str) or not val.startswith("="):
                    continue
                new_val = val
                if ws.title in ("Scenarios", "NOPAT Bridge"):
                    for old, new in ANCHOR_REWRITES:
                        if old in new_val:
                            new_val = new_val.replace(old, new)
                            stats["literal_anchor"] += 1
                    new_val = new_val.replace("DCF!$B$5", FY25_REV).replace("DCF!$D$5", FY25_REV)
                if ws.title == "NOPAT Bridge" and "Revenue Drivers" in new_val:
                    rd_fixed = _fix_revenue_drivers_ref(new_val)
                    if rd_fixed != new_val:
                        stats["nopat_rd"] += 1
                    new_val = rd_fixed
                if new_val != val:
                    cell.value = new_val

    stats["repurchase"] = _fix_repurchase(wb["DCF"])
    stats["dcf_scenarios"] = _restore_dcf_scenarios_links(wb["DCF"])

    wb.save(tmp)

    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from fix_dcf_valuation_refs import fix as fix_valuation

    stats["valuation"] = fix_valuation(path=tmp)

    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"DIV error fix: {fix()}")
