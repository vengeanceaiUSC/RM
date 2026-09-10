#!/usr/bin/env python3
"""Wire FY2025 DCF bridge rows and fix mis-formatted tax dollar rows.

FY25 column holds reported 10-K items; the UFCF build (taxes → NOPAT → D&A →
capex) intentionally starts at FY26. This pass fills FY25 bridge links so the
column is not blank, and corrects tax *dollar* rows formatted as percentages.

Run:  cd LULU && python3 scripts/fix_dcf_fy25_bridge.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

FY25_CAPEX = 680802  # 10-K cash flow statement ($000)
NUM = "#,##0;(#,##0)"
PCT = "0.0%"

# Row labels whose values are $000, not rates — override bad % formats
DOLLAR_TAX_ROWS = (
    "less: taxes",
    "unlevered tax on",
    "normalized ebit",
)


def _is_dollar_tax_label(label: str) -> bool:
    lab = label.lower()
    return any(k in lab for k in DOLLAR_TAX_ROWS)


def fix(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".fy25.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    nb = wb["NOPAT Bridge"]
    dcf = wb["DCF"]
    stats = {"nopat": 0, "dcf": 0, "formats": 0}

    # NOPAT FY25 tax + NOPAT
    if nb["B40"].value != "=B36*B39":
        nb["B40"] = "=B36*B39"
        stats["nopat"] += 1
    if nb["B41"].value != "=B36-B40":
        nb["B41"] = "=B36-B40"
        stats["nopat"] += 1

    # DCF FY25 bridge links
    links: list[tuple[str, str]] = [
        ("B13", "=-'NOPAT Bridge'!B40"),
        ("B14", "='NOPAT Bridge'!B41"),
        ("B17", "=B15"),
        ("B19", f"=-{FY25_CAPEX}"),
    ]
    for addr, formula in links:
        if dcf[addr].value != formula:
            dcf[addr] = formula
            stats["dcf"] += 1

    # Fix tax/EBIT dollar rows mis-tagged as %
    for ws_name in ("NOPAT Bridge", "DCF"):
        ws = wb[ws_name]
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "")
            if not _is_dollar_tax_label(label):
                continue
            cols = range(2, 8) if ws_name == "DCF" else range(2, 8)
            for c in cols:
                cell = ws.cell(r, c)
                if cell.number_format != NUM:
                    cell.number_format = NUM
                    stats["formats"] += 1

    # Tax rate rows stay as %
    for ws_name in ("NOPAT Bridge", "DCF"):
        ws = wb[ws_name]
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "").lower()
            if "tax rate" in label or "effective tax rate" in label:
                for c in range(2, 8):
                    ws.cell(r, c).number_format = PCT

    # Normalized EBIT row (contains "tax base" — not a rate)
    for c in range(2, 8):
        if nb.cell(36, c).number_format != NUM:
            nb.cell(36, c).number_format = NUM
            stats["formats"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"FY25 DCF bridge fix: {fix()}")
