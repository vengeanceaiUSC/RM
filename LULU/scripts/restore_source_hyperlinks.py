#!/usr/bin/env python3
"""Restore Source column hyperlinks from model18unaltered (12).xlsx.

Re-applies clickable URLs and internal sheet links on col C (main tabs) and col J
(Revenue Drivers), with blue italic underline styling. Does not change hardcoded
values or formulas.

Run:  cd LULU && python3 scripts/restore_source_hyperlinks.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import data as D  # noqa: E402
from restore_source_columns import (  # noqa: E402
    _apply_source_cell,
    _build_label_map,
    _clean_source_label,
    _collect_original_source,
    _match_row,
    _norm_label,
    _resolve_internal_link,
)

ROOT = SCRIPTS.parent
ORIGINAL = ROOT / "model18unaltered (12).xlsx"
TARGET = ROOT / "unbeiesgbar_final.xlsx"

BLUE = "0000CC"
FONT_NAME = "Garamond"

SHEET_SOURCE: dict[str, tuple[int, int]] = {
    "WACC": (3, 3),
    "Scenarios": (3, 3),
    "NOPAT Bridge": (3, 3),
    "Comps": (3, 3),
    "DCF": (3, 3),
    "Revenue Drivers": (11, 10),
}

# Revenue Drivers rows with FY25 hardcodes → add 10-K link when unaltered had text-only
RD_FY25_LINK_ROWS: dict[int, str] = {
    6: D.filing_url("FY2025"),
    9: D.filing_url("FY2025"),
    10: D.filing_url("FY2025"),
    13: D.filing_url("FY2025"),
    14: D.filing_url("FY2025"),
    17: D.filing_url("FY2025"),
    18: D.filing_url("FY2025"),
    28: D.filing_url("FY2025"),
    29: D.filing_url("FY2025"),
    31: D.filing_url("FY2025"),
    41: D.filing_url("FY2025"),
    55: D.filing_url("FY2025"),
}


def _link_font() -> Font:
    return Font(name=FONT_NAME, color=BLUE, italic=True, underline="single", size=9)


def _style_source_cell(cell) -> None:
    if cell.hyperlink:
        cell.font = _link_font()


def restore(path: Path = TARGET, original: Path = ORIGINAL) -> dict[str, int]:
    if not path.exists() or not original.exists():
        raise FileNotFoundError(f"Missing {path} or {original}")

    tmp = path.with_suffix(".restore_links.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    orig_wb = openpyxl.load_workbook(original, data_only=False)
    sheet_shifts: dict[str, dict[str, str]] = {}

    stats = {"restored": 0, "styled": 0, "rd_fy25_links": 0}

    for sheet, (orig_col, tgt_col) in SHEET_SOURCE.items():
        if sheet not in wb.sheetnames or sheet not in orig_wb.sheetnames:
            continue
        ws = wb[sheet]
        orig_ws = orig_wb[sheet]
        orig_map = _build_label_map(orig_ws)

        for r in range(1, ws.max_row + 1):
            label_raw = ws.cell(r, 1).value
            if label_raw is None:
                continue
            label = str(label_raw).strip()
            if not label:
                continue

            orig_row = _match_row(label, orig_map)
            if orig_row is None:
                continue

            orig_label, orig_url, orig_loc = _collect_original_source(
                orig_ws, orig_row, source_col=orig_col
            )
            if not orig_label and not orig_url and not orig_loc:
                continue

            cell = ws.cell(r, tgt_col)
            if orig_loc:
                orig_loc = _resolve_internal_link(orig_wb, wb, orig_loc, sheet_shifts)

            if _apply_source_cell(cell, orig_label, orig_url, orig_loc):
                stats["restored"] += 1
            elif orig_label:
                cell.value = orig_label
                stats["restored"] += 1

            if cell.hyperlink:
                _style_source_cell(cell)
                stats["styled"] += 1

    # Revenue Drivers: add 10-K links on FY25 hardcode rows we cited without URLs
    rd = wb["Revenue Drivers"]
    for row, url in RD_FY25_LINK_ROWS.items():
        cell = rd.cell(row, 10)
        if cell.hyperlink:
            _style_source_cell(cell)
            stats["styled"] += 1
            continue
        if not cell.value:
            continue
        cell.hyperlink = url
        cell.font = _link_font()
        stats["rd_fy25_links"] += 1
        stats["styled"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


def verify(path: Path = TARGET, original: Path = ORIGINAL) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    orig_wb = openpyxl.load_workbook(original, data_only=False)
    issues: list[str] = []

    for sheet, (orig_col, tgt_col) in SHEET_SOURCE.items():
        if sheet not in wb.sheetnames:
            continue
        ws, orig_ws = wb[sheet], orig_wb[sheet]
        orig_map = _build_label_map(orig_ws)
        orig_links = 0
        final_links = 0
        for r in range(1, orig_ws.max_row + 1):
            if orig_ws.cell(r, orig_col).hyperlink:
                orig_links += 1
                label = orig_ws.cell(r, 1).value
            else:
                continue
            if not label:
                continue
            fr = _match_row(str(label).strip(), _build_label_map(ws))
            if fr and ws.cell(fr, tgt_col).hyperlink:
                final_links += 1
        if final_links < orig_links - 2:  # allow derived rows without orig links
            issues.append(f"{sheet}: {final_links}/{orig_links} source hyperlinks")

    w = wb["WACC"]
    if not w["C3"].hyperlink:
        issues.append("WACC C3 FRED link missing")
    if str(w["C3"].font.color.rgb or "").upper() not in ("0000CC", "000000CC", "FF0000CC"):
        issues.append(f"WACC C3 not blue link color: {w['C3'].font.color.rgb}")

    if issues:
        raise AssertionError("Verify failed:\n" + "\n".join(issues))
    print("Verify OK: source hyperlinks and blue link styling restored")


if __name__ == "__main__":
    result = restore()
    print(f"Restored source hyperlinks: {result}")
    verify()
