#!/usr/bin/env python3
"""Remove hover-comment red triangles and inflated row heights.

Visible Source links (col C / J) and all hardcoded values are preserved.
Removes duplicate Analyst hover comments where Source column already documents
provenance, and resets row heights inflated by legacy hint/comment sizing.

Run:  cd LULU && python3 scripts/remove_comment_artifacts.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

DOC_SHEETS = ("WACC", "Scenarios", "NOPAT Bridge", "Comps", "DCF")
SOURCE_COL = 3
RD_SOURCE_COL = 10
DEFAULT_ROW_HEIGHT = 15.0
MAX_NORMAL_HEIGHT = 32.0  # rows taller than this get reset


def _has_source(ws, row: int, source_col: int) -> bool:
    cell = ws.cell(row, source_col)
    if cell.hyperlink is not None:
        return True
    val = cell.value
    if not isinstance(val, str) or not val.strip():
        return False
    low = val.strip().lower()
    if low in {"source", "notes"}:
        return False
    return True


def remove_artifacts(path: Path = TARGET) -> dict[str, int]:
    if not path.exists():
        raise FileNotFoundError(path)

    tmp = path.with_suffix(".cleaning.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)

    stats = {"comments_removed": 0, "heights_reset": 0, "comments_kept": 0}

    for ws in wb.worksheets:
        source_col = None
        if ws.title in DOC_SHEETS:
            source_col = SOURCE_COL
        elif ws.title == "Revenue Drivers":
            source_col = RD_SOURCE_COL

        for row in ws.iter_rows():
            for cell in row:
                if not cell.comment:
                    continue
                keep = source_col is None or not _has_source(ws, cell.row, source_col)
                if keep:
                    stats["comments_kept"] += 1
                    continue
                cell.comment = None
                stats["comments_removed"] += 1

        for r, dim in list(ws.row_dimensions.items()):
            if dim.height and dim.height > MAX_NORMAL_HEIGHT:
                dim.height = DEFAULT_ROW_HEIGHT
                stats["heights_reset"] += 1
            # Source col: single-line row; link stays clickable even if clipped
            if source_col and _has_source(ws, r, source_col):
                c = ws.cell(r, source_col)
                if c.alignment is None:
                    c.alignment = Alignment(vertical="center")
                else:
                    c.alignment = Alignment(
                        horizontal=c.alignment.horizontal,
                        vertical="center",
                        wrap_text=False,
                        shrink_to_fit=c.alignment.shrink_to_fit,
                    )

    wb.save(tmp)
    tmp.replace(path)
    return stats


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    issues: list[str] = []

    scn = wb["Scenarios"]
    if scn["F12"].value != 0.045:
        issues.append(f"D&A % removed: F12={scn['F12'].value}")
    if scn["F13"].value != 0.055:
        issues.append(f"Capex % removed: F13={scn['F13'].value}")
    if scn["C13"].hyperlink is None:
        issues.append("Capex Source link missing on C13")
    if scn["C15"].hyperlink is None and not scn["C15"].value:
        issues.append("DSO Source missing on C15")

    tall = [
        (ws.title, r, dim.height)
        for ws in wb.worksheets
        for r, dim in ws.row_dimensions.items()
        if dim.height and dim.height > MAX_NORMAL_HEIGHT
    ]
    if tall:
        issues.append(f"{len(tall)} rows still inflated (e.g. {tall[0]})")

    cmt_on_sourced = 0
    for sheet in DOC_SHEETS:
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            if not _has_source(ws, r, SOURCE_COL):
                continue
            for c in range(1, ws.max_column + 1):
                if ws.cell(r, c).comment:
                    cmt_on_sourced += 1
    if cmt_on_sourced:
        issues.append(f"{cmt_on_sourced} hover comments remain on sourced rows")

    if issues:
        raise AssertionError("Verification failed:\n" + "\n".join(issues))

    total_cmt = sum(
        1 for ws in wb.worksheets for row in ws.iter_rows() for cell in row if cell.comment
    )
    print(
        f"Verify OK: assumptions intact; {total_cmt} comments kept (no Source col); "
        "no tall rows; no red triangles on sourced lines"
    )


if __name__ == "__main__":
    stats = remove_artifacts()
    print(
        f"Cleaned {TARGET.name}: removed {stats['comments_removed']} hover comments, "
        f"reset {stats['heights_reset']} row heights, kept {stats['comments_kept']} comments"
    )
    verify()
