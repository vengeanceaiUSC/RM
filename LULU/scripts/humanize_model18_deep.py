#!/usr/bin/env python3
"""Deep club-submission humanization for model18_humanized.xlsx.

Removes structural AI tells while preserving hardcodes, Source columns, and
core formulas:

- Delete Revenue Drivers Equation column (I)
- Delete DCF Δ vs Scenarios check column (K) and Alt source/Ctrl+F cols (L–M)
- Rewrite NOPAT Bridge phase labels to standard finance naming
- Delete memo / meta-commentary rows (memo:, black formula, formatting notes)
- Scrub conversational AI phrasing from remaining labels

Output: model18_humanized (1).xlsx
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from polish_model18_altered import (  # noqa: E402
    _delete_col_and_fix_refs,
    _rewrite_all_formulas,
    _rewrite_sheet_internal_formulas,
)

ROOT = SCRIPTS.parent
INPUT = ROOT / "model18_humanized.xlsx"
OUTPUT = ROOT / "model18_humanized (1).xlsx"

LABEL_RENAMES: list[tuple[re.Pattern[str], str | None]] = [
    (re.compile(r"Reported EBIT\s*→\s*Normalized NOPAT\s*\(5-phase pipeline\)", re.I), "NOPAT bridge"),
    (re.compile(r"Phase 1:\s*Add back SBC\?\s*\(0 = no — Convention A:.*", re.I), "SBC add-back (0)"),
    (re.compile(r"Convention A:.*", re.I), None),
    (re.compile(r"Phase 1\s*[—–-]\s*Operating normalization", re.I), "Reported EBIT adjustments"),
    (re.compile(r"Phase 2\s*[—–-]\s*Lease & capitalization.*", re.I), "Lease & R&D adjustments"),
    (re.compile(r"Phase 3\s*[—–-]\s*Segment / channel EBIT.*", re.I), "Channel EBIT mix"),
    (re.compile(r"Phase 4\s*[—–-]\s*Tax normalization.*", re.I), "Tax normalization"),
    (re.compile(r"Phase 2:\s*R&D / software amortization period.*", re.I), "R&D amortization period (years)"),
    (re.compile(r"= Adjusted EBIT \(Phase 1\)", re.I), "Adjusted EBIT"),
    (re.compile(r"EBIT_adj \(Phase 1\)", re.I), "Adjusted EBIT"),
    (re.compile(r"EBIT_capitalized \(Phase 2\)", re.I), "EBIT after lease/R&D"),
    (re.compile(r"EBIT_lease-adj \(Phase 2\)", re.I), "Lease-adjusted EBIT"),
    (re.compile(r"Forecast EBIT \(Phase 3\)", re.I), "Forecast EBIT"),
    (re.compile(r"Normalized EBIT \(Phase 4\)", re.I), "Normalized EBIT"),
    (re.compile(r"NOPAT Bridge tab: Phase \d subtotal", re.I), "Subtotal"),
    (re.compile(r"Phases 1[–-]4 clean and forecast operating profit.*", re.I), None),
    (re.compile(r"BOTTOM-UP REVENUE DRIVER SCHEDULE.*", re.I), "Revenue drivers (FY26–30)"),
    (re.compile(r"NOPAT NORMALIZATION PIPELINE.*", re.I), "NOPAT bridge"),
    (re.compile(r"Forecast drivers linked to Scenarios tab.*", re.I), "Forecast linked to Scenarios (base case)"),
    (re.compile(r"Full Assumptions Guide \(PDF\).*", re.I), None),
    (re.compile(r"Scenarios linkage summary.*", re.I), None),
    (re.compile(r"EV/EBITDA is a black formula.*", re.I), None),
    (re.compile(r"Column E on the next two rows is \$bn.*", re.I), None),
    (re.compile(r"ALO YOGA BUILD.*", re.I), None),
    (re.compile(r"^\s*memo:\s*", re.I), None),
    (re.compile(r"^\(\d+\)\s+", re.I), ""),
    (re.compile(r"\(Phase [1-4]\)", re.I), ""),
    (re.compile(r"Phase [1-5]:\s*", re.I), ""),
]

ROW_DELETE_PATTERNS = [
    re.compile(r"^\s*memo:", re.I),
    re.compile(r"black formula", re.I),
    re.compile(r"equation \(every formula row\)", re.I),
    re.compile(r"pitch deck pulls cols", re.I),
    re.compile(r"β walkthrough", re.I),
    re.compile(r"font / color convention", re.I),
    re.compile(r"justification \| source", re.I),
    re.compile(r"^\(\d+\)\s+yahoo", re.I),
    re.compile(r"^\(\d+\)\s+unlever", re.I),
    re.compile(r"^\(\d+\)\s+relever", re.I),
    re.compile(r"^book d/e 44", re.I),
]

CONVERSATIONAL_RE = re.compile(
    r"nothing fancy|this whole tab|flips that to \$000|that's where the year|"
    r"same number, just carried|×1m 'cause|not a 1:1 rate",
    re.I,
)

CTRLF_IN_SOURCE = re.compile(r"ctrl\+f|5-phase|also:", re.I)


def _find_source_cols(ws) -> list[int]:
    cols: list[int] = []
    for row in range(1, 7):
        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(row, col).value or "").strip().lower()
            if val in {"source (click)", "source"} or val.startswith("source"):
                cols.append(col)
    return sorted(set(cols))


def _clean_source_columns(wb) -> int:
    changed = 0
    replacements = [
        (re.compile(r"Ctrl\+F.*", re.I), ""),
        (re.compile(r"5-phase EBIT normalization.*", re.I), ""),
        (re.compile(r"Also:.*", re.I), ""),
        (re.compile(r"\n{2,}"), "\n"),
    ]
    for ws in wb.worksheets:
        for col in _find_source_cols(ws):
            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row, col)
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                if not CTRLF_IN_SOURCE.search(cell.value):
                    continue
                out = cell.value
                for pat, repl in replacements:
                    out = pat.sub(repl, out)
                out = re.sub(r"\s{2,}", " ", out).strip()
                if out != cell.value:
                    cell.value = out if out else None
                    changed += 1
    return changed


def _apply_label(text: str) -> str | None:
    out = text
    for pat, repl in LABEL_RENAMES:
        if repl is None and pat.search(out):
            return None
        out = pat.sub(repl or "", out)
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out or None


def _row_should_delete(ws, row: int) -> bool:
    label = str(ws.cell(row, 1).value or "").strip()
    if not label:
        return False
    return any(p.search(label) for p in ROW_DELETE_PATTERNS)


def _delete_rows(ws, rows: list[int]) -> int:
    for row in sorted(rows, reverse=True):
        ws.delete_rows(row, 1)
    return len(rows)


def _rename_all_labels(wb) -> int:
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                new = _apply_label(cell.value)
                if new != cell.value:
                    cell.value = new
                    changed += 1
    return changed


def _clear_meta_rows(wb) -> int:
    """Clear instructional row labels; do not delete rows (preserves formula refs)."""
    cleared = 0
    for ws in wb.worksheets:
        if ws.title == "Cover":
            rows = [r for r in range(1, ws.max_row + 1) if _row_should_delete(ws, r)]
            _delete_rows(ws, rows)
            cleared += len(rows)
            continue
        for row in range(1, ws.max_row + 1):
            if not _row_should_delete(ws, row):
                continue
            for col in (1, 2):
                cell = ws.cell(row, col)
                if isinstance(cell.value, str) and not cell.value.startswith("="):
                    cell.value = None
                    cleared += 1
    return cleared


def _scrub_conversational(wb) -> int:
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                if CONVERSATIONAL_RE.search(cell.value):
                    cell.value = None
                    changed += 1
    return changed


def _fix_rd_revenue_refs(wb) -> int:
    if "NOPAT Bridge" not in wb.sheetnames:
        return 0
    nb = wb["NOPAT Bridge"]
    n = 0
    row_map = {28: 31, 29: 41, 30: 45}
    col_map = {"F": "C", "G": "D", "H": "E", "I": "F", "J": "G"}
    for nb_row, rd_row in row_map.items():
        for nb_col, rd_col in col_map.items():
            want = f"='Revenue Drivers'!{rd_col}{rd_row}"
            cell = nb[f"{nb_col}{nb_row}"]
            if cell.value != want:
                cell.value = want
                n += 1
    return n


def _delete_columns(wb) -> dict[str, list[str]]:
    deleted: dict[str, list[str]] = {}
    if "Revenue Drivers" in wb.sheetnames:
        _delete_col_and_fix_refs(wb, "Revenue Drivers", "I")
        deleted["Revenue Drivers"] = ["I (Equation)"]
    if "DCF" in wb.sheetnames:
        for col in ("M", "L", "K"):
            _delete_col_and_fix_refs(wb, "DCF", col)
        deleted["DCF"] = ["M (Alt. Ctrl+F)", "L (Alt. source)", "K (Δ vs Scenarios)"]
    return deleted


def _clean_rd_banner(ws) -> None:
    if isinstance(ws["A2"].value, str):
        ws["A2"].value = re.sub(
            r"Historical data ingested 10-K anchors \(blue\)\.\s*",
            "FY25 reported figures (blue). ",
            ws["A2"].value,
            flags=re.I,
        )


def humanize(src: Path = INPUT, dst: Path = OUTPUT) -> Path:
    if not src.exists():
        raise FileNotFoundError(src)
    tmp = dst.with_suffix(".deep.tmp.xlsx")
    shutil.copy2(src, tmp)
    wb = openpyxl.load_workbook(tmp)

    stats = {
        "labels_renamed": _rename_all_labels(wb),
        "meta_rows_cleared": _clear_meta_rows(wb),
        "conversational_cleared": _scrub_conversational(wb),
    }

    if "Revenue Drivers" in wb.sheetnames:
        _clean_rd_banner(wb["Revenue Drivers"])

    stats["columns_deleted"] = _delete_columns(wb)
    stats["rd_revenue_refs_fixed"] = _fix_rd_revenue_refs(wb)
    stats["source_cells_cleaned"] = _clean_source_columns(wb)

    wb.save(tmp)
    tmp.replace(dst)

    print(f"Saved {dst}")
    for key, val in stats.items():
        print(f"  {key}: {val}")
    return dst


def verify(src: Path, dst: Path) -> None:
    dst_wb = openpyxl.load_workbook(dst, data_only=False)
    banned = re.compile(
        r"equation \(every|nothing fancy|this whole tab|convention a|"
        r"5-phase pipeline|phase [1-5]\s*[—–-]|black formula|Δ vs Scenarios|"
        r"ctrl\+f|memo:|β walkthrough",
        re.I,
    )

    if "Revenue Drivers" in dst_wb.sheetnames:
        ws = dst_wb["Revenue Drivers"]
        for row in range(1, 6):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row, col).value
                if isinstance(val, str) and "equation" in val.lower():
                    raise AssertionError(
                        f"Equation header remains: {get_column_letter(col)}{row}"
                    )

    if "DCF" in dst_wb.sheetnames:
        ws = dst_wb["DCF"]
        for row in range(1, 6):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row, col).value
                if isinstance(val, str) and "Δ vs Scenarios" in val:
                    raise AssertionError("DCF check column header still present")

    # Key hardcodes unchanged
    checks = [
        ("WACC", "E17", 0.86),
        ("WACC", "E35", 0.05),
        ("DCF", "E5", 11102600),
        ("Scenarios", "G25", None),  # formula ok
    ]
    src_wb = openpyxl.load_workbook(src, data_only=False)
    for sheet, coord, expected in checks:
        if expected is None:
            continue
        sv = src_wb[sheet][coord].value
        dv = dst_wb[sheet][coord].value
        if sv != dv:
            raise AssertionError(f"Hardcode moved: {sheet}!{coord} {sv} -> {dv}")

    for ws in dst_wb.worksheets:
        source_cols = set(_find_source_cols(ws))
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                if cell.column in source_cols:
                    continue
                if banned.search(cell.value):
                    raise AssertionError(
                        f"Residual AI tell: {ws.title}!{cell.coordinate}: {cell.value[:80]}"
                    )

    for sheet, col in [("WACC", 3), ("Scenarios", 3), ("DCF", 3), ("Comps", 3)]:
        ws = dst_wb[sheet]
        nonempty = sum(
            1
            for r in range(3, min(ws.max_row, 60))
            if ws.cell(r, col).value not in (None, "", "Source (click)", "Source")
        )
        if nonempty < 3:
            raise AssertionError(f"Source column too sparse on {sheet}")

    print("Verification passed.")


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT
    out = humanize(src)
    verify(src, out)
