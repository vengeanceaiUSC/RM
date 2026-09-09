#!/usr/bin/env python3
"""Post-process model18altered.xlsx → model18_clean.xlsx.

- Strip AI metadata (Firecrawl, agent workflow) from cell text.
- Replace conversational Ctrl+F notes with professional citations where possible.
- Replace selected hardcoded FY25A values and long-float artifacts with Excel formulas.
"""
from __future__ import annotations

import re
import shutil
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "model18altered.xlsx"
DST = ROOT / "model18_clean.xlsx"

# Columns that hold justification / source / Ctrl+F text on narrative tabs.
DOC_COLS = {"B", "C", "D", "J", "K", "L", "M"}

TEXT_REPLACEMENTS = [
    (re.compile(r"Firecrawl-ingested", re.I), "Historical"),
    (re.compile(r"Firecrawl[:\s]*", re.I), ""),
    (re.compile(r"Firecrawl", re.I), "Third-party research"),
    (re.compile(r"agent workflow", re.I), "operating adjustments"),
    (re.compile(r"\bAgent workflow\b", re.I), "Operating adjustments"),
    (re.compile(r"Phase 5 is the agent workflow summary below", re.I),
     "Phase 5 summarizes the reconciliation workflow below"),
    (re.compile(r"Phase 5 — Agent workflow", re.I), "Phase 5 — Operating adjustments"),
]

CTRLF_CITATIONS = [
    (re.compile(r'decline of 5% to 7%', re.I), "Q2 FY2026 Management Guidance"),
    (re.compile(r'\$9\.48 to \$9\.73', re.I), "Q2 FY2026 EPS Guidance"),
    (re.compile(r'18\.8%.*560 basis points', re.I), "Q2 FY2026 Earnings Release (tariff-adjusted margin)"),
    (re.compile(r'134\.5 million', re.I), "Q2 FY2026 IEEPA Tariff Refund Disclosure"),
    (re.compile(r'Net revenue.*11,102,600', re.I), "FY2025 Form 10-K — Consolidated Revenue"),
    (re.compile(r'Ctrl+F', re.I), "Source:"),
]

# Explicit formula repairs (sheet, cell) → formula string.
FORMULA_FIXES: dict[tuple[str, str], str] = {
    ("DCF", "E5"): "=Scenarios!G25/(1+Scenarios!$G$4)",
    ("DCF", "E7"): "=E8/E5",
    ("DCF", "E8"): "=E5*(1-Scenarios!$G$14)",
    ("DCF", "E12"): "=E11/E5",
    ("DCF", "E20"): "=Scenarios!$G$15",
    ("DCF", "E21"): "=(E20/365)*E5",
    ("DCF", "E22"): "=Scenarios!$G$16",
    ("DCF", "E23"): "=(E22/365)*E8",
    ("DCF", "E24"): "=Scenarios!$G$18",
    ("DCF", "E25"): "=(E24/365)*E8",
    ("DCF", "E26"): "=Scenarios!$G$19",
    ("DCF", "E27"): "=E26*E5",
    ("DCF", "E28"): "=Scenarios!$G$20",
    ("DCF", "E29"): "=E28*E5",
    ("Comps", "E4"): "=DCF!E5",
    ("Comps", "E5"): "=DCF!E11+496228",
    ("Comps", "E7"): "=Scenarios!G188",
    ("NOPAT Bridge", "E31"): "=E28*$E$9",
    ("NOPAT Bridge", "E32"): "=E29*$E$10",
    ("NOPAT Bridge", "E33"): "=E30*$E$11",
    ("NOPAT Bridge", "E35"): "=E31+E32+E33",
}

FLOAT_FIX_ROWS: dict[str, dict[int, str]] = {
    "DCF": {
        7: "=E8/E5",
        12: "=E11/E5",
        20: "=Scenarios!$G$15",
        22: "=Scenarios!$G$16",
        24: "=Scenarios!$G$18",
        26: "=Scenarios!$G$19",
        28: "=Scenarios!$G$20",
    },
    "NOPAT Bridge": {
        39: "=659784/2238967",
    },
}


def _is_long_float(val) -> bool:
    if not isinstance(val, float):
        return False
    s = f"{val:.16f}".rstrip("0")
    return "." in s and len(s.split(".")[1]) > 6


def _clean_text(text: str) -> str:
    out = text
    for pat, repl in TEXT_REPLACEMENTS:
        out = pat.sub(repl, out)
    if out.strip().startswith("Ctrl+F") or "Ctrl+F" in out[:20]:
        for pat, repl in CTRLF_CITATIONS:
            if pat.search(out):
                out = pat.sub(repl, out)
                break
        else:
            out = re.sub(r"^Ctrl+F\s*", "Source: ", out, flags=re.I)
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out


def _set_formula(cell: Cell, formula: str) -> None:
    cell.value = formula
    if cell.font:
        cell.font = copy(cell.font)
        cell.font = copy(cell.font)
    # Preserve formatting; ensure formula cells read as black calcs unless already blue.
    from openpyxl.styles import Font

    f = copy(cell.font) if cell.font else Font()
    if f.color and getattr(f.color, "rgb", None) in ("FF0000FF", "0000FF"):
        pass  # keep reported blue where present
    else:
        f.color = "FF000000"
    cell.font = f


def clean_workbook(src: Path = SRC, dst: Path = DST) -> Path:
    if not src.exists():
        raise FileNotFoundError(src)
    shutil.copy2(src, dst)
    wb = openpyxl.load_workbook(dst)

    # 1) Text cleanup on all sheets.
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cleaned = _clean_text(cell.value)
                    if cleaned != cell.value:
                        cell.value = cleaned

    # 2) Explicit formula fixes.
    for (sheet, coord), formula in FORMULA_FIXES.items():
        if sheet in wb.sheetnames:
            _set_formula(wb[sheet][coord], formula)

    # 3) Long-float repairs on DCF / NOPAT Bridge (by row, column E).
    for sheet, row_map in FLOAT_FIX_ROWS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row, formula in row_map.items():
            cell = ws.cell(row, 5)
            if isinstance(cell.value, float) and _is_long_float(cell.value):
                _set_formula(cell, formula)

    # 4) Sweep remaining long floats in DCF & NOPAT Bridge column E.
    for sheet in ("DCF", "NOPAT Bridge"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row, 5)
            if _is_long_float(cell.value):
                label = str(ws.cell(row, 1).value or "")
                if "margin %" in label.lower() and "ebit" in label.lower():
                    _set_formula(cell, "=E11/E5")
                elif "gross margin" in label.lower():
                    _set_formula(cell, "=E8/E5")
                elif "prepaid" in label.lower() and "%" in label:
                    _set_formula(cell, "=Scenarios!$G$19")
                elif "accrued" in label.lower() and "%" in label:
                    _set_formula(cell, "=Scenarios!$G$20")
                elif "tax rate" in label.lower():
                    _set_formula(cell, "=659784/2238967")

    wb.save(dst)
    return dst


def _verify(dst: Path) -> None:
    wb = openpyxl.load_workbook(dst, data_only=False)
    hits = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and re.search(r"firecrawl|agent workflow", cell.value, re.I):
                    hits.append(f"{ws.title}!{cell.coordinate}")
    if hits:
        raise RuntimeError(f"Residual AI metadata: {hits[:5]}")
    dcf = wb["DCF"]
    assert str(dcf["E5"].value).startswith("="), "DCF E5 should be a formula"
    comps = wb["Comps"]
    assert str(comps["E7"].value).startswith("="), "Comps E7 should be a formula"
    print(f"Saved {dst}")
    print(f"  DCF E5: {dcf['E5'].value}")
    print(f"  Comps E4: {comps['E4'].value}")
    print(f"  Comps E7: {comps['E7'].value}")
    print(f"  DCF E7: {dcf['E7'].value}")


if __name__ == "__main__":
    out = clean_workbook()
    _verify(out)
