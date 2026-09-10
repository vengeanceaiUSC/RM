#!/usr/bin/env python3
"""Final polish pass on unbeiesgbar_final.xlsx — strip AI/template tells.

Preserves formulas, hardcoded inputs (except share-price refresh), and Excel
Analyst comments. Clears mangled in-cell URLs (sources live in cell notes).

Run:  cd LULU && python3 scripts/humanize_final_model.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"
SHARE_PRICE = 100.61  # NASDAQ last sale per model source notes (Sep-2026)
URLISH = re.compile(r"https?\s*:\s*/", re.I)
AI_PHRASES = [
    r"hangs together",
    r"bottom-up path",
    r"stays viable",
    r"valuation anchor",
    r"revisit drivers or Scenarios",
    r"high-conviction analyst overlay",
]
AI_SCRUB = re.compile("|".join(AI_PHRASES), re.I)


def _scrub_ai_phrasing(wb) -> int:
    """Remove known AI boilerplate from any visible text cell."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                if AI_SCRUB.search(cell.value):
                    cell.value = None
                    n += 1
    return n



def _clear_cell(cell) -> None:
    cell.value = None
    if cell.hyperlink:
        cell.hyperlink = None
    if cell.comment and cell.comment.text and URLISH.search(cell.comment.text):
        pass  # keep Analyst notes with proper URLs


def _set_or_clear(cell, text: str | None) -> None:
    cell.value = text.strip() if text and text.strip() else None


def humanize(path: Path = TARGET) -> Path:
    tmp = path.with_suffix(".polishing.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    changed = 0

    # --- Cover ---
    cov = wb["Cover"]
    for coord in ("B9", "B17"):
        if cov[coord].value is not None:
            _clear_cell(cov[coord])
            changed += 1

    # --- Revenue Drivers ---
    rd = wb["Revenue Drivers"]
    _set_or_clear(rd["A2"], None)
    _clear_cell(rd["A59"])  # variance block is self-explanatory; no footnote needed
    changed += 2

    # --- NOPAT Bridge ---
    nb = wb["NOPAT Bridge"]
    _set_or_clear(nb["A1"], "NOPAT RECONCILIATION (BASE CASE)")
    _clear_cell(nb["A8"])
    changed += 2

    # --- DCF ---
    dcf = wb["DCF"]
    _set_or_clear(dcf["H2"], "Check")
    for coord in ("A3", "B3", "E3", "H3", "A4"):
        _clear_cell(dcf[coord])
        changed += 1
    renames = {
        "A19": "Working capital schedule",
        "A40": "Scenarios linkage",
        "A53": "UFCF pre-interest; ASC 842 lease debt in EV bridge.",
    }
    for coord, text in renames.items():
        if dcf[coord].value != text:
            dcf[coord].value = text
            changed += 1
    _clear_cell(dcf["A63"])

    for row in dcf.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and not cell.value.startswith("="):
                if URLISH.search(cell.value) or cell.value.strip().startswith("http"):
                    _clear_cell(cell)
                    changed += 1
            elif cell.hyperlink:
                cell.hyperlink = None
                changed += 1

    # --- Comps ---
    comps = wb["Comps"]
    for row in comps.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and not cell.value.startswith("="):
                if URLISH.search(cell.value) or cell.value.strip().startswith("http"):
                    _clear_cell(cell)
                    changed += 1
            elif cell.hyperlink:
                cell.hyperlink = None
                changed += 1
    _clear_cell(comps["A32"])
    _set_or_clear(comps["A41"], "Peer multiples illustrative; terminal value from DCF exit multiple.")

    # --- Share price refresh (market hardcodes only) ---
    price_rows = [
        ("WACC", "B8", "Share price ($)"),
        ("DCF", "B57", "Current share price"),
        ("Comps", "B10", "Current share price"),
        ("Comps", "B62", "Current price"),
    ]
    for sheet, coord, label in price_rows:
        ws = wb[sheet]
        cell = ws[coord]
        if cell.value != SHARE_PRICE:
            cell.value = SHARE_PRICE
            changed += 1

    # Repurchase schedule starting price — same market print
    if wb["DCF"]["B72"].value == 100:
        wb["DCF"]["B72"].value = SHARE_PRICE
        changed += 1

    changed += _scrub_ai_phrasing(wb)

    wb.save(tmp)
    tmp.replace(path)
    print(f"Humanized {path.name} ({changed} edits)")
    return path


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    bad: list[str] = []
    checks = [
        ("Cover", "B9", None),
        ("Cover", "B17", None),
        ("Revenue Drivers", "A2", None),
        ("NOPAT Bridge", "A1", "NOPAT RECONCILIATION (BASE CASE)"),
        ("DCF", "H2", "Check"),
        ("DCF", "A3", None),
        ("DCF", "A4", None),
        ("WACC", "B8", SHARE_PRICE),
    ]
    for sheet, coord, expected in checks:
        val = wb[sheet][coord].value
        if val != expected:
            bad.append(f"{sheet}!{coord}: {val!r} (expected {expected!r})")

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and not v.startswith("="):
                    if URLISH.search(v) or "(blue)" in v.lower() or "(red)" in v.lower():
                        bad.append(f"{ws.title}!{cell.coordinate}: {v[:60]!r}")
                    if "pipeline" in v.lower() and "NOPAT" in v:
                        bad.append(f"{ws.title}!{cell.coordinate}: pipeline header remains")
                    if "LULU_Assumptions_Memo" in v:
                        bad.append(f"{ws.title}!{cell.coordinate}: memo filename in cell")
                    if AI_SCRUB.search(v):
                        bad.append(f"{ws.title}!{cell.coordinate}: AI phrasing remains")

    n_comments = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.comment)
    if n_comments < 50:
        bad.append(f"Only {n_comments} Analyst comments (expected >= 50)")

    if bad:
        raise AssertionError("Verify failed:\n" + "\n".join(bad[:15]))
    print(f"Verify OK: {n_comments} Analyst comments preserved, AI tells removed")


if __name__ == "__main__":
    humanize()
    verify()
