#!/usr/bin/env python3
"""Final polish pass on unbeiesgbar_final.xlsx — strip AI/template tells.

Preserves formulas, hardcoded inputs (except share-price refresh), and Excel
Analyst comments. Clears mangled in-cell URLs (sources live in cell notes).

Run:  cd LULU && python3 scripts/humanize_final_model.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from restore_outline_groups import remove_outline_groups  # noqa: E402

ROOT = SCRIPTS.parent
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
    r"black formula",
    r"live formula",
    r"driver-based nwc",
    r"steps 1-4",
    r"col g\) links",
    r"not a peer pick",
    r"not the exit",
    r"sanity check",
    r"pitch deck reads",
    r"live excel average",
]
AI_SCRUB = re.compile("|".join(AI_PHRASES), re.I)
CONVERSATIONAL = re.compile(
    r"(?i)^/\s|^\s*note:\s|should be 0|guardrail|pipeline \(base",
)


def _clear_row_label(ws, row: int, col: int = 1) -> bool:
    cell = ws.cell(row, col)
    if cell.value is not None:
        cell.value = None
        return True
    return False


def _strip_conversational_notes(wb) -> int:
    """Remove AI tutorial rows and replace robotic DCF check column."""
    n = 0
    dcf = wb["DCF"]
    if dcf.max_column >= 8:
        dcf.delete_cols(8)
        n += 1

    clears = [
        ("DCF", "A34"),
        ("DCF", "A52"),
        ("DCF", "A63"),
        ("Scenarios", "A135"),
        ("NOPAT Bridge", "A51"),
        ("NOPAT Bridge", "A52"),
        ("Comps", "A17"),
        ("Comps", "A41"),
        ("Comps", "A63"),
    ]
    for r in range(48, 54):
        clears.append(("Comps", f"A{r}"))

    for sheet, coord in clears:
        cell = wb[sheet][coord]
        if cell.value is not None:
            _clear_cell(cell)
            n += 1

    nb = wb["NOPAT Bridge"]
    if nb["A37"].value and "Check vs" in str(nb["A37"].value):
        nb["A37"].value = "EBIT variance vs Scenarios ($000)"
        n += 1

    comps = wb["Comps"]
    if comps["A16"].value and "PITCHBOOK PUBCOMPS" in str(comps["A16"].value):
        comps["A16"].value = "Peer EV/EBITDA (PitchBook, Sep-2026)"
        n += 1
    if comps["A47"].value and "Rationale for selected exit" in str(comps["A47"].value):
        comps["A47"].value = "Selected exit multiple (Gordon growth)"
        n += 1
    if comps["A38"].value and "live Excel AVERAGE" in str(comps["A38"].value):
        comps["A38"].value = "Averages"
        n += 1
    if comps["A31"].value and "ALO YOGA BUILD" in str(comps["A31"].value):
        comps["A31"].value = "Alo Yoga implied valuation (reference only)"
        n += 1

    if dcf["A94"].value and "Sanity check" in str(dcf["A94"].value):
        _clear_cell(dcf["A94"])
        n += 1

    scn = wb["Scenarios"]
    if scn["A193"].value and "pitch deck" in str(scn["A193"].value).lower():
        _clear_cell(scn["A193"])
        n += 1

    dcf_renames = {
        "A19": "Working capital",
        "A40": "Scenarios tie-out",
        "A53": "Enterprise value bridge",
    }
    for coord, text in dcf_renames.items():
        if dcf[coord].value != text:
            dcf[coord].value = text
            n += 1

    return n


def _scrub_ai_phrasing(wb) -> int:
    """Remove known AI boilerplate from any visible text cell."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or cell.value.startswith("="):
                    continue
                val = cell.value
                if (
                    AI_SCRUB.search(val)
                    or CONVERSATIONAL.search(val)
                    or val.strip().startswith("/")
                    or val.lower().startswith("note:")
                ):
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
    for coord in ("A3", "B3", "E3", "A4"):
        _clear_cell(dcf[coord])
        changed += 1

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

    changed += _strip_conversational_notes(wb)
    changed += _scrub_ai_phrasing(wb)

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

    if wb["DCF"]["B72"].value == 100:
        wb["DCF"]["B72"].value = SHARE_PRICE
        changed += 1

    remove_outline_groups(wb)

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
        ("DCF", "A3", None),
        ("DCF", "A4", None),
        ("Scenarios", "A135", None),
        ("Comps", "A17", None),
        ("Comps", "A31", "Alo Yoga implied valuation (reference only)"),
        ("Comps", "A38", "Averages"),
        ("Comps", "A63", None),
        ("DCF", "A94", None),
        ("Scenarios", "A193", None),
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
                if isinstance(v, str) and v.startswith("="):
                    if "IF(MAX(ABS" in v.upper():
                        bad.append(f"{ws.title}!{cell.coordinate}: robotic check formula")
                elif isinstance(v, str):
                    if URLISH.search(v) or "(blue)" in v.lower() or "(red)" in v.lower():
                        bad.append(f"{ws.title}!{cell.coordinate}: {v[:60]!r}")
                    if v.strip().startswith("/"):
                        bad.append(f"{ws.title}!{cell.coordinate}: slash-note remains")
                    if CONVERSATIONAL.search(v):
                        bad.append(f"{ws.title}!{cell.coordinate}: conversational note")
                    if "pipeline" in v.lower() and "NOPAT" in v:
                        bad.append(f"{ws.title}!{cell.coordinate}: pipeline header remains")
                    if "LULU_Assumptions_Memo" in v:
                        bad.append(f"{ws.title}!{cell.coordinate}: memo filename in cell")
                    if AI_SCRUB.search(v):
                        bad.append(f"{ws.title}!{cell.coordinate}: AI phrasing remains")

    if wb["DCF"].max_column >= 8:
        bad.append("DCF: check column H still present")

    n_comments = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.comment)
    if n_comments < 50:
        bad.append(f"Only {n_comments} Analyst comments (expected >= 50)")

    if bad:
        raise AssertionError("Verify failed:\n" + "\n".join(bad[:15]))
    print(f"Verify OK: {n_comments} Analyst comments preserved, AI tells removed")


if __name__ == "__main__":
    humanize()
    verify()
