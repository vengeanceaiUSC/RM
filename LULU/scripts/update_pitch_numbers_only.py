#!/usr/bin/env python3
"""Update ONLY numeric values in the pitch deck to match recalculated Excel DCF.

Preserves all original slide copy verbatim — no narrative rewrites.
Run:  cd LULU && python3 scripts/update_pitch_numbers_only.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

from pptx import Presentation

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))

from recalc_model18 import read_dcf_outputs, recalc_workbook

SRC = ROOT / "LULU_Investment_Pitch_Deck.pptx"
FALLBACK = Path("/tmp/orig_deck.pptx")
OUT = ROOT / "LULU_Investment_Pitch_Deck.pptx"


def _replace_text(text: str, mapping: list[tuple[str, str]]) -> str:
    for old, new in mapping:
        text = text.replace(old, new)
    text = re.sub(r"\$134(?!\.\d)", "$133.52", text)
    return text


def _fmt_m(v: int) -> str:
    return f"{v:,}"


def main() -> None:
    wb = recalc_workbook(ROOT / "model18_wsp_formulas.xlsx")
    data = read_dcf_outputs(wb)
    base = data["scenarios"]["base"]
    bear = data["scenarios"]["bear"]
    bull = data["scenarios"]["bull"]
    bridge = data["bridge"]
    b = data["base"]

    px = f"{base['implied_px']:.2f}"
    bear_px = f"{bear['implied_px']:.2f}"
    bull_px = f"{bull['implied_px']:.2f}"
    upside = f"{base['upside_pct']:.1f}"
    eps = b["eps"]

    mapping = [
        # base DCF
        ("$133.64", f"${px}"),
        ("133.64", px),
        ("(+33.6% upside)", f"(+{upside}% upside)"),
        ("33.6% upside", f"{upside}% upside"),
        # bear / bull — exact Excel recalc (rows 131–133)
        ("$54.00 bear", f"${bear_px} bear"),
        ("$54 bear", f"${bear_px} bear"),
        ("$56.00 bear", f"${bear_px} bear"),
        ("$56 bear", f"${bear_px} bear"),
        ("Bear $54", f"Bear ${bear_px}"),
        ("Bear $56", f"Bear ${bear_px}"),
        ("bear $54", f"bear ${bear_px}"),
        ("bear $56", f"bear ${bear_px}"),
        ("$54 / bull", f"${bear_px} / bull"),
        ("Bear $54 / bull $208", f"Bear ${bear_px} / bull ${bull_px}"),
        ("Bear $56 / bull $202", f"Bear ${bear_px} / bull ${bull_px}"),
        ("bull $208", f"bull ${bull_px}"),
        ("bull $202", f"bull ${bull_px}"),
        ("$208", f"${bull_px}"),
        # EPS
        ("$14.57", f"${eps[4]:.2f}"),
        ("14.57", f"{eps[4]:.2f}"),
        ("$10.09", f"${eps[1]:.2f}"),
        ("10.09", f"{eps[1]:.2f}"),
        # EV bridge
        ("3,944", _fmt_m(bridge["pv_fcf_m"])),
        ("14,876", _fmt_m(bridge["ev_m"])),
        ("14,885", _fmt_m(bridge["equity_m"])),
        # IS / CF tables
        ("1,087", _fmt_m(b["ni_m"][0])),
        ("1,057", _fmt_m(b["ni_m"][1])),
        ("1,056", _fmt_m(b["fcf_m"][3])),
    ]

    # Update in place when deck exists; otherwise seed from GIS fallback copy.
    if not OUT.exists() and FALLBACK.exists():
        shutil.copy2(FALLBACK, OUT)
    prs = Presentation(str(OUT))

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text:
                            run.text = _replace_text(run.text, mapping)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text:
                            cell.text = _replace_text(cell.text, mapping)

    # Slide 14 — income statement forecast cols FY26-30
    for shape in prs.slides[13].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            t = shape.table
            rows_map = {
                1: b["revenue_m"],
                3: b["ebit_m"],
                5: b["ni_m"],
                6: eps,
            }
            for ri, vals in rows_map.items():
                for ci, val in enumerate(vals, start=5):
                    if ri == 6:
                        t.cell(ri, ci).text = f"{val:.2f}"
                    else:
                        t.cell(ri, ci).text = _fmt_m(int(val))

    # Slide 16 — cash flow FCF row
    for shape in prs.slides[15].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            t = shape.table
            for ci, val in enumerate(b["fcf_m"], start=5):
                t.cell(4, ci).text = _fmt_m(val)

    prs.save(str(OUT))
    print(f"Updated numbers only → {OUT}")
    print(f"  Base ${px} (+{upside}%) | Bear ${bear_px} | Bull ${bull_px}")
    print(f"  EPS FY26-FY30: {eps}")
    print(f"  FCF FY29: {b['fcf_m'][3]}M | NI FY27: {b['ni_m'][1]}M")


if __name__ == "__main__":
    main()
