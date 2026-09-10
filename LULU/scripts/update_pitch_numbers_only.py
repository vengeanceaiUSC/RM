#!/usr/bin/env python3
"""Update ONLY numeric values in the pitch deck to match model18 DCF.

Preserves all original slide copy verbatim — no narrative rewrites.
Run:  cd LULU && python3 scripts/update_pitch_numbers_only.py
"""
from __future__ import annotations

import re
import shutil
import sys
from copy import copy
from pathlib import Path

from pptx import Presentation

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))

from eval_model18 import _col_inputs, _eval_scenario, evaluate

SRC = ROOT / "LULU_Investment_Pitch_Deck.pptx"
# Prefer branch deck with original GIS copy if present
FALLBACK = Path("/tmp/orig_deck.pptx")
OUT = ROOT / "LULU_Investment_Pitch_Deck.pptx"


def _sensitivity_grid() -> dict[tuple[str, str], int]:
    from openpyxl import load_workbook

    wb = load_workbook(ROOT / "model18_wsp_formulas.xlsx", data_only=False)
    scn, dcf = wb["Scenarios"], wb["DCF"]
    inp = _col_inputs(scn, dcf, "G")
    base_w = evaluate(ROOT / "model18_wsp_formulas.xlsx")["wacc"]
    grid = {}
    for w in (0.095, 0.10, 0.105, 0.11):
        for g in (0.015, 0.02, 0.0225, 0.025, 0.03):
            gi = dict(inp)
            gi["wacc"] = w
            gi["g10"] = g
            px = round(_eval_scenario(gi, base_w)["implied_px"])
            grid[(f"{w*100:.1f}%", f"{g*100:.2f}%".rstrip("0").rstrip(".") + "%" if g != 0.0225 else "2.25%")] = px
    # normalize keys for 2.25%
    out = {}
    for (w, g), px in grid.items():
        gkey = "2.25%" if abs(float(g.rstrip("%")) - 2.25) < 0.01 else g
        out[(w, gkey)] = px
    return out


def _replace_text(text: str, mapping: list[tuple[str, str]]) -> str:
    for old, new in mapping:
        text = text.replace(old, new)
    # $134 that is NOT $134.5M (DCF rounded references)
    text = re.sub(r"\$134(?!\.\d)", "$133.52", text)
    return text


def main() -> None:
    ev = evaluate(ROOT / "model18_wsp_formulas.xlsx")
    base = ev["scenarios"]["base"]
    bear = round(ev["scenarios"]["bear"]["implied_px"])
    bull = round(ev["scenarios"]["bull"]["implied_px"])
    upside = round(base["upside_pct"], 1)
    px = f"{base['implied_px']:.2f}"
    eps = base["eps"]

    dcf_cash = 1_807_202
    debt = 1_798_441
    eq_m = round((base["ev"] + dcf_cash - debt) / 1000)

    mapping = [
        ("$133.64", f"${px}"),
        ("133.64", px),
        ("(+33.6% upside)", f"(+{upside}% upside)"),
        ("33.6% upside", f"{upside}% upside"),
        ("$14.57", f"${eps[4]:.2f}"),
        ("14.57", f"{eps[4]:.2f}"),
        ("$10.09", f"${eps[1]:.2f}"),
        ("10.09", f"{eps[1]:.2f}"),
        ("$56.00 bear", f"${bear}.00 bear"),
        ("Bear $56", f"Bear ${bear}"),
        ("bull $202", f"bull ${bull}"),
        ("$202 bracket", f"${bull} bracket"),
        ("3,944", f"{round(base['pv_explicit']/1000):,}"),
        ("14,876", f"{round(base['ev']/1000):,}"),
        ("14,885", f"{eq_m:,}"),
        ("1,087", "1,086"),
        ("1,056", "1,057"),
    ]

    src = FALLBACK if FALLBACK.exists() else SRC
    shutil.copy2(src, OUT)
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

    # EPS table row (slide 14) — precise
    is_slide = prs.slides[13]
    for shape in is_slide.shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            cells = shape.table.rows[6].cells
            for ci, val in enumerate(eps, start=5):
                cells[ci].text = f"{val:.2f}"

    # Sensitivity table slide 20
    sens = _sensitivity_grid()
    dcf_slide = prs.slides[19]
    for shape in dcf_slide.shapes:
        if shape.has_table and shape.table.cell(0, 0).text.startswith("WACC"):
            for ri in range(1, 5):
                w = shape.table.cell(ri, 0).text
                for ci, g in enumerate(["1.5%", "2.0%", "2.25%", "2.5%", "3.0%"], start=1):
                    key = (w, g)
                    if key in sens:
                        shape.table.cell(ri, ci).text = f"${sens[key]}"

    prs.save(str(OUT))
    print(f"Updated numbers only → {OUT}")
    print(f"  Base DCF ${px} (+{upside}%) | Bear ${bear} | Bull ${bull}")
    print(f"  EPS FY26-FY30: {[round(x,2) for x in eps]}")


if __name__ == "__main__":
    main()
