#!/usr/bin/env python3
"""Sync slide 20 WACC × g sensitivity table directly from DCF sheet cells.

Reads DCF row 98 (terminal-g headers) and rows 99–103 (WACC × price grid)
after LibreOffice recalc. Rebuilds the table as the last deck mutation so GIS
formatting passes cannot blank the header row.

Run:  cd LULU && python3 scripts/sync_sensitivity_from_dcf.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
DECK = ROOT / "LULU_Investment_Pitch_Deck.pptx"
MODEL = ROOT / "model18_wsp_formulas.xlsx"

sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT.parent / "GIS"))

from gis_pitch import INK, LGREY, NAVY, WHITE, _set_font  # noqa: E402
from recalc_model18 import recalc_workbook  # noqa: E402


def _fmt_g(v: float) -> str:
    pct = float(v) * 100
    if abs(pct - round(pct, 1)) < 1e-9:
        return f"{pct:.1f}%"
    return f"{pct:.2f}%"


def read_dcf_sensitivity(dcf) -> tuple[list[str], list[list[str]]]:
    """Pull headers + body straight from DCF grid (row 98 / rows 99–103)."""
    headers = ["WACC \\ g"] + [_fmt_g(dcf.cell(98, c).value) for c in range(6, 11)]
    rows: list[list[str]] = []
    for r in range(99, 104):
        w = dcf.cell(r, 1).value
        if not isinstance(w, (int, float)):
            continue
        prices = [dcf.cell(r, c).value for c in range(6, 11)]
        if not all(isinstance(p, (int, float)) for p in prices):
            continue
        rows.append([f"{w * 100:.1f}%"] + [f"${round(p)}" for p in prices])
    return headers, rows


def _delete_shape(shape) -> None:
    shape.element.getparent().remove(shape.element)


def _is_sensitivity_table(shape) -> bool:
    if not shape.has_table:
        return False
    h0 = shape.table.rows[0].cells[0].text.strip().replace(" vs ", " \\ ")
    if h0 in ("WACC \\ g", "WACC vs g"):
        return True
    # legacy blank-header table parked at sensitivity position
    return abs(shape.top - Inches(5.62)) < 50000 and len(shape.table.rows) in (5, 6)


def _write_cell(cell, text: str, *, size: float, color, bold: bool, fill) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if cell._tc.getparent().index(cell._tc) > 0 else PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    _set_font(r, size, color, bold=bold)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)


def _build_table(slide, headers: list[str], rows: list[list[str]]) -> None:
    ncol = len(headers)
    nrows = len(rows) + 1
    top = Inches(5.62)
    left = Inches(0.50)
    width = Inches(12.35)
    height = Inches(round(0.20 * nrows, 2))

    shape = slide.shapes.add_table(nrows, ncol, left, top, width, height)
    table = shape.table
    col0w = Inches(1.0)
    table.columns[0].width = int(col0w)
    rest = int((width - col0w) / (ncol - 1))
    for c in range(1, ncol):
        table.columns[c].width = rest

    row_h = int(height / nrows)
    for row in table.rows:
        row.height = row_h

    for c, htxt in enumerate(headers):
        fill = NAVY if c == 0 else LGREY
        color = WHITE if c == 0 else NAVY
        _write_cell(table.cell(0, c), htxt, size=10, color=color, bold=True, fill=fill)

    for ri, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            fill = WHITE if ri % 2 else LGREY
            align_left = c == 0
            cell = table.cell(ri, c)
            _write_cell(cell, val, size=9, color=INK, bold=(c == 0), fill=fill)
            if not align_left:
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT


def sync(path: Path = DECK, model: Path = MODEL) -> dict:
    wb = recalc_workbook(model)
    headers, rows = read_dcf_sensitivity(wb["DCF"])

    prs = Presentation(str(path))
    slide = prs.slides[19]
    for shape in list(slide.shapes):
        if _is_sensitivity_table(shape):
            _delete_shape(shape)

    _build_table(slide, headers, rows)
    prs.save(str(path))

    return {"headers": headers, "rows": rows, "deck": str(path)}


if __name__ == "__main__":
    out = sync()
    print(f"Synced slide 20 sensitivity → {out['deck']}")
    print("headers:", out["headers"])
    for row in out["rows"]:
        print(" ", row)
