#!/usr/bin/env python3
"""Full GIS layout + formatting compliance pass on LULU pitch deck.

Fixes overlap/bleed issues on slides 14–22, shortens footers, col G→C,
applies Garamond, and enforces GIS spacing.

Run:  cd LULU && python3 scripts/fix_pitch_gis_compliance.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
DECK = ROOT / "LULU_Investment_Pitch_Deck.pptx"

sys.path.insert(0, str(ROOT.parent / "GIS"))
from gis_pitch import CARD, FONT, GREY, INK, LGREY, NAVY, WHITE, _set_font  # noqa: E402

FIN_FOOTER = "Source: LULU_DCF_Valuation_Model.xlsx · Hist: 10-K · Forecast: Scenarios col C"
FOOTER_Y = Inches(7.12)
FOOTER_H = Inches(0.30)
FOOTER_W = Inches(11.65)


def _delete_shape(shape) -> None:
    shape.element.getparent().remove(shape.element)


def _set_footer(shape, text: str) -> None:
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text[:95]
    _set_font(r, 8, GREY)


def _footer_shapes(slide):
    return [
        s for s in slide.shapes
        if s.has_text_frame and s.top >= FOOTER_Y - 1000 and "Investment Research" not in s.text_frame.text
    ]


def _run_color(run) -> object:
    try:
        return run.font.color.rgb
    except AttributeError:
        return INK


def _shrink_table_font(table, size: float = 8.0, *, skip_header: bool = False) -> None:
    for ri, row in enumerate(table.rows):
        if skip_header and ri == 0:
            continue
        for cell in row.cells:
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    _set_font(
                        r, size, _run_color(r),
                        bold=bool(r.font.bold),
                        italic=bool(r.font.italic),
                    )


def _fix_navy_header_row(table, size: float = 9.0) -> None:
    for cell in table.rows[0].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                _set_font(r, size, WHITE, bold=True)


def _fix_financial_slides(prs: Presentation) -> None:
    """Slides 14–16: gap between data table and source table."""
    main_h = Inches(1.82)
    main_top = Inches(3.72)
    src_top = Inches(5.72)
    src_h = Inches(0.90)

    for si in (14, 15, 16):
        slide = prs.slides[si - 1]
        for shape in slide.shapes:
            if shape.has_table:
                h0 = shape.table.rows[0].cells[0].text.strip()
                if h0 == "US$ M":
                    shape.height = int(main_h)
                    shape.top = int(main_top)
                elif h0 == "Line item":
                    shape.top = int(src_top)
                    shape.height = int(src_h)
                    _compact_table_rows(shape.table, 0.90)
                    _shrink_table_font(shape.table, 8.0)
            if shape.has_text_frame and shape.top >= FOOTER_Y - 1000:
                t = shape.text_frame.text
                if t and not t.strip().isdigit():
                    _set_footer(shape, FIN_FOOTER)


def _compact_table_rows(table, total_h_in: float) -> None:
    """Set uniform row heights so declared height matches rendered content."""
    n = len(table.rows)
    rh = int(Inches(total_h_in) / n)
    for row in table.rows:
        row.height = rh


def _fix_slide17(prs: Presentation) -> None:
    slide = prs.slides[16]
    # Keep MODEL SOURCE MAP — shrink tables so they fit (never delete)
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        h0 = shape.table.rows[0].cells[0].text.strip()
        if h0 in ("Cap stack line item", "WACC input"):
            shape.width = int(Inches(6.0))
            shape.height = int(Inches(1.05))
            _compact_table_rows(shape.table, 1.05)
            _shrink_table_font(shape.table, 7.0)
    for shape in slide.shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text.strip() == "Component":
            # EV bridge table — cap height so note below clears
            if shape.top > Inches(2.5):
                shape.height = int(Inches(0.72))
                _compact_table_rows(shape.table, 0.72)
        if shape.has_text_frame and shape.text_frame.text.startswith("No funded bank debt"):
            shape.top = int(Inches(4.22))
            shape.height = int(Inches(0.48))
    for shape in _footer_shapes(slide):
        if not shape.text_frame.text.strip().isdigit():
            _set_footer(shape, "Source: LULU_DCF_Valuation_Model.xlsx, WACC tab (col E) · 10-K cap stack")


def _fix_slide20(prs: Presentation) -> None:
    """Slide 20 layout — keep base-case callout; sensitivity grid stays left (7.4in)."""
    slide = prs.slides[19]
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text
            if "SCENARIOS COL G" in t:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = r.text.replace("SCENARIOS COL G", "SCENARIOS COL C")
            if t.strip().startswith("TERMINAL VALUE"):
                shape.top = int(Inches(4.28))
                shape.height = int(Inches(0.22))
            if "SENSITIVITY" in t and "Base-case" in t:
                # Strip duplicate base-case line added by prior runs
                paras = [p for p in shape.text_frame.paragraphs if "Base-case" not in p.text]
                while len(shape.text_frame.paragraphs) > len(paras):
                    shape.text_frame._txBody.remove(shape.text_frame.paragraphs[-1]._p)
            if "SENSITIVITY" in t:
                shape.top = int(Inches(5.42))
                shape.left = int(Inches(0.50))
                shape.width = int(Inches(12.35))
                shape.height = int(Inches(0.18))
            if t.startswith("TV = 73%"):
                shape.top = int(Inches(4.82))
                shape.left = int(Inches(6.85))
                shape.width = int(Inches(6.00))
                shape.height = int(Inches(0.52))
                _shrink_text_frame(shape.text_frame, 7.5)
        if shape.has_table:
            h0 = shape.table.rows[0].cells[0].text.strip()
            if h0 == "Method":
                shape.top = int(Inches(4.02))
                shape.left = int(Inches(6.85))
                shape.width = int(Inches(6.00))
                shape.height = int(Inches(0.46))
                _compact_table_rows(shape.table, 0.46)
                _shrink_table_font(shape.table, 7.5)
            if h0 == "WACC vs g":
                shape.top = int(Inches(5.80))
                shape.left = int(Inches(0.50))
                shape.width = int(Inches(7.40))
                shape.height = int(Inches(1.08))
                _compact_table_rows(shape.table, 1.08)
                _shrink_table_font(shape.table, 9, skip_header=True)
                _fix_navy_header_row(shape.table, 9)
    for shape in _footer_shapes(slide):
        if not shape.text_frame.text.strip().isdigit():
            _set_footer(
                shape,
                "Source: LULU_DCF_Valuation_Model.xlsx, Scenarios col C + DCF · base WACC 9.0% g 2.25%",
            )


def _shrink_text_frame(tf, size: float = 8.5) -> None:
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            _set_font(
                r, size, _run_color(r),
                bold=bool(r.font.bold),
                italic=bool(r.font.italic),
            )


def _fix_slide22(prs: Presentation) -> None:
    slide = prs.slides[21]
    for shape in slide.shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text.strip() == "Target":
            shape.top = int(Inches(1.28))
            shape.height = int(Inches(1.68))
            _compact_table_rows(shape.table, 1.68)
            for row in shape.table.rows:
                for cell in row.cells:
                    _shrink_text_frame(cell.text_frame, 8.5)
        if shape.has_text_frame:
            t = shape.text_frame.text
            if "Private precedent framework" in t:
                shape.top = int(Inches(3.72))
                shape.left = int(Inches(0.50))
                shape.width = int(Inches(7.55))
                shape.height = int(Inches(2.35))
                _shrink_text_frame(shape.text_frame, 8.5)
            if t.strip().startswith("Valuation anchors"):
                shape.top = int(Inches(3.72))
                shape.left = int(Inches(8.30))
                shape.width = int(Inches(4.55))
                shape.height = int(Inches(2.35))
                _shrink_text_frame(shape.text_frame, 8.5)
    for shape in _footer_shapes(slide):
        if not shape.text_frame.text.strip().isdigit():
            _set_footer(shape, "Source: Wolverine IR, Bloomberg, BusinessWire (precedent context only)")


def _fix_other_footers(prs: Presentation) -> None:
    shorts = {
        18: "Source: LULU_DCF_Valuation_Model.xlsx, Comps / football field · SOTP FY30E",
        19: "Source: LULU_DCF_Valuation_Model.xlsx, SOTP · 10-K geo mix · Scenarios col C",
        21: "Source: PitchBook comps 04-Sep-2026 · LULU_DCF_Valuation_Model.xlsx",
    }
    for si, text in shorts.items():
        for shape in _footer_shapes(prs.slides[si - 1]):
            if not shape.text_frame.text.strip().isdigit():
                _set_footer(shape, text)


def _global_replacements(prs: Presentation) -> None:
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if "col G" in r.text or "COL G" in r.text:
                            r.text = r.text.replace("col G", "col C").replace("COL G", "COL C")
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if "col G" in r.text or "COL G" in r.text:
                                    r.text = r.text.replace("col G", "col C").replace("COL G", "COL C")


def _apply_garamond(prs: Presentation) -> int:
    n = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            frames = []
            if shape.has_text_frame:
                frames.append(shape.text_frame)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        frames.append(cell.text_frame)
            for tf in frames:
                for para in tf.paragraphs:
                    for run in para.runs:
                        if not run.text:
                            continue
                        try:
                            color = run.font.color.rgb
                        except AttributeError:
                            color = INK
                        _set_font(
                            run,
                            run.font.size.pt if run.font.size else 10,
                            color,
                            bold=bool(run.font.bold),
                            italic=bool(run.font.italic),
                        )
                        n += 1
    return n


def _rebuild_toc(slide) -> None:
    from fix_pitch_gis_formatting import _read_toc_entries, _add_toc_table, TOC_LEFTS

    entries = _read_toc_entries(slide)
    if not entries:
        return
    split = (len(entries) + 1) // 2
    for shape in list(slide.shapes):
        if shape.has_table and shape.table.rows[0].cells[0].text.strip() == "#":
            _delete_shape(shape)
    _add_toc_table(slide, entries[:split], TOC_LEFTS[0])
    _add_toc_table(slide, entries[split:], TOC_LEFTS[1])


def main() -> None:
    prs = Presentation(str(DECK))
    _rebuild_toc(prs.slides[0])
    _fix_financial_slides(prs)
    _fix_slide17(prs)
    _fix_slide20(prs)
    _fix_slide22(prs)
    _fix_other_footers(prs)
    _global_replacements(prs)
    runs = _apply_garamond(prs)
    prs.save(str(DECK))
    print(f"GIS compliance pass → {DECK} ({runs} runs set to Garamond)")


if __name__ == "__main__":
    main()
