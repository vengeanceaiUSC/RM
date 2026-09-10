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
    """Slides 14–16: restore reference table heights and source-map spacing."""
    main_h = Inches(1.98)
    main_top = Inches(3.72)
    src_top = Inches(5.74)
    src_h = Inches(0.95)

    for si in (14, 15, 16):
        slide = prs.slides[si - 1]
        for shape in slide.shapes:
            if shape.has_table:
                h0 = shape.table.rows[0].cells[0].text.strip()
                if h0 == "US$ M":
                    shape.height = int(main_h)
                    shape.top = int(main_top)
                    _compact_table_rows(shape.table, 1.98)
                elif h0 == "Line item":
                    shape.top = int(src_top)
                    shape.height = int(src_h)
                    _compact_table_rows(shape.table, 0.95)
                    _shrink_table_font(shape.table, 7.0)
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
    # Restore MODEL SOURCE MAP tables to reference geometry (never delete)
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        h0 = shape.table.rows[0].cells[0].text.strip()
        if h0 == "Cap stack line item":
            shape.left = int(Inches(0.50))
            shape.top = int(Inches(5.26))
            shape.width = int(Inches(6.05))
            shape.height = int(Inches(0.72))
            _compact_table_rows(shape.table, 0.72)
            _shrink_table_font(shape.table, 7.0)
        elif h0 == "WACC input":
            shape.left = int(Inches(6.78))
            shape.top = int(Inches(5.26))
            shape.width = int(Inches(5.98))
            shape.height = int(Inches(1.08))
            _compact_table_rows(shape.table, 1.08)
            _shrink_table_font(shape.table, 7.0)
    for shape in slide.shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text.strip() == "Component":
            # EV bridge table — reference height
            if shape.top > Inches(2.5):
                shape.height = int(Inches(0.78))
                _compact_table_rows(shape.table, 0.78)
        if shape.has_text_frame and shape.text_frame.text.startswith("No funded bank debt"):
            shape.top = int(Inches(4.22))
            shape.height = int(Inches(0.48))
    for shape in _footer_shapes(slide):
        if not shape.text_frame.text.strip().isdigit():
            _set_footer(shape, "Source: LULU_DCF_Valuation_Model.xlsx, WACC tab (col E) · 10-K cap stack")


def _align_sensitivity_table(table) -> None:
    """Center g-rate headers; right-align price cells for a clean grid."""
    for ci, cell in enumerate(table.rows[0].cells):
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
    for ri in range(1, len(table.rows)):
        for ci, cell in enumerate(table.rows[ri].cells):
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def _fix_slide20(prs: Presentation) -> None:
    """Slide 20 — GIS grid: left col 0.50×6.15, right col 6.85×6.00, sensitivity + callout."""
    slide = prs.slides[19]
    left, lwide = Inches(0.50), Inches(6.15)
    right, rwide = Inches(6.85), Inches(6.00)

    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text
            if "SCENARIOS COL G" in t:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = r.text.replace("SCENARIOS COL G", "SCENARIOS COL C")
            if t.strip().startswith("BASE-CASE ASSUMPTIONS"):
                shape.left = int(left)
                shape.top = int(Inches(1.24))
                shape.width = int(lwide)
                shape.height = int(Inches(0.22))
            if t.strip().startswith("TERMINAL VALUE"):
                shape.left = int(left)
                shape.top = int(Inches(4.32))
                shape.width = int(lwide)
                shape.height = int(Inches(0.25))
            if "SENSITIVITY" in t and "Base-case" in t:
                paras = [p for p in shape.text_frame.paragraphs if "Base-case" not in p.text]
                while len(shape.text_frame.paragraphs) > len(paras):
                    shape.text_frame._txBody.remove(shape.text_frame.paragraphs[-1]._p)
            if "SENSITIVITY" in t and "IMPLIED" in t:
                shape.left = int(left)
                shape.top = int(Inches(5.58))
                shape.width = int(Inches(12.35))
                shape.height = int(Inches(0.20))
            if t.strip().startswith("VALUATION OUTPUT") or (
                "PV of explicit FCF" in t and "Enterprise value" in t
            ):
                shape.left = int(right)
                shape.top = int(Inches(1.48))
                shape.width = int(rwide)
                shape.height = int(Inches(2.75))
            if t.startswith("TV = 73%"):
                shape.left = int(right)
                shape.top = int(Inches(4.58))
                shape.width = int(rwide)
                shape.height = int(Inches(0.92))
                _shrink_text_frame(shape.text_frame, 8.0)
            if t.strip().startswith("Base-case cell"):
                shape.left = int(Inches(8.10))
                shape.top = int(Inches(5.80))
                shape.width = int(Inches(4.75))
                shape.height = int(Inches(1.08))
        if shape.has_table:
            h0 = shape.table.rows[0].cells[0].text.strip()
            if h0 == "Assumption":
                shape.left = int(left)
                shape.top = int(Inches(1.48))
                shape.width = int(lwide)
                shape.height = int(Inches(2.75))
                _compact_table_rows(shape.table, 2.75)
            if h0 == "Method":
                shape.left = int(left)
                shape.top = int(Inches(4.58))
                shape.width = int(lwide)
                shape.height = int(Inches(0.92))
                _compact_table_rows(shape.table, 0.92)
                _shrink_table_font(shape.table, 8.0)
            if h0 == "WACC vs g":
                shape.left = int(left)
                shape.top = int(Inches(5.80))
                shape.width = int(Inches(7.40))
                shape.height = int(Inches(1.08))
                _compact_table_rows(shape.table, 1.08)
                _shrink_table_font(shape.table, 9, skip_header=True)
                _fix_navy_header_row(shape.table, 9)
                _align_sensitivity_table(shape.table)
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
            shape.height = int(Inches(2.15))
            _compact_table_rows(shape.table, 2.15)
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
                        # Preserve white text on navy table headers
                        parent = run._r.getparent()
                        while parent is not None:
                            if parent.tag.endswith("}tc"):
                                break
                            parent = parent.getparent()
                        if parent is not None:
                            try:
                                from pptx.oxml.ns import qn

                                solid = parent.find(f".//{qn('a:solidFill')}")
                                if solid is not None:
                                    srgb = solid.find(qn("a:srgbClr"))
                                    if srgb is not None and srgb.get("val", "").upper() == "1F2A44":
                                        color = WHITE
                            except Exception:
                                pass
                        _set_font(
                            run,
                            run.font.size.pt if run.font.size else 10,
                            color,
                            bold=bool(run.font.bold),
                            italic=bool(run.font.italic),
                        )
                        n += 1
    return n


def _rebuild_toc(slide, entries: list[tuple[str, str]] | None = None) -> None:
    from fix_pitch_gis_formatting import _read_toc_entries, _add_toc_table, TOC_LEFTS

    if entries is None:
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
    toc_entries = None
    from fix_pitch_gis_formatting import _read_toc_entries

    toc_entries = _read_toc_entries(prs.slides[0])
    _fix_financial_slides(prs)
    _fix_slide17(prs)
    _fix_slide20(prs)
    _fix_slide22(prs)
    _fix_other_footers(prs)
    _global_replacements(prs)
    runs = _apply_garamond(prs)
    _rebuild_toc(prs.slides[0], toc_entries)
    prs.save(str(DECK))
    print(f"GIS compliance pass → {DECK} ({runs} runs set to Garamond)")


if __name__ == "__main__":
    main()
