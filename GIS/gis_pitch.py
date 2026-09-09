"""GIS Investment Research pitch deck template engine.

Reusable slide chrome, theme colors, and layout helpers for python-pptx decks.
Follows GIS formatting rules:
  - Garamond throughout
  - 0.3" navy header rectangle, title 15pt right-aligned
  - One-sentence descriptor per slide (no trailing period)
  - USC cardinal / gold / navy palette
  - Running header: Investment Research Division + company line
  - Footer source line + page number
"""
from __future__ import annotations

import os

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- Theme colors ----
NAVY = RGBColor(0x1F, 0x2A, 0x44)
CARD = RGBColor(0x99, 0x00, 0x00)   # USC cardinal
GOLD = RGBColor(0xFF, 0xC7, 0x2C)   # USC gold
INK = RGBColor(0x26, 0x26, 0x26)
GREY = RGBColor(0x8C, 0x8C, 0x8C)
LGREY = RGBColor(0xF0, 0xF1, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1E, 0x7A, 0x3C)
LIGHT_BLUE = RGBColor(0xCF, 0xD6, 0xE4)
FONT = "Garamond"

# Standard 22-slide GIS structure (matches IR selection template)
STANDARD_TOC = [
    "1.  Investment thesis summary",
    "2.  Situation overview",
    "3.  Market narrative",
    "4.  Company overview",
    "5.  Business model & unit economics",
    "6.  Industry overview",
    "7.  Thesis I: Priced for terminal decline",
    "8.  Thesis II: International growth engine",
    "9.  Thesis III: Elite economics & capital return",
    "10. Risks & mitigants",
    "11. Catalyst timeline",
    "12. Financials: income statement",
    "13. Financials: balance sheet",
    "14. Financials: cash flow",
    "15. Capital structure & WACC",
    "16. Valuation summary (football field)",
    "17. DCF valuation",
    "18. Comparable companies",
    "19. Appendix: bull / bear scenarios",
]

DEFAULT_SOURCE = "Source: company SEC filings (Form 10-K, CIK 0000000000)"


def _set_font(run, size, color=INK, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    return tb, tf


def _no_bullet(p):
    pPr = p._pPr
    if pPr is None:
        pPr = p._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buAutoNum"):
        e = pPr.find(qn(tag))
        if e is not None:
            pPr.remove(e)
    pPr.append(pPr.makeelement(qn("a:buNone"), {}))


def _add_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.set("indent", str(Pt(-12)))
    pPr.set("marL", str(Pt(14)))
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": FONT}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "\u2022"}))


def add_para(
    tf,
    text,
    size=14,
    color=INK,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    bullet=False,
    level=0,
    space_after=4,
    first=False,
):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.level = level
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    r = p.add_run()
    r.text = text
    _set_font(r, size, color, bold, italic)
    if bullet:
        _add_bullet(p)
    else:
        _no_bullet(p)
    return p


def rect(slide, l, t, w, h, fill=NAVY, line=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(0.75)
    sp.shadow.inherit = False
    return sp


class PitchDeck:
    """GIS pitch deck builder with shared chrome and page counter."""

    def __init__(self, company_header: str = "Company Name (NASDAQ: TICK)"):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank = self.prs.slide_layouts[6]
        self.sw = self.prs.slide_width
        self.sh = self.prs.slide_height
        self.company_header = company_header
        self._page = 0
        self.default_source = DEFAULT_SOURCE

    def pg(self) -> int:
        self._page += 1
        return self._page

    def slide_base(self, title, descriptor, sources=None, page=None):
        s = self.prs.slides.add_slide(self.blank)
        # Left header stops at 5.90" so it never sits under the navy title bar.
        tb, tf = textbox(s, Inches(0.45), Inches(0.16), Inches(5.40), Inches(0.44))
        add_para(tf, "Investment Research Division", 11, CARD, bold=True, first=True, space_after=0)
        add_para(tf, self.company_header, 12.5, NAVY, bold=True, space_after=0)

        title_clean = (title or "").strip()
        if title_clean:
            title_h = 0.40 if len(title_clean) > 40 else 0.30
            title_size = 12 if len(title_clean) > 52 else (13 if len(title_clean) > 38 else 15)
            hr = rect(s, Inches(6.05), Inches(0.16), Inches(6.85), Inches(title_h), fill=NAVY)
            htf = hr.text_frame
            htf.word_wrap = True
            htf.vertical_anchor = MSO_ANCHOR.MIDDLE
            htf.margin_top = Pt(1)
            htf.margin_bottom = Pt(1)
            htf.margin_left = Pt(8)
            htf.margin_right = Pt(8)
            hp = htf.paragraphs[0]
            hp.alignment = PP_ALIGN.RIGHT
            hr_run = hp.add_run()
            hr_run.text = title_clean
            _set_font(hr_run, title_size, WHITE, bold=True)

        desc = (descriptor or "").strip()
        desc_size = 10.5 if len(desc) > 95 else 12.5
        tb2, tf2 = textbox(s, Inches(0.45), Inches(0.64), Inches(12.45), Inches(0.48))
        if desc:
            add_para(tf2, desc, desc_size, GREY, italic=True, first=True, space_after=0)

        rect(s, Inches(0.45), Inches(1.18), Inches(12.45), Pt(1.6), fill=GOLD)

        # Footer source stops short of the page number.
        ftb, ftf = textbox(s, Inches(0.45), Inches(7.12), Inches(11.65), Inches(0.30))
        src = sources or self.default_source
        add_para(ftf, src, 8, GREY, first=True, space_after=0)
        if page is not None:
            ptb, ptf = textbox(s, Inches(12.25), Inches(7.12), Inches(0.70), Inches(0.30))
            add_para(ptf, str(page), 9, GREY, align=PP_ALIGN.RIGHT, first=True, space_after=0)
        return s

    def body_box(self, slide, l=Inches(0.5), t=Inches(1.28), w=Inches(12.35), h=Inches(5.75)):
        return textbox(slide, l, t, w, h)

    def title_slide(
        self,
        company_name: str,
        ticker_line: str,
        recommendation: str,
        current_price: str,
        target_price: str,
        upside_pct: str,
        descriptor: str,
        date_line: str = "September 2026        Prepared for the GIS IR selection process",
    ):
        s = self.prs.slides.add_slide(self.blank)
        rect(s, 0, 0, self.sw, self.sh, fill=NAVY)
        rect(s, 0, Inches(2.55), self.sw, Inches(0.06), fill=GOLD)
        rect(s, 0, Inches(4.35), self.sw, Inches(0.06), fill=CARD)

        tb, tf = textbox(s, Inches(0.9), Inches(0.7), Inches(11.5), Inches(1.4))
        add_para(tf, "GLOBAL INVESTMENT SOCIETY", 20, GOLD, bold=True, first=True, space_after=2)
        add_para(tf, "Investment Research Division", 15, WHITE, space_after=0)

        tb, tf = textbox(s, Inches(0.9), Inches(2.7), Inches(11.5), Inches(1.6))
        add_para(tf, company_name, 40, WHITE, bold=True, first=True, space_after=0)
        add_para(tf, ticker_line, 20, GOLD, bold=True, space_after=0)

        tb, tf = textbox(s, Inches(0.9), Inches(4.6), Inches(11.6), Inches(2.2))
        add_para(tf, f"Recommendation:  {recommendation}", 22, GOLD, bold=True, first=True, space_after=8)
        add_para(
            tf,
            f"Current price:  {current_price}        Price target:  {target_price}  ({upside_pct})",
            17,
            WHITE,
            space_after=6,
        )
        add_para(tf, descriptor, 13.5, LIGHT_BLUE, italic=True, space_after=10)
        add_para(tf, date_line, 11, GREY, space_after=0)
        return s

    def toc_slide(self, items: list[str] | None = None, descriptor: str = "What this pitch will cover"):
        toc = items or STANDARD_TOC
        s = self.slide_base("Table of Contents", descriptor, page=self.pg())
        tb, tf = textbox(s, Inches(0.6), Inches(1.2), Inches(6.0), Inches(5.8))
        mid = (len(toc) + 1) // 2
        for i, item in enumerate(toc[:mid]):
            add_para(tf, item, 14.5, NAVY, bold=True, first=(i == 0), space_after=8)
        tb2, tf2 = textbox(s, Inches(6.9), Inches(1.2), Inches(6.0), Inches(5.8))
        for j, item in enumerate(toc[mid:]):
            add_para(tf2, item, 14.5, NAVY, bold=True, first=(j == 0), space_after=8)
        return s

    def placeholder_slide(
        self,
        title: str,
        descriptor: str = "[One-sentence descriptor :  no trailing period]",
        sections: list[tuple[str, list[str]]] | None = None,
        sources: str | None = None,
    ):
        """Generic content slide with placeholder bullets for blank templates."""
        s = self.slide_base(title, descriptor, sources=sources, page=self.pg())
        tb, tf = self.body_box(s)
        if sections:
            for i, (heading, bullets) in enumerate(sections):
                add_para(tf, heading, 14.5, CARD, bold=True, first=(i == 0), space_after=5)
                for b in bullets:
                    add_para(tf, b, 13, INK, bullet=True, space_after=5)
        else:
            add_para(tf, "[Section header]", 14.5, CARD, bold=True, first=True, space_after=5)
            for _ in range(3):
                add_para(tf, "[Bullet point]", 13, INK, bullet=True, space_after=5)
        return s

    def stmt_table(
        self,
        slide,
        rows,
        headers,
        col0w=3.6,
        top=1.35,
        height=5.4,
        left=0.5,
        width=12.35,
        red_rows=(),
        bold_rows=(),
        font_size=10.5,
        header_font_size=None,
        row_notes=(),
        note_font_size=None,
    ):
        ncol = len(headers)
        hdr_fs = header_font_size if header_font_size is not None else (8 if ncol > 12 else 11)
        left_in = Inches(left)
        width_in = Inches(width)
        tbl_shape = slide.shapes.add_table(len(rows) + 1, ncol, left_in, Inches(top), width_in, Inches(height))
        table = tbl_shape.table
        table.columns[0].width = Inches(col0w)
        restw = (width - col0w) / (ncol - 1)
        for c in range(1, ncol):
            table.columns[c].width = Inches(restw)
        for c, htxt in enumerate(headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = htxt
            _set_font(r, hdr_fs, WHITE, bold=True)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        note_fs = note_font_size if note_font_size is not None else max(6.5, font_size - 2)
        for ri, row in enumerate(rows, start=1):
            is_red = ri - 1 in red_rows
            is_bold = ri - 1 in bold_rows
            note = row_notes[ri - 1] if ri - 1 < len(row_notes) else None
            for c, val in enumerate(row):
                cell = table.cell(ri, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if ri % 2 else LGREY
                tf = cell.text_frame
                tf.clear()
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
                r = p.add_run()
                r.text = val
                color = CARD if (is_red and c > 0) else (NAVY if is_bold else INK)
                _set_font(r, font_size, color, bold=is_bold)
                if c == 0 and note:
                    np = tf.add_paragraph()
                    np.alignment = PP_ALIGN.LEFT
                    nr = np.add_run()
                    nr.text = note
                    _set_font(nr, note_fs, GREY, italic=True)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_top = Pt(1)
                cell.margin_bottom = Pt(1)
        return table

    def save(self, path: str) -> int:
        self.prs.save(path)
        return len(self.prs.slides._sldIdLst)


def export_pdf(pptx_path: str, out_dir: str | None = None) -> str:
    """Export PPTX to PDF via LibreOffice (soffice). Returns PDF path."""
    import subprocess

    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir or os.path.dirname(pptx_path))
    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    pdf_path = os.path.join(out_dir, pdf_name)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", pptx_path, "--outdir", out_dir],
        check=True,
        capture_output=True,
    )
    if not os.path.isfile(pdf_path):
        raise RuntimeError(f"PDF export failed: {pdf_path} not created")
    return pdf_path
