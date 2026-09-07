"""Generate LULU_Assumptions_Memo.pdf — full assumptions guide for every model tab.

Pulls justification, source label, URL, and Ctrl+F proof from data.py (JUST,
ASSUMPTION_SRC, SOURCE_HINT) so the PDF stays in sync with the workbooks.
"""
import os
import re
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import data as D

OUT = os.path.join(os.path.dirname(__file__), "..", "LULU_Assumptions_Memo.pdf")

# Tab order in the PDF (prefix match on JUST keys)
TAB_SECTIONS = [
    ("WACC", "DCF workbook — WACC tab", "wacc_"),
    ("Scenarios", "DCF workbook — Scenarios tab", "sc_"),
    ("Revenue Drivers", "DCF workbook — Revenue Drivers tab", "drv_"),
    ("DCF", "DCF workbook — DCF tab (valuation & sensitivity)", ("dcf_", "sens_")),
    ("Comps", "DCF workbook — Comps / Football Field tab", "comps_"),
    ("3-Statement", "LULU_3_Statement_Model — Assumptions tab", "3s_"),
]

LABEL_OVERRIDES = {
    "wacc_rf": "Risk-free rate (10-yr UST)",
    "wacc_erp": "Equity risk premium",
    "wacc_beta_obs": "Observed beta (5Y) — not used",
    "wacc_beta_ind": "Unlevered retail beta (Damodaran)",
    "wacc_beta": "Beta used (relevered for lease debt)",
    "wacc_kd": "Pre-tax cost of debt (lease-equivalent)",
    "wacc_tax": "Tax rate",
    "wacc_mkt_px": "Share price ($)",
    "wacc_mkt_eq": "Market value of equity",
    "wacc_mkt_shares": "Shares outstanding (000)",
    "wacc_lease_d": "Operating lease liabilities (debt equiv.)",
    "wacc_fund_d": "Funded debt (term loans / bonds)",
    "wacc_we": "Equity weight",
    "wacc_wd": "Debt weight (leases + funded)",
    "sc_g1": "FY2026 revenue growth",
    "sc_gterm": "FY2027–30 revenue growth",
    "sc_m1": "FY2026 clean EBIT margin",
    "sc_tariff": "FY2026 tariff refund (one-time)",
    "sc_mterm": "Terminal (FY2030E) EBIT margin",
    "sc_wacc": "WACC (base links to WACC tab)",
    "sc_g": "Terminal growth (g)",
    "sc_tax": "Tax rate",
    "sc_da_pct": "D&A (% of revenue)",
    "sc_capex_pct": "Capex (% of revenue)",
    "sc_capex_sales": "FY2026 sales (capex denominator)",
    "sc_ar": "AR (% of sales) — inside NWC",
    "sc_nwc_pct": "NWC (% of Δ revenue)",
    "dcf_exitm": "Selected exit EV/EBITDA (Gordon implied)",
    "dcf_debt_bridge": "EV bridge: debt & operating leases",
    "sens_axes": "Sensitivity grid axes",
    "sens_wacc": "Sensitivity — WACC axis",
    "sens_g": "Sensitivity — terminal g axis",
    "drv_f_geo_scale": "Geo revenue scale formula",
    "drv_f_var_pct": "Variance (%) formula",
}


def _label(key):
    if key in LABEL_OVERRIDES:
        return LABEL_OVERRIDES[key]
    s = key
    for prefix in ("wacc_", "sc_", "dcf_", "sens_", "comps_", "drv_f_", "drv_", "3s_"):
        if s.startswith(prefix):
            s = s[len(prefix):]
            break
    return s.replace("_", " ").title()


def _keys_for_section(prefixes):
    if isinstance(prefixes, str):
        prefixes = (prefixes,)
    keys = []
    for key in D.JUST:
        if any(key.startswith(p) for p in prefixes):
            keys.append(key)
    return keys


def _source_cell(key):
    src = D.ASSUMPTION_SRC.get(key)
    if not src:
        return "—"
    label, url = src
    label = escape(label or "")
    if url:
        return f'<a href="{escape(url)}" color="#0563C1">{label}</a>'
    return label


def _para(text, style):
    return Paragraph(escape(str(text)).replace("\n", "<br/>"), style)


def _link_para(html, style):
    return Paragraph(html, style)


def build():
    doc = SimpleDocTemplate(
        OUT,
        pagesize=letter,
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.35 * inch,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "title", parent=styles["Heading1"], fontSize=14, textColor=colors.HexColor("#1F2A44"),
        spaceAfter=6, leading=16,
    )
    sub = ParagraphStyle(
        "sub", parent=styles["Normal"], fontSize=9, textColor=colors.HexColor("#C8102E"),
        spaceAfter=8, leading=11,
    )
    body = ParagraphStyle("body", parent=styles["Normal"], fontSize=8, leading=10)
    tab_hdr = ParagraphStyle(
        "tabhdr", parent=styles["Heading2"], fontSize=11, textColor=colors.HexColor("#1F2A44"),
        spaceBefore=6, spaceAfter=4, leading=13,
    )
    cell = ParagraphStyle("cell", parent=body, fontSize=7.2, leading=8.5)
    cell_bold = ParagraphStyle("cell_bold", parent=cell, fontName="Helvetica-Bold")

    story = [
        Paragraph("GLOBAL INVESTMENT SOCIETY — FULL ASSUMPTIONS GUIDE", title),
        Paragraph(
            "lululemon athletica inc. (NASDAQ: LULU)  |  Every red assumption on every tab  |  "
            "FY2025A reported; FY2026E–FY2030E projected",
            sub,
        ),
        Paragraph(
            "<b>Convention:</b> Blue = reported (10-K / earnings). Red = analyst assumption. "
            "Black = formulas. This PDF mirrors columns <i>Justification</i>, <i>Source</i>, and "
            "<i>Ctrl+F</i> on each workbook tab. Click source links to open the filing or data page.",
            body,
        ),
        Spacer(1, 6),
        Paragraph("<b>DCF workbook tabs:</b> WACC → Scenarios → Revenue Drivers → DCF → Comps", body),
        Paragraph("<b>3-statement workbook tabs:</b> Assumptions → Income Statement → Balance Sheet → Cash Flow", body),
        Spacer(1, 8),
    ]

    col_w = [1.35 * inch, 2.35 * inch, 1.45 * inch, 2.35 * inch]
    total_rows = 0

    for i, (tab_name, tab_desc, prefixes) in enumerate(TAB_SECTIONS):
        keys = _keys_for_section(prefixes)
        if not keys:
            continue
        if i > 0:
            story.append(PageBreak())
        story.append(Paragraph(f"{tab_name}", tab_hdr))
        story.append(Paragraph(tab_desc, body))
        story.append(Spacer(1, 4))

        data = [[
            Paragraph("<b>Assumption</b>", cell_bold),
            Paragraph("<b>Justification (~20 words)</b>", cell_bold),
            Paragraph("<b>Source (click)</b>", cell_bold),
            Paragraph("<b>Ctrl+F / proof</b>", cell_bold),
        ]]
        for key in keys:
            just = D.JUST.get(key, "")
            hint = D.SOURCE_HINT.get(key, "")
            data.append([
                Paragraph(f"<b>{escape(_label(key))}</b>", cell_bold),
                _para(just, cell),
                _link_para(_source_cell(key), cell),
                _para(hint or "—", cell),
            ])
            total_rows += 1

        t = Table(data, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2A44")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)

    story.append(PageBreak())
    story.append(Paragraph("Primary source links", tab_hdr))
    story.append(Spacer(1, 4))
    for label, url in [
        ("FY2025 Form 10-K", D.filing_url("FY2025")),
        ("Q2 FY2026 earnings release", D.SOURCES["earnings_sep2026"]),
        ("FRED DGS10 (risk-free)", D.SOURCES["fred_dgs10"]),
        ("FRED GDPC1 (GDP)", D.SOURCES["fred_gdpc1"]),
        ("Damodaran ERP", D.SOURCES["damodaran_erp"]),
        ("Damodaran betas", D.SOURCES["damodaran_betas"]),
        ("StockAnalysis LULU stats", D.SOURCES["lulu_stats"]),
        ("NASDAQ LULU quote", D.SOURCES["nasdaq_quote"]),
    ]:
        story.append(Paragraph(
            f'• <a href="{escape(url)}" color="#0563C1">{escape(label)}</a>', body,
        ))

    doc.build(story)
    print(f"Saved {os.path.abspath(OUT)} ({total_rows} assumptions across {len(TAB_SECTIONS)} sections)")


if __name__ == "__main__":
    build()
