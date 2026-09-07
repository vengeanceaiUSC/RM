"""Generate LULU_Assumptions_Memo.pdf — single-page assumptions summary (FICO-style layout)."""
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import data as D

OUT = os.path.join(os.path.dirname(__file__), "..", "LULU_Assumptions_Memo.pdf")

ROWS = [
    ("REVENUE — CONSOLIDATED", "", "", ""),
    ("FY2026 revenue growth", "−6.1%", "Q2 FY2026 earnings release", '"decline of 5% to 7%"'),
    ("FY2027–30 revenue growth", "+2.3% avg", "StockAnalysis 3Y forecast", '"Revenue Growth Forecast (3Y)" → 2.26%'),
    ("Store fleet (FY25)", "811 stores", "FY2025 10-K", '"Total company-operated stores" → 811'),
    ("China store openings (FY26E)", "+20 net", "Analyst assumption", "Ctrl+F store table; expansion market"),
    ("E-commerce revenue (FY25)", "$4,919M", "FY2025 10-K", '"E-commerce" → 4,918,697'),
    ("Digital conversion (FY26E)", "3.3%", "Analyst assumption", "Calibrated to 10-K channel split"),
    ("Women's / Men's / Acc. mix", "63 / 24 / 13%", "FY2025 10-K", '"women\'s, men\'s, and accessories"'),
    ("MARGINS & FCF", "", "", ""),
    ("FY26 clean EBIT margin", "13.2%", "Q2 FY2026 release", '"18.8%" minus "560 basis points"'),
    ("FY26 tariff refund (one-time)", "$134.5M", "Q2 FY2026 release", '"134.5 million"'),
    ("Terminal EBIT margin (FY30)", "15.5%", "FY2025 10-K anchor", '"19.9%" (unique hit)'),
    ("Capex (% sales, FY26)", "7.0%", "FY2025 10-K guide", '"$725 million and $745 million"'),
    ("NWC (% Δ revenue)", "7.5%", "Analyst assumption", "Single AR+inv+OCA−AP−accrued line"),
    ("WACC & TERMINAL VALUE", "", "", ""),
    ("Risk-free rate", "4.8%", "FRED DGS10", '"2026-09-03" → 4.77'),
    ("Equity risk premium", "6.0%", "Damodaran overlay", '"Implied ERP (FCFE)" 4.23%; model 6.0%'),
    ("Beta (unlevered retail)", "0.95", "Damodaran betas", '"Retail (Special Lines)" → 0.95'),
    ("Beta (relevered for leases)", "~1.06", "WACC tab", "βu × (1 + (1−T) × lease debt / market equity)"),
    ("WACC (base)", "~10.1%", "CAPM build", "Lease-adjusted weights; rf + β×ERP blend"),
    ("Terminal growth (g)", "2.25%", "FRED GDPC1", "Real GDP ≈ 2.1%; model 2.25%"),
    ("Exit EV/EBITDA (selected)", "~6.0x Gordon", "DCF identity", "(UFCF/EBITDA)×(1+g)/(WACC−g)"),
    ("Pubcomps (PitchBook 04-Sep-26)", "4.7–12.7x", "PitchBook Comps Set", "EV/EBITDA = daily EV / TTM EBITDA"),
]


def build():
    doc = SimpleDocTemplate(
        OUT,
        pagesize=letter,
        leftMargin=0.45 * inch,
        rightMargin=0.45 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.35 * inch,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "title", parent=styles["Heading1"], fontSize=13, textColor=colors.HexColor("#1F2A44"),
        spaceAfter=4, leading=15,
    )
    sub = ParagraphStyle(
        "sub", parent=styles["Normal"], fontSize=8.5, textColor=colors.HexColor("#C8102E"),
        spaceAfter=6,
    )
    body = ParagraphStyle("body", parent=styles["Normal"], fontSize=7.2, leading=8.5)
    hdr = ParagraphStyle("hdr", parent=body, fontName="Helvetica-Bold", textColor=colors.white)

    story = [
        Paragraph("GLOBAL INVESTMENT SOCIETY — ASSUMPTIONS MEMO", title),
        Paragraph(
            "lululemon athletica inc. (NASDAQ: LULU)  |  Integrated DCF &amp; 3-Statement Model  |  "
            "FY2025A reported; FY2026E–FY2030E projected",
            sub,
        ),
        Paragraph(
            "<b>Convention:</b> Blue = reported (10-K / earnings). Red = analyst assumption. "
            "Every material input has a Ctrl+F anchor on the linked source page. "
            "Bottom-up revenue drivers on the <i>Revenue Drivers</i> tab; valuation on <i>DCF</i>.",
            body,
        ),
        Spacer(1, 4),
    ]

    data = [["Assumption", "Value", "Source", "Ctrl+F / proof"]]
    section_rows = []
    for i, row in enumerate(ROWS):
        data.append(list(row))
        if row[1] == "":
            section_rows.append(i + 1)

    col_w = [2.05 * inch, 0.85 * inch, 1.55 * inch, 2.55 * inch]
    t = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2A44")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    for sr in section_rows:
        style_cmds += [
            ("BACKGROUND", (0, sr), (-1, sr), colors.HexColor("#D9E1F2")),
            ("FONTNAME", (0, sr), (-1, sr), "Helvetica-Bold"),
            ("SPAN", (0, sr), (-1, sr)),
        ]
    t.setStyle(TableStyle(style_cmds))
    story.append(t)
    story.append(Spacer(1, 5))
    story.append(Paragraph(
        f"<b>Primary filings:</b> {D.filing_url('FY2025')}  |  "
        f"<b>Earnings:</b> {D.SOURCES['earnings_sep2026']}",
        body,
    ))
    doc.build(story)
    print("Saved", os.path.abspath(OUT))


if __name__ == "__main__":
    build()
