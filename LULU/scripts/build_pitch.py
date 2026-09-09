"""Builds LULU_Investment_Pitch_Deck.pptx using the shared GIS pitch template.

Recommendation: LONG / OVERWEIGHT on lululemon athletica (NASDAQ: LULU).

Figures are sourced from data/pitch_values.json (run pitch_values.py after model changes).
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "GIS"))
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from gis_pitch import (
    PitchDeck,
    add_para,
    textbox,
    rect,
    fit_font_size,
    _set_font,
    _add_bullet,
    NAVY,
    CARD,
    GOLD,
    INK,
    GREY,
    LGREY,
    WHITE,
    GREEN,
)
import data as D

OUT = os.path.join(os.path.dirname(__file__), "..", "LULU_Investment_Pitch_Deck.pptx")
PITCH_VALUES_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "pitch_values.json")
with open(PITCH_VALUES_PATH) as _pf:
    PV = json.load(_pf)
V = PV["valuation"]
IS_BASE = PV["income_statement"]["base"]
BS_BASE = PV["balance_sheet"]["base"]
CF_BASE = PV["cash_flow"]["base"]
PROJ_YEARS = PV["proj_years"]
FF = PV["football_field"]
MR = PV.get("model_refs", {})
WB = PV.get("wacc_build", {})
WR = WB.get("rows", {})
DCF_BASE = PV.get("dcf_base", {})
SOTP = PV.get("sotp", {})
COMPS = PV.get("comps_analysis", {})
PREC = PV.get("precedent", {})


def _d(v):
    return f"${int(round(v))}"


def _m(v):
    return f"{v:,.0f}"


def _up(px, base=100):
    pct = round((px - base) / base * 100)
    sign = "\u2212" if pct < 0 else "+"
    return f"{sign}{abs(pct)}%"


def _pct(v, d=1):
    return f"{v * 100:.{d}f}%"

_BASE_DCF = 133.64
_BASE_UPSIDE = 33.6
_WACC_PCT = 9.0
_MODEL = "LULUMODEL18.xlsx"
_MODEL3 = "LULUMODEL18_3.xlsx"
_MODEL6 = "LULUMODEL18_6.xlsx"
_Q2_SUPP = "https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf"
_LULU_PR = "https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733"
_TIKR = "https://www.tikr.com/blog/lululemon-stock-crashed-17-on-friday-the-guidance-cut-was-the-real-story"
_CONTENT_TOP = 1.28
_FOOTER_Y = 7.08  # GIS footer line (page number)
_LINK_GAP = 0.16  # clearance between body content and sources block


def _links_block_height(n_links: int) -> float:
    """Vertical space reserved for numbered sources above the page footer."""
    return 0.32 + n_links * 0.22


def _links_top(n_links: int) -> float:
    return _FOOTER_Y - _links_block_height(n_links) - 0.08


def _content_bottom(n_links: int) -> float:
    return _links_top(n_links) - _LINK_GAP


def _add_links_box(slide, links, top=None):
    """Numbered source list tucked above the slide footer."""
    n = len(links)
    height = _links_block_height(n)
    if top is None:
        top = _links_top(n)
    tb, tf = textbox(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(height))
    add_para(tf, "Links & Sources", 8, CARD, bold=True, first=True, space_after=1)
    for num, text in links:
        add_para(tf, f"[{num}] {text}", 6.5, INK, bullet=True, space_after=1)


def _add_rich_runs(p, text, size=12, color=INK, default_bold=False):
    """Render **bold** segments inside narrative body copy."""
    parts = re.split(r"\*\*(.+?)\*\*", text)
    for i, part in enumerate(parts):
        if not part:
            continue
        r = p.add_run()
        r.text = part
        _set_font(r, size, color, bold=default_bold or i % 2 == 1)


def _add_labeled_bullet(tf, label, body, size=13, color=INK, first=False, space_after=10):
    """Institutional bullet: **Label**: body text with optional [n] citations."""
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    r_label = p.add_run()
    r_label.text = f"{label}: "
    _set_font(r_label, size, color, bold=True)
    _add_rich_runs(p, body, size=size, color=color)
    _add_bullet(p)
    return p


_CARD_GAP = 0.14      # vertical gutter between bullet panels
_CARD_ACCENT_W = 0.06  # navy rule down the left edge of each panel
_CARD_PAD_L = 0.34     # text inset past the accent rule
_CARD_PAD_R = 0.30


def _narrative_slide(title, descriptor, bullets, links, body_top=_CONTENT_TOP):
    """One panel per bullet, spread evenly over the full content band.

    The copy is fixed and short, so a plain bullet list leaves most of the page
    empty. Giving each bullet its own panel spends that space on structure
    instead of adding words, and matches the timeline slide's card treatment.
    """
    bottom = _content_bottom(len(links))
    s = slide_base(title, descriptor, page=pg(), sources=" ")
    n = len(bullets)
    band = max(0.8, bottom - body_top)
    card_h = (band - (n - 1) * _CARD_GAP) / n
    text_w = 12.35 - _CARD_PAD_L - _CARD_PAD_R
    plain = [b[0] + ": " + b[1] if isinstance(b, tuple) else b for b in bullets]
    size = fit_font_size(plain, text_w, card_h - 0.30)

    for i, bullet in enumerate(bullets):
        top = body_top + i * (card_h + _CARD_GAP)
        rect(s, Inches(0.5), Inches(top), Inches(12.35), Inches(card_h), fill=LGREY)
        rect(s, Inches(0.5), Inches(top), Inches(_CARD_ACCENT_W), Inches(card_h), fill=NAVY)
        _, tf = textbox(
            s,
            Inches(0.5 + _CARD_PAD_L),
            Inches(top),
            Inches(text_w),
            Inches(card_h),
            anchor=MSO_ANCHOR.MIDDLE,
        )
        if isinstance(bullet, tuple):
            _add_labeled_bullet(tf, bullet[0], bullet[1], size=size, first=True)
        else:
            add_para(tf, bullet, size, INK, first=True, space_after=0)
    _add_links_box(s, links)
    return s

deck = PitchDeck(company_header="lululemon athletica (NASDAQ: LULU)")
deck.default_source = "Source: company SEC filings (Form 10-K, CIK 0001397187)"
prs = deck.prs

slide_base = deck.slide_base
body_box = deck.body_box
pg = deck.pg
stmt_table = deck.stmt_table

# =====================================================================
# SLIDES 1-12 (narrative - user-specified copy before Timeline / financials)
# =====================================================================

# Slide 1 - Table of Contents (full deck outline)
_DECK_TOC = [
    (1, "Table of Contents"),
    (2, "Situation Overview and Current Investment Setup"),
    (3, "Market Narrative and Analyst Sentiment Surrounding the Stock"),
    (4, "Investment Thesis Summary and Target Price"),
    (5, "Geographic Segments and Revenue Breakdown"),
    (6, "Business Model Unit Economics and Competitive Moats"),
    (7, "Industry Overview - Trends and Structure"),
    (8, "Industry Overview - Barriers to Entry and Profitability"),
    (9, "Investment Thesis I"),
    (10, "Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor"),
    (11, "Investment Thesis III: Geographic Growth Divergence"),
    (12, "Risk & Mitigants"),
    (13, "Timeline of Recovery"),
    (14, "Financials | Income Statement"),
    (15, "Financials | Balance Sheet"),
    (16, "Financials | Cash Flow"),
    (17, "Financials | Capital Structure & WACC"),
    (18, "Valuation Summary (Football Field)"),
    (19, "Sum of the Parts"),
    (20, "DCF Valuation"),
    (21, "Comps Analysis"),
    (22, "Precedent Transactions"),
]
_toc_top = _CONTENT_TOP
_toc_bottom = _content_bottom(0)
_toc_h = _toc_bottom - _toc_top
_toc_split = (len(_DECK_TOC) + 1) // 2
s = slide_base(
    "Table of Contents",
    "What you will be focusing on during this pitch",
    page=pg(),
    sources=" ",
)
rect(s, Inches(0.5), Inches(_toc_top), Inches(12.35), Inches(_toc_h), fill=LGREY)
_toc_tbl_h = _toc_h - 0.08
_toc_tbl_top = _toc_top + 0.04
for col_i, entries in enumerate((_DECK_TOC[:_toc_split], _DECK_TOC[_toc_split:])):
    stmt_table(
        s,
        [[str(num), title] for num, title in entries],
        ["#", "Slide"],
        col0w=0.42,
        top=_toc_tbl_top,
        height=_toc_tbl_h,
        left=0.5 + col_i * 6.3,
        width=6.02,
        font_size=9,
        header_font_size=9.5,
        text_cols=(1,),
    )

# Slide 2 - Situation Overview
_narrative_slide(
    "Situation Overview and Current Investment Setup",
    "This slide outlines why this investment opportunity exists and the historical financial context driving the current setup",
    [
        "The recent guidance cut triggered a massive cyclical panic pushing Lululemon down to roughly 100 dollars [1]",
        "Despite historically compounding double digit growth the market capitulated over a guidance cut of 5 to 7 percent [2]",
        "Lululemon still retains durable cash flows with clean run rate operating margins of 13.2 percent [3]",
    ],
    [
        (1, f"TIKR LULU Stock Crashed 17%: {_TIKR}"),
        (2, f"Lululemon Q2 FY2026 Guidance Release: {_Q2_SUPP}"),
        (3, f"Provided Valuation Model: {_MODEL}"),
    ],
)

# Slide 3 - Market Narrative
_narrative_slide(
    "Market Narrative and Analyst Sentiment Surrounding the Stock",
    "This slide breaks down current market sentiment and exactly what analysts are saying about the recent guidance cut",
    [
        "Morgan Stanley issued a highly pessimistic forecast after management aggressively cut 2026 revenue guidance to 10.35 billion [1]",
        "Analysts are paralyzed by cyclical fears as the consensus price target was slashed from 176 dollars to 136 dollars [2]",
        "Firms like BTIG maintain neutral ratings due to short-term turbulence but ignore the durable competitive advantage we see [3]",
    ],
    [
        (1, "MarketBeat: Morgan Stanley Issues Pessimistic Forecast for lululemon athletica: https://www.marketbeat.com/instant-alerts/analyst-morgan-stanley-issues-pessimistic-forecast-for-lululemon-athletica-nasdaq-lulu-stock-price-2026-09-04/"),
        (2, "Simply Wall St: lululemon athletica Stock Analysis: https://simplywall.st/stocks/us/consumer-durables/nasdaq-lulu/lululemon-athletica"),
        (3, "GuruFocus: LULU Reiterates by BTIG - Rating Maintained at Neutral: https://www.gurufocus.com/news/9068272/lulu-reiterates-by-btig-rating-maintained-at-neutral"),
    ],
)

# Slide 4 - Investment Thesis Summary
_narrative_slide(
    "Investment Thesis Summary and Target Price",
    "This slide breaks down our $133.64 intrinsic DCF fair value, our $140 twelve-month target price, and the three core pillars supporting our overweight recommendation",
    [
        "Our base-case DCF yields an intrinsic fair value of $133.64 per share (+33.6% upside), supporting our 12-month target price of $140.00 as market multiples modestly re-rate [1]",
        "The first pillar is profitability because adjusting out the tariff refunds reveals LULU still maintains a highly resilient 13.2 percent clean run-rate operating margin [2]",
        "The second pillar is our mathematically sound 9.0 percent WACC which strictly bounds our 2.3 percent long-term revenue growth assumption [3]",
        "Finally, international expansion remains the crucial growth engine, as FY25 China segment revenue grew approximately 29% year-over-year versus flat Americas, while modeled FY26-FY30 five-year averages of 10.2% China comps versus approximately flat Americas comps and 16 versus 6 annual gross store openings offset near-term North American stagnation [4]",
    ],
    [
        (1, f"Provided Valuation Model (Base Case Implied Value): {_MODEL3}"),
        (2, f"Lululemon Q2 FY2026 Earnings Release (13.2% Margin Calc): {_Q2_SUPP}"),
        (3, f"Provided Valuation Model (WACC & Revenue Drivers): {_MODEL3}"),
        (4, f"Lululemon Q2 FY2026 Earnings Release (International Growth): {_Q2_SUPP}"),
    ],
)

# Slide 5 - Geographic segments. The user copy gives no title, so the title bar
# reuses the table-of-contents entry rather than leaving the bar blank.
_narrative_slide(
    _DECK_TOC[4][1],
    "This slide outlines Lululemon's geographic segments and revenue breakdown while detailing why China growth offsets temporary US declines",
    [
        "Lululemon generated 11.1 billion dollars in total revenue with Americas contributing 7.85 billion dollars or 70.68 percent [1]",
        "Americas comparable sales fell 3 percent due to temporary cyclical macro pressure rather than structural brand degradation [2]",
        "China Mainland reached 1.75 billion dollars in segment revenue on +20% FY25 comparable sales growth, proving high-growth international expansion easily offsets US temporary weakness [3]",
    ],
    [
        (1, f"Lululemon FY2025 Form 10-K (Segment Revenue and Percentages): {_Q2_SUPP}"),
        (2, f"Lululemon FY2025 Form 10-K (Americas Comparable Sales): {_Q2_SUPP}"),
        (3, f"Provided Valuation Model (China Mainland Growth & Revenue Driver): {_MODEL}"),
    ],
)

# Slide 6 - Unit economics
_narrative_slide(
    "Business Model Unit Economics and Competitive Moats",
    "This slide analyzes Lululemon's unit economics and competitive moats that protect its long-term market leadership",
    [
        "Lululemon sustains a 56.6 percent gross margin providing a massive competitive moat against retail price competition [1]",
        "E-commerce unit economics remain elite with direct-to-consumer EBIT margins hitting 23.6 percent without retail occupancy drag [2]",
        "Company-operated stores generate 1,426 dollars per square foot and an 18.6 percent EBIT margin establishing high capital efficiency [3]",
    ],
    [
        (1, f"Provided Valuation Model (Gross Margin): {_MODEL6}"),
        (2, f"Provided Valuation Model (E-commerce EBIT Margin Driver): {_MODEL6}"),
        (3, f"Provided Valuation Model (Store EBIT Margin & SPSF): {_MODEL6}"),
    ],
)

# Slide 7 - Industry overview
_narrative_slide(
    "Industry Overview - Trends and Structure",
    "This slide explores the ongoing athleisure industry trends including market fragmentation and the barriers to entry",
    [
        "I argue the global athleisure market remains highly fragmented meaning most players lack a durable competitive advantage [1]",
        "The premium segment maintains high barriers to entry protecting global champions from temporary cyclical noise [2]",
        "Lululemon mathematically proves its moat by generating a 30.25 percent ROE far outpacing the 6.81 percent industry median [3]",
    ],
    [
        (1, "Market.us Media (Athleisure Market Fragmentation): https://media.market.us/athleisure-industry-statistics/"),
        (2, "Fortune Business Insights (Premium Athleisure Trends): https://www.fortunebusinessinsights.com/athleisure-market-110642"),
        (3, "FinanceCharts (Lululemon ROE & Retail Median): https://www.financecharts.com/stocks/LULU/growth/roe"),
    ],
)

# Slide 8 - Industry margins (Slide 7 Part 2)
_narrative_slide(
    "Industry Overview - Barriers to Entry and Profitability",
    "This slide dissects capital efficiency metrics comparing Lululemon's gross margin directly against legacy apparel competitors",
    [
        "Lululemon commands a 56.6 percent gross margin far exceeding traditional athletic apparel peers like Nike and Under Armour [1]",
        "Nike struggles to maintain a 43 percent margin while Under Armour hovers around 46 percent reflecting their wholesale dependence [2]",
        "Lululemon's direct-to-consumer scale prevents this structural margin compression and forms an impenetrable economic moat against industry price wars [3]",
    ],
    [
        (1, f"Provided Valuation Model (Gross Margin Assumptions): {_MODEL}"),
        (2, "Investing.com (Nike & Under Armour Historical Gross Margins): https://www.investing.com/pro/NYSE:NKE/explorer/gp_margin"),
        (3, "ProAnalyst LULU Market Report (Competitor Margins & DTC Moat): https://lulu.proanalyst.ai/business"),
    ],
)

# Slide 9 - Investment Thesis I
_narrative_slide(
    "Investment Thesis I",
    "This slide outlines the core contrarian investment thesis utilizing numerical evidence from the LULUMODEL18.xlsx file",
    [
        "The 2026 price drop exceeding 50% has left Lululemon critically undervalued despite durable cash flows [1]",
        "Its premium Direct-to-Consumer revenue mix protects a massive 54.9% adjusted gross margin, mathematically justifying my intrinsic valuation thesis [2]",
        "Furthermore, even with Americas growth stagnating, sheer cash flow generation creates a robust intrinsic valuation buffer, where a conservative 2.25% terminal growth rate still yields over 30% upside to our $133.64 target price [3]",
    ],
    [
        (1, "https://everythingmoney.com/blog/lululemon-is-collapsing-burry-s-biggest-bet-4334"),
        (2, "https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733"),
        (3, f"{_MODEL}, DCF Terminal Value & Valuation Summary Schedule"),
    ],
)

# Slide 10 - Investment Thesis II
_narrative_slide(
    "Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor",
    "This slide details how partial margin recovery supported by core brand loyalty still drives a highly compelling valuation",
    [
        "Historical pandemic-era peak EBIT margins reached 23.7% in FY24 showing prior peak earnings power [1]",
        "Our valuation assumes a floor built on resilient baseline brand loyalty, proving Lululemon does not need to remain the hottest viral trend 100% of the time to sustain a 13.2% trough margin [2]",
        "A modest partial recovery to just 15.5% EBIT margin by FY30 still yields $133.64 per share [3]",
        "This proves returning to peak COVID profitability is completely unnecessary to unlock substantial market upside [4]",
    ],
    [
        (1, f"{_MODEL} / SEC EDGAR Form 10-K (FY24 Peak Margins)"),
        (2, f"{_MODEL}, Clean Run-Rate EBIT Margin & Brand Royalty Assumptions"),
        (3, f"{_MODEL}, Base Case DCF Valuation Summary"),
        (4, f"{_MODEL}, Discounted Cash Flow Valuation Summary"),
    ],
)

# Slide 11 - Investment Thesis III
_narrative_slide(
    "Investment Thesis III: Geographic Growth Divergence",
    "This slide examines how Lululemon's top-line projections rely disproportionately on Chinese market expansion to conceal domestic North American stagnation.",
    [
        "The revenue build reveals Americas facing near-term contraction with -4.0% comps in FY26 flatlining at a terminal 2.0% growth rate by FY30 [1]",
        "To offset this domestic anchor, the model assumes FY26-FY30 China footprint expansion averaging 16 gross store openings annually versus 6 in Americas, plus 10.2% average same-store comp growth (14% to 7% FY26-FY30) on existing doors [2]",
        "Consequently, if the Chinese consumer softens, this model's core top-line projections will not be optimal enough to meet our target [3]",
    ],
    [
        (1, f"{_MODEL}, Americas FY26-FY30 Comparable Sales Growth Assumptions"),
        (2, f"{_MODEL}, FY26 Store Openings & Mainland China Revenue Build"),
        (3, f"{_MODEL}, Revenue Drivers Schedule"),
    ],
)

# Slide 12 - Risk & Mitigants
_narrative_slide(
    "Risk & Mitigants",
    "This slide evaluates core downside risks and demonstrates how share repurchases compound EPS to drive share price recovery",
    [
        "China deceleration risks a $56.00 bear floor, but $45.64 cumulative cash per share recovers 45% of entry price [1]",
        "Slashed CapEx saves $360M annually by relying on online e-commerce's 23.6% EBIT margin plus $101.5M inventory releases [2]",
        "Deploying $750M annually into buybacks retires 32.9 million shares, compounding EPS to $14.57, more than enough to recover our $133.64 DCF-implied share price [3]",
    ],
    [
        (1, f"{_MODEL}, Bear Case DCF Valuation Summary / https://www.barrons.com/articles/lululemon-stock-earnings-guidance-a7a7c5c0"),
        (2, f"{_MODEL}, CapEx & E-Commerce Channel EBIT Assumptions / {_LULU_PR}"),
        (3, f"{_MODEL}, Share Repurchase & EPS Accretion Schedule / {_LULU_PR}"),
    ],
)

# Slide 13 - Timeline of Recovery (last narrative slide before 3-statement financials on slide 14)
_timeline_links = [
    (1, f"{_MODEL}, Scenarios & Revenue Drivers / {_LULU_PR}"),
    (2, f"{_MODEL}, NOPAT Bridge & Scenarios / {_LULU_PR}"),
    (3, f"{_MODEL}, Scenarios & Unlevered Free Cash Flow Schedule / {_LULU_PR}"),
    (4, f"{_MODEL}, Revenue Drivers & DCF Valuation Summary / https://stockanalysis.com/stocks/lulu/forecast/"),
]
_timeline_top = 1.28
_timeline_rows = 4
_timeline_row_gap = 0.06
_timeline_bottom = _content_bottom(len(_timeline_links))
_timeline_row_h = (
    _timeline_bottom - _timeline_top - (_timeline_rows - 1) * _timeline_row_gap
) / _timeline_rows
_timeline_step = _timeline_row_h + _timeline_row_gap

s = slide_base(
    "Timeline of Recovery",
    " ",
    page=pg(),
    sources=" ",
)
cats = [
    (
        "2026 (Q3 FY2026 Trough)",
        "Q3 FY2026 revenue laps guidance trough while China Double 11 sales confirm holiday store traffic floor stabilization across key markets [1]",
    ),
    (
        "2027 (FY2026 Year-End)",
        "First full-year reset absorbs steep prior declines while $134.5M tariff refunds and targeted SG&A cost actions protect earnings per share [2]",
    ),
    (
        "2027 (FY2027 Margin Inflection)",
        "Operating margins expand to 13.8% (+60 bps) as promotional headwinds anniversary, while $750M annual buybacks compound EPS to $10.09 to boost sentiment [3]",
    ),
    (
        "2028 (FY2027-FY2028 Multiple Re-Rating)",
        "China adds ~17 net stores in FY27 while +12% FY27 and +10% FY28 same-store comps offset Americas -4% FY26 comps, driving re-rating toward $133.64 [4]",
    ),
]
_timeline_label_w = 3.55
_timeline_body_l = 4.2
_timeline_body_w = 8.65
# Match the body type on slides 2-12 so the narrative run reads as one section.
_timeline_body_size = fit_font_size(
    [what for _, what in cats], _timeline_body_w - 0.3, _timeline_row_h - 0.24
)
_timeline_label_size = fit_font_size(
    [when for when, _ in cats], _timeline_label_w - 0.3, _timeline_row_h - 0.24,
    candidates=(13, 12, 11, 10, 9),
)
top = _timeline_top
for when, what in cats:
    b = rect(s, Inches(0.5), Inches(top), Inches(_timeline_label_w), Inches(_timeline_row_h), fill=NAVY)
    bt = b.text_frame
    bt.word_wrap = True
    bt.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt, when, _timeline_label_size, GOLD, bold=True, first=True, space_after=0)
    b2 = rect(s, Inches(_timeline_body_l), Inches(top), Inches(_timeline_body_w), Inches(_timeline_row_h), fill=LGREY)
    bt2 = b2.text_frame
    bt2.word_wrap = True
    bt2.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt2, what, _timeline_body_size, INK, first=True, space_after=0)
    top += _timeline_step
_add_links_box(s, _timeline_links)

def m(v):
    return f"{v:,.0f}"


HY = ["FY2022", "FY2023", "FY2024", "FY2025"]
DCF_MODEL = "LULU_DCF_Valuation_Model.xlsx"
SCEN_TAB = "Scenarios"
SCEN_G = f"{DCF_MODEL}: {SCEN_TAB}, col G (base)"
HIST_10K = "SEC Form 10-K (FY2022-FY2025)"
FIN_FOOTNOTE = (
    "Hist = 10-K. Forecast (26-30) = DCF Scenarios pitch-bridge block, col G (base case only). "
    "Per-row model refs shown under line items."
)


def _scen_rng(key):
    """Excel range for a pitch line item in Scenarios col G."""
    r = MR[key]
    return f"{SCEN_TAB}!G{r['r1']}:G{r['r5']}"


def _row_note(hist_src, fcst_src):
    return f"Hist: {hist_src} · Fcst: {fcst_src}"


def _hist_m(section, key, years=HY):
    """Historical line item in US$M: same units as forecast JSON."""
    return [m(section[key][y] / 1000) for y in years]


def _hist_om(years=HY):
    """Operating margin % = EBIT / revenue (matches forecast row)."""
    return [
        f"{D.IS['operating_income'][y] / D.IS['revenue'][y] * 100:.1f}%"
        for y in years
    ]


def _hist_eps(years=HY):
    return [f"{D.IS['diluted_eps'][y]:.2f}" for y in years]


def _hist_capex_out(years=HY):
    """Capex as parenthetical outflow (10-K: capex stored positive)."""
    return [f"({m(D.CF['capex'][y] / 1000)})" for y in years]


def _hist_fcf(years=HY):
    """FCF = CFO − capex: same definition as forecast (CFO + negative capex)."""
    return [m((D.CF["cfo"][y] - D.CF["capex"][y]) / 1000) for y in years]


def _hist_buybacks_out(years=HY):
    return [f"({m(D.CF['buybacks'][y] / 1000)})" for y in years]


def _fin_headers():
    hdr = ["US$ M", "22", "23", "24", "25"]
    for fy in PROJ_YEARS:
        hdr.append(fy[4:6])
    return hdr


def _hist_range(section, key, y0="FY2022", y1="FY2025"):
    return f"FY22 {m(section[key][y0] / 1000)}M to FY25 {m(section[key][y1] / 1000)}M"


def _hist_line(label, detail):
    """One explainer bullet: metric label + historical FY22 to FY25 detail."""
    return f"{label}: {detail}"


def _base_fcst_line(label, v26, v30, fmt="num"):
    """One explainer bullet: base case FY26 to FY30 (Scenarios col G)."""
    if fmt == "pct":
        return f"{label}: FY26 {v26} to FY30 {v30}"
    if fmt == "eps":
        return f"{label}: FY26 ${v26:.2f} to FY30 ${v30:.2f}"
    if fmt == "outflow":
        return f"{label}: FY26 ({m(abs(v26))}) to FY30 ({m(abs(v30))})"
    return f"{label}: FY26 {m(v26)}M to FY30 {m(v30)}M"


def _base_year_vals(sec, key, by_year=True, fmt="num"):
    vals = []
    for fy in PROJ_YEARS:
        v = sec[fy][key] if by_year else sec[key][fy]
        vals.append(_fmt_cell(v, fmt))
    return vals


def _aligned_explainers(pairs):
    """Build hist/fcst bullet lists with one metric per line (same order)."""
    hist = [_hist_line(label, h) for label, h, _ in pairs]
    fcst = [f for _, _, f in pairs]
    return hist, fcst


def _fin_explainer_boxes(slide, hist_title, hist_bullets, fcst_title, fcst_bullets, box_height=2.45):
    """Side-by-side Historical vs forecast explainer panels above the financial table."""
    bullet_fs = 9 if len(hist_bullets) > 5 else 10
    for x, title, bullets, fill in (
        (0.5, hist_title, hist_bullets, LGREY),
        (6.7, fcst_title, fcst_bullets, NAVY),
    ):
        box = rect(slide, Inches(x), Inches(1.26), Inches(6.05), Inches(box_height), fill=fill)
        btf = box.text_frame
        btf.word_wrap = True
        title_color = CARD if fill == LGREY else GOLD
        body_color = INK if fill == LGREY else WHITE
        add_para(btf, title, 11, title_color, bold=True, first=True, space_after=2)
        for i, bullet in enumerate(bullets):
            add_para(btf, bullet, bullet_fs, body_color, bullet=True, first=(i == 0), space_after=1)


def _fmt_cell(v, fmt="num"):
    if fmt == "pct":
        return v
    if fmt == "eps":
        return f"{v:.2f}"
    if fmt == "neg_paren":
        return f"({abs(v):,})"
    return _m(v)


def _fin_source_table(slide, source_rows, top=5.92, col0w=1.85, height=0.95):
    """Per-line model source: Hist | Fcst (Scenarios G)."""
    headers = ["Line item", "Hist source (22-25)", "Fcst source, Scenarios col G"]
    rows = [[line, hist, fcst] for line, hist, fcst in source_rows]
    stmt_table(
        slide, rows, headers, col0w=col0w, top=top, height=height,
        font_size=7, header_font_size=7.5, bold_rows=(),
        text_cols=(1, 2),
    )


def pitch_financial_slide(
    title,
    subtitle,
    rows,
    row_notes,
    hist_title,
    hist_bullets,
    fcst_title,
    fcst_bullets,
    source_rows,
    bold_rows=(),
    italic_note=None,
    col0w=2.35,
    table_top=3.72,
    table_height=1.98,
    explainer_height=2.32,
):
    """Build one financial slide: historicals + 5yr base-case forecast (Scenarios G)."""
    hdr = _fin_headers()
    footnote = FIN_FOOTNOTE
    if italic_note:
        footnote = f"{italic_note}  {FIN_FOOTNOTE}"
    s = slide_base(title, subtitle, page=pg(), sources=footnote)
    _fin_explainer_boxes(s, hist_title, hist_bullets, fcst_title, fcst_bullets,
                         box_height=explainer_height)
    stmt_table(s, rows, hdr, col0w=col0w, top=table_top, height=table_height,
               bold_rows=bold_rows, font_size=8.5, header_font_size=8)
    _fin_source_table(s, source_rows, top=5.74, height=0.95)
    return s


# =====================================================================
# FINANCIAL STATEMENTS & VALUATION (starts slide 13 - income statement; narrative ends at slide 12)
# =====================================================================
_IS26B, _IS30B = IS_BASE["FY2026E"], IS_BASE["FY2030E"]
_OM22 = D.IS["operating_income"]["FY2022"] / D.IS["revenue"]["FY2022"] * 100
_OM25 = D.IS["operating_income"]["FY2025"] / D.IS["revenue"]["FY2025"] * 100
_OM24 = D.IS["operating_income"]["FY2024"] / D.IS["revenue"]["FY2024"] * 100

_IS_HIST, _IS_FCST = _aligned_explainers([
    (
        "Net revenue",
        f"FY22 ${m(D.IS['revenue']['FY2022']/1000)}M to FY25 ${m(D.IS['revenue']['FY2025']/1000)}M",
        _base_fcst_line("Net revenue", _IS26B["revenue"], _IS30B["revenue"]),
    ),
    (
        "Gross profit",
        _hist_range(D.IS, "gross_profit"),
        _base_fcst_line("Gross profit", _IS26B["gross_profit"], _IS30B["gross_profit"]),
    ),
    (
        "Operating income",
        _hist_range(D.IS, "operating_income"),
        _base_fcst_line("Operating income", _IS26B["operating_income"], _IS30B["operating_income"]),
    ),
    (
        "Operating margin %",
        f"FY22 {_OM22:.1f}% to FY25 {_OM25:.1f}% (peak {_OM24:.1f}% FY24)",
        _base_fcst_line("Operating margin %", _IS26B["operating_margin"], _IS30B["operating_margin"], fmt="pct"),
    ),
    (
        "Net income",
        _hist_range(D.IS, "net_income"),
        _base_fcst_line("Net income", _IS26B["net_income"], _IS30B["net_income"]),
    ),
    (
        "Diluted EPS",
        f"FY22 ${D.IS['diluted_eps']['FY2022']:.2f} to FY25 ${D.IS['diluted_eps']['FY2025']:.2f}",
        _base_fcst_line("Diluted EPS", _IS26B["eps"], _IS30B["eps"], fmt="eps"),
    ),
])

_BS_HIST, _BS_FCST = _aligned_explainers([
    (
        "Cash & equivalents",
        _hist_range(D.BS, "cash"),
        _base_fcst_line("Cash & equivalents", BS_BASE["cash"]["FY2026E"], BS_BASE["cash"]["FY2030E"]),
    ),
    (
        "Inventories",
        _hist_range(D.BS, "inventories"),
        _base_fcst_line("Inventories", BS_BASE["inventories"]["FY2026E"], BS_BASE["inventories"]["FY2030E"]),
    ),
    (
        "Total assets",
        _hist_range(D.BS, "total_assets"),
        _base_fcst_line("Total assets", BS_BASE["total_assets"]["FY2026E"], BS_BASE["total_assets"]["FY2030E"]),
    ),
    (
        "Total liabilities",
        _hist_range(D.BS, "total_liab"),
        _base_fcst_line("Total liabilities", BS_BASE["total_liab"]["FY2026E"], BS_BASE["total_liab"]["FY2030E"]),
    ),
    (
        "Total equity",
        _hist_range(D.BS, "total_equity"),
        _base_fcst_line("Total equity", BS_BASE["total_equity"]["FY2026E"], BS_BASE["total_equity"]["FY2030E"]),
    ),
    (
        "Funded debt",
        "$0 across FY22-FY25 (net-cash)",
        "Funded debt: $0 (no term debt in model)",
    ),
])

_CF_HIST, _CF_FCST = _aligned_explainers([
    (
        "Cash from operations",
        _hist_range(D.CF, "cfo"),
        _base_fcst_line("Cash from operations", CF_BASE["cfo"]["FY2026E"], CF_BASE["cfo"]["FY2030E"]),
    ),
    (
        "D&A (add-back)",
        _hist_range(D.CF, "d_and_a"),
        _base_fcst_line("D&A (add-back)", CF_BASE["dna"]["FY2026E"], CF_BASE["dna"]["FY2030E"]),
    ),
    (
        "Capital expenditures",
        f"FY25 ({m(D.CF['capex']['FY2025']/1000)}) outflow",
        _base_fcst_line("Capital expenditures", CF_BASE["capex"]["FY2026E"], CF_BASE["capex"]["FY2030E"], fmt="outflow"),
    ),
    (
        "Free cash flow",
        f"FY25 {m((D.CF['cfo']['FY2025']-D.CF['capex']['FY2025'])/1000)}M (CFO \u2212 capex)",
        _base_fcst_line("Free cash flow", CF_BASE["fcf"]["FY2026E"], CF_BASE["fcf"]["FY2030E"]),
    ),
    (
        "Share repurchases (CFF)",
        f"FY25 ({m(D.CF['buybacks']['FY2025']/1000)}) outflow",
        _base_fcst_line("Share repurchases (CFF)", CF_BASE["buybacks"]["FY2026E"], CF_BASE["buybacks"]["FY2030E"], fmt="outflow"),
    ),
])

_IS_SOURCES = [
    ("Net revenue", "10-K IS, Net revenue", f"{_scen_rng('revenue')} (Revenue yr 1-5)"),
    ("Gross profit", "10-K IS, Gross profit", f"{_scen_rng('revenue')} × {_scen_rng('gross_margin')} (rev × GM%)"),
    ("Operating income", "10-K IS, Operating income", f"{_scen_rng('ebit')} (EBIT yr 1-5)"),
    ("Operating margin %", "10-K IS, OI ÷ revenue", f"{_scen_rng('ebit')} ÷ {_scen_rng('revenue')}"),
    ("Net income", "10-K IS, Net income", f"{_scen_rng('net_income')} (pitch bridge NI)"),
    ("Diluted EPS", "10-K IS, Diluted EPS", f"NI ÷ buyback-adjusted shares ($750M/yr; {_scen_rng('eps')})"),
]
_IS_NOTES = [
    _row_note("10-K IS, Net revenue", f"{_scen_rng('revenue')}"),
    _row_note("10-K IS, Gross profit", f"{_scen_rng('revenue')} × {_scen_rng('gross_margin')}"),
    _row_note("10-K IS, Operating income", f"{_scen_rng('ebit')}"),
    _row_note("10-K IS, OI ÷ revenue", f"{_scen_rng('ebit')} ÷ {_scen_rng('revenue')}"),
    _row_note("10-K IS, Net income", f"{_scen_rng('net_income')}"),
    _row_note("10-K IS, Diluted EPS", f"NI ÷ diluted sh (750M/yr buybacks); {_scen_rng('eps')}"),
]

_BS_SOURCES = [
    ("Cash & equivalents", "10-K BS, Cash & equivalents", f"{_scen_rng('cash')} (pitch bridge cash roll-forward)"),
    ("Inventories", "10-K BS, Inventories", f"{_scen_rng('inventories')} (NWC schedule)"),
    ("Total assets", "10-K BS, Total assets", f"{_scen_rng('total_assets')} (FY25 TA × rev growth)"),
    ("Total liabilities", "10-K BS, Total liabilities", f"{_scen_rng('total_liab')} (FY25 TL × rev growth)"),
    ("Total equity", "10-K BS, Total equity", f"{_scen_rng('total_equity')} (FY25 TE × rev growth)"),
    ("Funded debt", "10-K BS, Long-term debt ($0)", "Model assumption, $0 (no term debt)"),
]
_BS_NOTES = [
    _row_note("10-K BS, Cash", f"{_scen_rng('cash')}"),
    _row_note("10-K BS, Inventories", f"{_scen_rng('inventories')}"),
    _row_note("10-K BS, Total assets", f"{_scen_rng('total_assets')}"),
    _row_note("10-K BS, Total liabilities", f"{_scen_rng('total_liab')}"),
    _row_note("10-K BS, Total equity", f"{_scen_rng('total_equity')}"),
    _row_note("10-K BS, Debt ($0)", "Model, $0"),
]

_CF_SOURCES = [
    ("Cash from operations", "10-K CF, Operating activities", f"{_scen_rng('cfo')} (pitch bridge CFO)"),
    ("D&A (add-back)", "10-K CF, Depreciation & amortization", f"{_scen_rng('dna')} (D&A yr 1-5)"),
    ("Capital expenditures", "10-K CF, Capital expenditures", f"{_scen_rng('capex')} (Capex yr 1-5)"),
    ("Free cash flow", "10-K CF, CFO − capex", f"{_scen_rng('fcf')} (pitch bridge FCF/CFS)"),
    ("Share repurchases", "10-K CF, Repurchases (financing)", f"{_scen_rng('buybacks')} (G21=$750M/yr fixed)"),
]
_CF_NOTES = [
    _row_note("10-K CF, CFO", f"{_scen_rng('cfo')}"),
    _row_note("10-K CF, D&A", f"{_scen_rng('dna')}"),
    _row_note("10-K CF, Capex", f"{_scen_rng('capex')}"),
    _row_note("10-K CF, CFO − capex", f"{_scen_rng('fcf')}"),
    _row_note("10-K CF, Buybacks", f"{_scen_rng('buybacks')}"),
]

pitch_financial_slide(
    "Financials | Income Statement",
    "Reported history (10-K) vs base-case operating forecast: FY26-30 (US$ M)",
    [
        ["Net revenue"] + _hist_m(D.IS, "revenue") + _base_year_vals(IS_BASE, "revenue"),
        ["Gross profit"] + _hist_m(D.IS, "gross_profit") + _base_year_vals(IS_BASE, "gross_profit"),
        ["Operating income"] + _hist_m(D.IS, "operating_income") + _base_year_vals(IS_BASE, "operating_income"),
        ["Operating margin %"] + _hist_om() + _base_year_vals(IS_BASE, "operating_margin", fmt="pct"),
        ["Net income"] + _hist_m(D.IS, "net_income") + _base_year_vals(IS_BASE, "net_income"),
        ["Diluted EPS ($)"] + _hist_eps() + _base_year_vals(IS_BASE, "eps", fmt="eps"),
    ],
    row_notes=_IS_NOTES,
    hist_title="HISTORICALS (FY2022-FY2025) | SEC 10-K",
    hist_bullets=_IS_HIST,
    fcst_title="FORECAST (FY2026E-FY2030E) | base case (Scenarios col G)",
    fcst_bullets=_IS_FCST,
    source_rows=_IS_SOURCES,
    bold_rows=(0, 2, 5),
    italic_note=(
        f"Forecast: {DCF_MODEL}, Scenarios col G. EPS = NI \u00f7 diluted shares after "
        "$750M/yr repurchases (DCF $134 still uses 111.4M day-one basic shares)."
    ),
)

pitch_financial_slide(
    "Financials | Balance Sheet",
    "Net-cash history vs base-case funded growth: FY26-30 (US$ M)",
    [
        ["Cash & equivalents"] + _hist_m(D.BS, "cash") + _base_year_vals(BS_BASE, "cash", by_year=False),
        ["Inventories"] + _hist_m(D.BS, "inventories") + _base_year_vals(BS_BASE, "inventories", by_year=False),
        ["Total assets"] + _hist_m(D.BS, "total_assets") + _base_year_vals(BS_BASE, "total_assets", by_year=False),
        ["Total liabilities"] + _hist_m(D.BS, "total_liab") + _base_year_vals(BS_BASE, "total_liab", by_year=False),
        ["Total equity"] + _hist_m(D.BS, "total_equity") + _base_year_vals(BS_BASE, "total_equity", by_year=False),
        ["Funded debt"] + ["0"] * 4 + ["0"] * len(PROJ_YEARS),
    ],
    row_notes=_BS_NOTES,
    hist_title="HISTORICALS (FY2022-FY2025) | SEC 10-K",
    hist_bullets=_BS_HIST,
    fcst_title="FORECAST (FY2026E-FY2030E) | base case (Scenarios col G)",
    fcst_bullets=_BS_FCST,
    source_rows=_BS_SOURCES,
    bold_rows=(2, 4, 5),
    italic_note=(
        "BS forecast: cash roll-forward includes $750M/yr buybacks; equity scales with revenue "
        "(buyback-driven equity reduction shown via CF / IS share count, not TE line)."
    ),
)

pitch_financial_slide(
    "Financials | Cash Flow",
    "Operating FCF (DCF input) vs financing: historical 10-K vs base case (US$ M)",
    [
        ["Cash from operations"] + _hist_m(D.CF, "cfo") + _base_year_vals(CF_BASE, "cfo", by_year=False),
        ["D&A (add-back)"] + _hist_m(D.CF, "d_and_a") + _base_year_vals(CF_BASE, "dna", by_year=False),
        ["Capital expenditures"] + _hist_capex_out() + _base_year_vals(CF_BASE, "capex", by_year=False, fmt="neg_paren"),
        ["Free cash flow"] + _hist_fcf() + _base_year_vals(CF_BASE, "fcf", by_year=False),
        ["Share repurchases (CFF)"] + _hist_buybacks_out() + _base_year_vals(CF_BASE, "buybacks", by_year=False, fmt="neg_paren"),
    ],
    row_notes=_CF_NOTES,
    hist_title="HISTORICALS (FY2022-FY2025) | SEC 10-K",
    hist_bullets=_CF_HIST,
    fcst_title="FORECAST (FY2026E-FY2030E) | base case (Scenarios col G)",
    fcst_bullets=_CF_FCST,
    source_rows=_CF_SOURCES,
    bold_rows=(3,),
    italic_note=(
        "Buybacks: $750M/yr fixed (Scenarios col G). Cash roll-forward = prior cash + FCF \u2212 buybacks."
    ),
)

# =====================================================================
# 16. FINANCIALS: CAPITAL STRUCTURE & WACC
# =====================================================================
def _wacc_ref(key):
    r = WR.get(key)
    return f"WACC!E{r}" if r else f"{DCF_MODEL}: WACC tab"


def _pct_wb(v, d=2):
    return f"{v * 100:.{d}f}%"


_CAP_SOURCES = [
    ("Market equity", _wacc_ref("mkt_eq")),
    ("ASC 842 operating leases", f"{_wacc_ref('lease_d')} (FY25 10-K)"),
    ("Funded debt", _wacc_ref("fund_d")),
    ("Total capital", _wacc_ref("debt_tot")),
    ("Cash & equivalents", "10-K BS, Cash"),
    ("Net debt", "DCF EV bridge"),
]
_WACC_SOURCES = [
    ("Risk-free rate", _wacc_ref("rf")),
    ("Equity risk premium", _wacc_ref("erp")),
    ("Tax rate", _wacc_ref("tax")),
    ("Beta (relevered)", _wacc_ref("beta")),
    ("Cost of equity", _wacc_ref("coe")),
    ("Pre-tax Kd", _wacc_ref("kd")),
    ("After-tax Kd", _wacc_ref("kd_at")),
    ("Equity / debt weights", f"{_wacc_ref('we')} / {_wacc_ref('wd')}"),
    ("WACC", _wacc_ref("wacc")),
]
_cap_src_rows = [(line, src) for line, src in _CAP_SOURCES]
_wacc_src_rows = [(line, src) for line, src in _WACC_SOURCES]

s = slide_base(
    "Financials | Capital Structure",
    "Cap stack and WACC build: lease-adjusted; no funded term debt",
    page=pg(),
    sources=f"Cap stack & WACC: {DCF_MODEL}, WACC tab (col E). Cash/leases: FY2025 10-K.",
)

# Section headers (GIS financial-slide style)
tb, tf = textbox(s, Inches(0.5), Inches(1.24), Inches(6.0), Inches(0.22))
add_para(tf, "CAPITAL STACK (US$ M)", 11, CARD, bold=True, first=True, space_after=0)
tb, tf = textbox(s, Inches(6.7), Inches(1.24), Inches(6.15), Inches(0.22))
add_para(tf, "WACC BUILD (CAPM)", 11, CARD, bold=True, first=True, space_after=0)

# --- Cap stack table (left): WACC weights only on equity + lease debt ---
cap_headers = ["Component", "US$ M", "WACC weight"]
cap_rows = [
    ["Market equity (price \u00d7 shares)", f"{WB['mkt_eq_m']:,}", f"{WB['equity_pct']:.1f}%"],
    ["ASC 842 operating lease liabilities", f"{WB['lease_debt_m']:,}", f"{WB['debt_pct']:.1f}%"],
    ["Funded debt (term loans / bonds)", f"{WB['funded_debt_m']:,}", "0.0%"],
    ["Total capital (WACC basis)", f"{WB['total_cap_m']:,}", "100.0%"],
]
stmt_table(
    s, cap_rows, cap_headers, col0w=2.85, top=1.48, height=1.35, left=0.5, width=6.05,
    font_size=9, header_font_size=9, bold_rows=(3,),
)
tb, tf = textbox(s, Inches(0.5), Inches(2.86), Inches(6.05), Inches(0.20))
add_para(tf, "EV BRIDGE (NOT IN WACC WEIGHTS)", 9, CARD, bold=True, first=True, space_after=0)
bridge_rows = [
    ["Cash & equivalents", f"{WB['cash_m']:,}", "Added back in EV, equity bridge"],
    ["Net debt (leases \u2212 cash)", f"{WB['net_debt_m']:,}", "Near net-cash; leases in WACC above"],
]
stmt_table(
    s, bridge_rows, ["Component", "US$ M", "Note"],
    col0w=2.85, top=3.08, height=0.78, left=0.5, width=6.05,
    font_size=8.5, header_font_size=8.5,
)
tb, tf = textbox(s, Inches(0.5), Inches(4.12), Inches(6.05), Inches(0.5))
add_para(
    tf,
    "No funded bank debt; ASC 842 store leases are the only debt equivalent (~14% WACC weight). "
    f"Cash ${WB['cash_m']:,}M exceeds lease debt, net-cash on a funded-debt basis.",
    8.5, INK, italic=True, first=True, space_after=0,
)

# --- WACC build (right): template LGREY panel + compact 2-col table ---
rect(s, Inches(6.7), Inches(1.50), Inches(6.15), Inches(3.53), fill=LGREY)
wacc_compact = [
    ["Risk-free rate (10-yr UST)", _pct_wb(WB["rf"])],
    ["Equity risk premium", _pct_wb(WB["erp"])],
    ["Tax rate", _pct_wb(WB["tax"])],
    ["Beta (relevered; lease adj.)", f"{WB['beta']:.2f}"],
    ["Cost of equity (CAPM)", _pct_wb(WB["coe"])],
    ["Pre-tax Kd (lease-equivalent)", _pct_wb(WB["kd"])],
    ["After-tax Kd", _pct_wb(WB["kd_at"])],
    ["Equity / debt weight", f"{_pct_wb(WB['we'], 1)} / {_pct_wb(WB['wd'], 1)}"],
]
stmt_table(
    s, wacc_compact,
    ["Input", "Value"],
    col0w=3.35, top=1.56, height=2.22, left=6.78, width=5.98,
    font_size=9, header_font_size=9, bold_rows=(4,),
)
tb, tf = textbox(s, Inches(6.85), Inches(3.88), Inches(5.85), Inches(0.24))
add_para(
    tf,
    f"\u03b2: Yahoo \u03b2L {WB['beta_obs']:.2f} to unlevered {WB['beta_unlev']:.2f} "
    f"@ D/E {WB['de_unlev']:.2f} to relever {WB['beta']:.2f} @ lease D/E {WB['de_relev']:.2f}",
    8, INK, italic=True, first=True, space_after=0,
)
box = rect(s, Inches(6.7), Inches(4.14), Inches(6.15), Inches(0.84), fill=NAVY)
btf = box.text_frame
btf.word_wrap = True
btf.vertical_anchor = MSO_ANCHOR.MIDDLE
add_para(btf, f"WACC = {_pct_wb(WB['wacc'])}", 17, WHITE, bold=True, first=True, space_after=2)
add_para(
    btf,
    f"Base-case discount rate: Scenarios col G \u00b7 implied price {_d(V['base_dcf'])}",
    9, WHITE, space_after=0,
)

# --- Model source map (footer, matches other financial slides) ---
# Split across two columns under their own halves of the slide: all 15 rows in
# one full-width table stack past the bottom of the page.
tb, tf = textbox(s, Inches(0.5), Inches(5.06), Inches(12.35), Inches(0.2))
add_para(tf, "MODEL SOURCE MAP", 9, CARD, bold=True, first=True, space_after=0)
_SRC_MAP_TOP = 5.26
_SRC_MAP_HDR = "DCF model source (WACC tab, col E)"
stmt_table(
    s, [[a, b] for a, b in _cap_src_rows],
    ["Cap stack line item", _SRC_MAP_HDR],
    col0w=2.4, top=_SRC_MAP_TOP, height=0.72, left=0.5, width=6.05,
    font_size=7, header_font_size=7.5,
    text_cols=(1,),
)
stmt_table(
    s, [[a, b] for a, b in _wacc_src_rows],
    ["WACC input", _SRC_MAP_HDR],
    col0w=2.4, top=_SRC_MAP_TOP, height=1.08, left=6.78, width=5.98,
    font_size=7, header_font_size=7.5,
    bold_rows=(len(_wacc_src_rows) - 1,),
    text_cols=(1,),
)

# =====================================================================
# 17. VALUATION SUMMARY (FOOTBALL FIELD)
# =====================================================================
s = slide_base(
    "Valuation Summary",
    "Football field: implied share-price ranges by methodology (base case DCF)",
    page=pg(),
    sources=f"Source: {DCF_MODEL}, Comps / football field tab; geographic SOTP on FY30E base revenue mix",
)

_base_px = int(V["base_dcf"])
_sotp_lo = SOTP.get("implied_px_lo", _base_px)
_sotp_hi = SOTP.get("implied_px_hi", _base_px)
_gordon_exit = SOTP.get("gordon_exit_multiple", DCF_BASE.get("exit_multiple", 7.36))
ff_methods = [
    ("P / E (10-18x FY2026E)", round(FF["P / E"]["low"]), round(FF["P / E"]["high"])),
    ("EV / EBITDA (5.0-8.0x FY30E)", round(FF["EV / EBITDA"]["low"]), round(FF["EV / EBITDA"]["high"])),
    ("Geographic SOTP (FY30E; Gordon anchor)", _sotp_lo, _sotp_hi),
    ("Unlevered DCF (base / Gordon g)", _base_px, _base_px),
    ("52-week range", 100, 226),
]
chart_l, chart_r = 3.35, 12.5
vmin, vmax = 50, 240


def _xpos(v):
    return chart_l + (v - vmin) / (vmax - vmin) * (chart_r - chart_l)


# axis
for tick in (50, 100, 150, 200):
    tx = _xpos(tick)
    rect(s, Inches(tx), Inches(1.38), Pt(1), Inches(4.35), fill=LGREY)
    tb, tf = textbox(s, Inches(tx - 0.25), Inches(5.78), Inches(0.55), Inches(0.22))
    add_para(tf, f"${tick}", 8, GREY, align=PP_ALIGN.CENTER, first=True, space_after=0)

top = 1.55
for name, lo, hi in ff_methods:
    tb, tf = textbox(s, Inches(0.5), Inches(top - 0.02), Inches(2.75), Inches(0.48), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, name, 11, NAVY, bold=True, first=True, space_after=0)
    if lo == hi:
        mx = _xpos(lo)
        rect(s, Inches(mx - 0.015), Inches(top + 0.04), Pt(3), Inches(0.34), fill=NAVY)
        # Label above marker - avoids collision with $140 target line at nearby x-positions
        tb2, tf2 = textbox(s, Inches(mx - 0.42), Inches(top - 0.22), Inches(0.84), Inches(0.24))
        add_para(tf2, f"${lo}", 9.5, NAVY, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
    else:
        bar_w = max(_xpos(hi) - _xpos(lo), 0.15)
        rect(s, Inches(_xpos(lo)), Inches(top), Inches(bar_w), Inches(0.42), fill=GOLD)
        # Low label outside bar (left) - keeps $139 clear of target line at $140
        tb_lo, tf_lo = textbox(
            s, Inches(_xpos(lo) - 0.5), Inches(top), Inches(0.46), Inches(0.42), anchor=MSO_ANCHOR.MIDDLE
        )
        add_para(tf_lo, f"${lo}", 9.5, NAVY, bold=True, align=PP_ALIGN.RIGHT, first=True, space_after=0)
        tb2, tf2 = textbox(s, Inches(_xpos(hi) + 0.02), Inches(top), Inches(0.85), Inches(0.42), anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf2, f"${hi}", 9.5, NAVY, bold=True, first=True, space_after=0)
    top += 0.72

cp_x = _xpos(100)
pt_x = _xpos(140)
rect(s, Inches(cp_x), Inches(1.45), Pt(2), Inches(4.25), fill=INK)
rect(s, Inches(pt_x), Inches(1.45), Pt(2), Inches(4.25), fill=CARD)
tb, tf = textbox(s, Inches(cp_x - 0.6), Inches(1.22), Inches(1.2), Inches(0.24))
add_para(tf, "Current $100", 8.5, INK, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
tb, tf = textbox(s, Inches(pt_x - 0.7), Inches(1.22), Inches(1.2), Inches(0.24))
add_para(tf, "Target $140", 8.5, CARD, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)

tb, tf = textbox(s, Inches(0.5), Inches(6.12), Inches(12.35), Inches(0.88))
add_para(
    tf,
    f"12-month target $140 sits above base-case DCF {_d(_base_px)} and inside the comps/SOTP ranges: a partial re-rating, not a return to peak multiples.",
    12.5, NAVY, bold=True, first=True, space_after=4,
)
add_para(
    tf,
    f"DCF row = Scenarios col G (Gordon growth). Bear {_d(V['bear'])} / bull {_d(V['bull'])} bracket scenario range from the model.",
    11, INK, italic=True, space_after=0,
)

# =====================================================================
# 18. SUM OF THE PARTS (GEOGRAPHIC)
# =====================================================================
s = slide_base(
    "Sum of the Parts",
    f"Geographic segments on FY30E base revenue: EV/EBITDA spreads vs Gordon-implied exit ({_gordon_exit:.1f}x)",
    page=pg(),
    sources="Segment revenue: FY2025 10-K geography; FY30E scaled to base-case consolidated revenue (Scenarios col G)",
)
seg_rows = []
for seg in SOTP.get("segments", []):
    seg_rows.append([
        seg["segment"],
        f"{seg['fy30_rev_m']:,}",
        f"{seg['ebitda_margin_pct']:.1f}%",
        f"{seg['fy30_ebitda_m']:,}",
        f"{seg['ev_ebitda_lo']:.1f}x-{seg['ev_ebitda_hi']:.1f}x",
        f"{seg['ev_lo_m']:,}-{seg['ev_hi_m']:,}",
    ])
seg_rows.append([
    "Total segment EV",
    f"{SOTP.get('fy30_rev_m', 0):,}",
    "",
    f"{SOTP.get('fy30_ebitda_m', 0):,}",
    f"{SOTP.get('consolidated_multiple_lo', _gordon_exit):.1f}x-{SOTP.get('consolidated_multiple_hi', _gordon_exit):.1f}x",
    f"{SOTP.get('total_ev_lo_m', 0):,}-{SOTP.get('total_ev_hi_m', 0):,}",
])
seg_rows.append([
    "Corporate / HQ (no separate carve-out)",
    "", "", "", "n/a", "$0",
])
seg_rows.append([
    "+ Cash / \u2212 lease debt (EV bridge)",
    "", "", "", "n/a",
    f"+{SOTP.get('cash_m', 0):,} / \u2212{SOTP.get('debt_m', 0):,}",
])
seg_rows.append([
    "Implied equity value / share",
    "", "", "", "n/a",
    f"${SOTP.get('implied_px_lo', 0)}-${SOTP.get('implied_px_hi', 0)}",
])
stmt_table(
    s, seg_rows,
    ["Segment", "FY30E rev", "EBITDA %", "FY30E EBITDA", "EV/EBITDA", "Segment EV ($M)"],
    col0w=2.15, top=1.2, height=3.05, font_size=9, header_font_size=9,
    bold_rows=(3, len(seg_rows) - 1),
)
tb, tf = textbox(s, Inches(0.5), Inches(4.38), Inches(7.8), Inches(1.85))
add_para(tf, "Methodology", 12, CARD, bold=True, first=True, space_after=3)
for t in [
    "Single-brand retailer: geography is the cleanest SOTP cut (Americas / China / RoW per 10-K)",
    "FY30E segment revenue = FY25 geo mix \u00d7 base-case consolidated FY30 revenue (Scenarios col G)",
    f"Segment EV/EBITDA spreads anchor to Gordon-implied exit {_gordon_exit:.1f}x (selected DCF TV identity: not the 5-8x comps football-field band)",
    "Americas: Gordon \u2212 1.0x to \u2212 0.25x (mature); China: +0.5x to +2.0x (growth); RoW: \u22120.25x to +0.75x",
    f"Consolidated base-case DCF {_d(_base_px)} uses Gordon growth (g={DCF_BASE.get('terminal_g', 0.0225)*100:.2f}%); SOTP is terminal-year EBITDA triangulation only",
]:
    add_para(tf, t, 10.5, INK, bullet=True, space_after=3)

box = rect(s, Inches(8.5), Inches(4.38), Inches(4.35), Inches(1.55), fill=LGREY)
btf = box.text_frame
btf.word_wrap = True
add_para(btf, "SOTP vs DCF", 12, CARD, bold=True, first=True, space_after=4)
add_para(btf, f"SOTP range: ${_sotp_lo}-${_sotp_hi}", 13, NAVY, bold=True, space_after=3)
add_para(btf, f"Base-case DCF: {_d(_base_px)}", 13, NAVY, bold=True, space_after=3)
add_para(btf, "Overlap is expected: SOTP applies FY30 EBITDA multiples; DCF discounts explicit FCF + Gordon growth TV.", 9.5, INK, space_after=0)

# =====================================================================
# 19. DCF VALUATION (BASE CASE)
# =====================================================================
_rev_g1 = DCF_BASE.get("rev_growth_fy26", -0.061)
_rev_gt = DCF_BASE.get("rev_growth_fy27_30", 0.023)
_exit_m = DCF_BASE.get("exit_multiple", 7.36)
_tg = DCF_BASE.get("terminal_g", 0.0225)
_tv_pct = DCF_BASE.get("tv_pct_ev", 0.735)

s = slide_base(
    "DCF Valuation",
    f"Base-case unlevered DCF (Scenarios col G): Gordon growth terminal value; {_exit_m:.1f}x is implied exit identity",
    page=pg(),
    sources=f"Source: {DCF_MODEL}, Scenarios col G + DCF tab",
)

# --- Assumptions (left) ---
tb, tf = textbox(s, Inches(0.5), Inches(1.24), Inches(6.2), Inches(0.22))
add_para(tf, "BASE-CASE ASSUMPTIONS (SCENARIOS COL G)", 11, CARD, bold=True, first=True, space_after=0)
arows = [
    ["Assumption", "Value"],
    ["FY2026 revenue growth", f"{_rev_g1 * 100:.1f}%"],
    ["FY2027-30 revenue growth (avg)", f"{_rev_gt * 100:+.1f}%"],
    ["Clean EBIT margin (FY26 run-rate)", f"{DCF_BASE.get('ebit_margin_clean', 0.132) * 100:.1f}%"],
    ["Terminal EBIT margin (FY2030E)", f"{DCF_BASE.get('ebit_margin_terminal', 0.155) * 100:.1f}%"],
    ["FY26 tariff refunds (one-time)", f"${DCF_BASE.get('tariff_refund_k', 134500) / 1000:.1f}M"],
    ["Cash tax rate", f"{DCF_BASE.get('tax', 0.30) * 100:.0f}%"],
    ["Capex % of revenue", f"{DCF_BASE.get('capex_pct', 0.055) * 100:.1f}%"],
    ["WACC", _pct(V["wacc"])],
    ["Terminal growth (g)", f"{_tg * 100:.2f}%"],
    ["Gordon-implied exit EV/EBITDA (identity)", f"{_exit_m:.1f}x"],
]
stmt_table(s, arows[1:], arows[0], col0w=3.5, top=1.48, height=2.75, left=0.5, width=6.15,
           font_size=9, header_font_size=9, bold_rows=())

# --- Terminal value approaches (left bottom) ---
tb, tf = textbox(s, Inches(0.5), Inches(4.32), Inches(6.15), Inches(0.25))
add_para(tf, "TERMINAL VALUE | GORDON GROWTH (SELECTED)", 10, CARD, bold=True, first=True, space_after=0)
tv_rows = [
    ["Method", "Formula / input", "FY30 TV ($M)"],
    [
        "Gordon growth (selected)",
        f"FCF\u2085 \u00d7 (1+g) / (WACC\u2212g); g={_tg*100:.2f}%",
        f"{DCF_BASE.get('gordon_tv_m', 0):,}",
    ],
    [
        "Exit multiple (identity check)",
        f"FY30 EBITDA \u00d7 {_exit_m:.1f}x = Gordon TV \u00f7 EBITDA (same number)",
        f"{DCF_BASE.get('exit_tv_m', 0):,}",
    ],
]
stmt_table(s, tv_rows[1:], tv_rows[0], col0w=1.85, top=4.58, height=0.92, left=0.5, width=6.15,
           font_size=8.5, header_font_size=8.5, bold_rows=())

# --- Output (right) ---
box = rect(s, Inches(6.85), Inches(1.35), Inches(6.0), Inches(2.55), fill=LGREY)
btf = box.text_frame
btf.word_wrap = True
add_para(btf, "VALUATION OUTPUT (US$ M)", 11, CARD, bold=True, first=True, space_after=5)
for t2, v in [
    ("PV of explicit FCF (FY26-FY30)", _m(round(V["pv_fcf_m"]))),
    ("PV of terminal value", _m(round(V["pv_tv_m"]))),
    ("Enterprise value", _m(round(V["ev_m"]))),
    ("Plus: cash", _m(round(V["cash_m"]))),
    ("Less: lease debt (ASC 842)", f"({_m(round(WB.get('total_debt_m', 0)))})"),
    ("Equity value", _m(round(V["equity_m"]))),
    ("\u00f7 Diluted shares (M)", f"{V['shares_m']:.1f}"),
]:
    p = btf.add_paragraph()
    p.space_after = Pt(3)
    r = p.add_run()
    r.text = t2
    _set_font(r, 10.5, INK, bold=("Enterprise" in t2 or "Equity" in t2))
    r2 = p.add_run()
    r2.text = f"  {v}"
    _set_font(r2, 10.5, NAVY, bold=True)
p = btf.add_paragraph()
p.space_before = Pt(4)
r = p.add_run()
r.text = f"Implied value:  {_d(_base_px)} / share"
_set_font(r, 15, GREEN, bold=True)

box2 = rect(s, Inches(6.85), Inches(4.05), Inches(6.0), Inches(1.25), fill=NAVY)
b2 = box2.text_frame
b2.word_wrap = True
add_para(
    b2,
    f"TV = {_tv_pct * 100:.0f}% of EV  \u00b7  {_exit_m:.1f}x = Gordon TV \u00f7 FY30 EBITDA ${_m(DCF_BASE.get('fy30_ebitda_m', 0))}M (not a peer pick)",
    10.5, WHITE, bold=True, first=True, space_after=3,
)
add_para(
    b2,
    f"Selected TV is Gordon growth (g). The {_exit_m:.1f}x exit multiple is the identity check (TV \u00f7 FY30 EBITDA), not a peer average.",
    9, WHITE, space_after=0,
)

# --- Sensitivity ---
tb, tf = textbox(s, Inches(0.5), Inches(5.58), Inches(12.35), Inches(0.20))
add_para(tf, "SENSITIVITY | IMPLIED SHARE PRICE (WACC vs TERMINAL g)", 10, CARD, bold=True, first=True, space_after=0)
_g_cols = ["1.5%", "2.0%", "2.25%", "2.5%", "3.0%"]
sens = [["WACC \\ g"] + _g_cols]
_base_wacc_idx = min(
    range(len(PV["sensitivity"])),
    key=lambda i: abs(float(PV["sensitivity"][i]["wacc"].rstrip("%")) - V["wacc"] * 100),
)
for row in PV["sensitivity"]:
    sens.append([row["wacc"]] + [_d(p) for p in row["prices"]])
stmt_table(
    s, sens[1:], sens[0], col0w=1.0, top=5.80, height=1.08, left=0.5, width=7.4,
    font_size=9, header_font_size=9,
    bold_rows=(),
)
# highlight base cell via note
tb, tf = textbox(s, Inches(8.1), Inches(5.80), Inches(4.75), Inches(1.08))
add_para(tf, "Base-case cell", 9, CARD, bold=True, first=True, space_after=2)
add_para(
    tf,
    f"WACC {_pct(V['wacc'])} \u00d7 g {_tg*100:.2f}% gives {_d(_base_px)}. Grid brackets \u00b1100bps WACC and 1.5-3.0% g.",
    9, INK, space_after=0,
)

# =====================================================================
# 20. COMPS ANALYSIS
# =====================================================================
_ca = COMPS
_core = _ca.get("core_stats", {})
_lulu_t = _ca.get("lulu_trading", {})
s = slide_base(
    "Comps Analysis",
    "Implied LULU valuation from PitchBook core athletic / apparel set (04-Sep-2026)",
    page=pg(),
    sources=_ca.get("source", "PitchBook Comps Set 04-Sep-2026"),
)
# --- Peer multiples (left) ---
tb, tf = textbox(s, Inches(0.5), Inches(1.24), Inches(7.2), Inches(0.20))
add_para(tf, "COMPARABLE COMPANIES | TRADING MULTIPLES", 10, CARD, bold=True, first=True, space_after=0)
peer_hdr = ["Company", "EV/Rev", "EV/EBITDA", "EV/EBIT", "P/E"]
peer_rows = []
for p in _ca.get("peers", []):
    if not p.get("core") or p["name"] == "lululemon (LULU)":
        continue
    short = p["name"].split("(")[0].strip()
    peer_rows.append([
        short,
        f"{p['ev_rev']:.2f}x" if p.get("ev_rev") else "n/a",
        f"{p['ev_ebitda']:.1f}x" if p.get("ev_ebitda") else "n/a",
        f"{p['ev_ebit']:.1f}x" if p.get("ev_ebit") else "n/a",
        f"{p['pe_fwd']:.1f}x" if p.get("pe_fwd") else "n/a",
    ])
med = _core
if med:
    peer_rows.append([
        "Core median (ex-LULU)",
        f"{med.get('ev_rev', {}).get('median', 0):.2f}x",
        f"{med.get('ev_ebitda', {}).get('median', 0):.1f}x",
        f"{med.get('ev_ebit', {}).get('median', 0):.1f}x",
        f"{med.get('pe_fwd', {}).get('median', 0):.1f}x",
    ])
peer_rows.append([
    "LULU @ $100 (current)",
    f"{_lulu_t.get('ev_rev', 0):.2f}x",
    f"{_lulu_t.get('ev_ebitda', 0):.1f}x" if _lulu_t.get("ev_ebitda") else "n/a",
    f"{_lulu_t.get('ev_ebit', 0):.1f}x",
    f"{_lulu_t.get('pe_fwd', 0):.1f}x",
])
stmt_table(
    s, peer_rows, peer_hdr, col0w=2.0, top=1.46, height=2.40, left=0.5, width=7.25,
    font_size=8.5, header_font_size=8.5, bold_rows=(len(peer_rows) - 1,),
)

# --- Implied valuation (right) ---
tb, tf = textbox(s, Inches(7.95), Inches(1.24), Inches(5.0), Inches(0.20))
add_para(tf, "IMPLIED LULU VALUATION (CORE PEER MEDIAN)", 10, CARD, bold=True, first=True, space_after=0)
imp_hdr = ["Metric", "Peer med.", "LULU base", "Implied px"]
imp_rows = []
for row in _ca.get("implied", []):
    base = row["lulu_base"]
    base_s = f"${base:.2f}" if row["metric"] == "P / E" else f"${int(base):,}M"
    imp_rows.append([
        row["metric"],
        f"{row['peer_median']:.1f}x" if row["metric"] != "EV / Revenue" else f"{row['peer_median']:.2f}x",
        base_s,
        f"${row['implied_px_low']}-${row['implied_px_high']} (med ${row['implied_px_median']})",
    ])
stmt_table(
    s, imp_rows, imp_hdr, col0w=1.2, top=1.46, height=2.40, left=7.95, width=5.05,
    font_size=7.5, header_font_size=8, bold_rows=(),
)
tb, tf = textbox(s, Inches(7.95), Inches(3.92), Inches(5.05), Inches(0.48))
add_para(
    tf,
    "Implied EV = peer median \u00d7 LULU base (cols 2-3); P/E = median \u00d7 EPS (no EV). "
    "EV + cash \u2212 lease debt \u00f7 shares = equity/sh.",
    7.5, INK, italic=True, first=True, space_after=0,
)

tb, tf = textbox(s, Inches(0.5), Inches(4.48), Inches(12.35), Inches(2.20))
_ebitda_med = _core.get("ev_ebitda", {}).get("median", 0)
_rev_med = _core.get("ev_rev", {}).get("median", 0)
_pe_med = _core.get("pe_fwd", {}).get("median", 0)
_lulu_ebitda = _lulu_t.get("ev_ebitda", 0)
_exit_m = DCF_BASE.get("exit_multiple", 7.36)
_tgt_px = 140
for title, body in [
    (
        "Peer benchmark methodology",
        "Core peer set (NKE, ADS, DECK, CROX, LEVI, KTB) evaluates relative valuation using PitchBook data "
        "(04-Sep-2026). Implied equity = peer median \u00d7 LULU base, EV, then + cash \u2212 ASC 842 lease debt \u00f7 shares.",
    ),
    (
        "Multiples confirm severe undervaluation",
        f"LULU trades at {_lulu_t.get('ev_rev', 0):.2f}x EV/Revenue, {_lulu_ebitda:.1f}x EV/EBITDA (PitchBook TTM), and "
        f"{_lulu_t.get('pe_fwd', 0):.1f}x forward P/E vs core peer medians of {_rev_med:.2f}x EV/Revenue, "
        f"{_ebitda_med:.1f}x EV/EBITDA, and {_pe_med:.1f}x P/E.",
    ),
    (
        "DCF terminal valuation discount",
        f"Base-case DCF uses Gordon growth g = {DCF_BASE.get('terminal_g', 0.0225)*100:.2f}%, implying {_exit_m:.1f}x FY30 "
        f"EV/EBITDA exit multiple: well below the {_ebitda_med:.1f}x peer median (right table uses TTM EBITDA).",
    ),
    (
        "Conservative target price re-rating",
        f"Our {_d(_base_px)} base-case DCF and ${_tgt_px} 12-month target (+{_tgt_px - int(_base_px)} vs DCF, ~{round((_tgt_px/int(_base_px)-1)*100)}% "
        f"above model fair value) reflect undervaluation with only a modest re-rating: not Nike-level multiples.",
    ),
]:
    add_para(tf, title, 11, CARD, bold=True, first=(title == "Peer benchmark methodology"), space_after=2)
    add_para(tf, body, 10, INK, space_after=6)

# =====================================================================
# 21. PRECEDENT TRANSACTIONS
# =====================================================================
s = slide_base(
    "Precedent Transactions",
    "Private-market context for athleisure M&A: not used to imply a control premium on LULU",
    page=pg(),
    sources="Source: Wolverine IR/SEC (Sweaty Betty); Bloomberg (Gymshark); BusinessWire/Reuters (Vuori); Reuters/Forbes (Alo)",
)
_prec_deals = PREC.get("deals", [])
_lulu_es = PREC.get("lulu_ev_sales", 0)
prec_hdr = ["Target", "Year", "Type", "EV ($M)", "EV/Rev", "EV/EBITDA", "Status"]
prec_rows = []
for d in _prec_deals:
    status = (d.get("status", "") or "")
    if len(status) > 48:
        status = status[:45] + "..."
    prec_rows.append([
        d.get("target", ""),
        d.get("year", ""),
        d.get("deal_type", ""),
        f"{d.get('ev_usd_m', 0):,}" if d.get("ev_usd_m") else "n/a",
        f"{d.get('ev_sales', 0):.1f}x" if d.get("ev_sales") else "n/a",
        f"{d.get('ev_ebitda', 0):.1f}x" if d.get("ev_ebitda") else "n/a",
        status,
    ])
prec_rows.append([
    "LULU (public @ $100)",
    "n/a",
    "Public equity",
    f"{PREC.get('lulu_ev_m', _lulu_t.get('ev_m', 0)):,}",
    f"{_lulu_es:.2f}x",
    f"{_lulu_t.get('ev_ebitda', 0):.1f}x",
    "Trading comps only (not a premium benchmark)",
])
stmt_table(
    s, prec_rows, prec_hdr, col0w=1.55, top=1.28, height=2.15, left=0.5, width=12.35,
    font_size=8, header_font_size=8, bold_rows=(len(prec_rows) - 1,),
)

tb, tf = textbox(s, Inches(0.5), Inches(3.55), Inches(7.85), Inches(2.55))
for title, body in [
    (
        "Private precedent framework",
        "Non-Alo private transactions establish a contextual floor of what strategics and PE pay for premium athleisure. "
        "They do not imply a control premium on LULU's public stock.",
    ),
    (
        "Closed M&A benchmark",
        "Wolverine paid $410M (~16x EBITDA, ~1.6x sales) for Sweaty Betty (2021), a concrete operating-company "
        "precedent for women's DTC activewear.",
    ),
    (
        "Growth capital rounds",
        "Gymshark ($1.3B implied EV, 2020 minority stake) and Vuori ($4B post-money, 2021) show institutional demand "
        "and realistic baselines for premium DTC brands.",
    ),
    (
        "Excluded unclosed process",
        "Alo's ~$10B ask (~5.0x EV/Sales) is LULU's closest yoga peer, but unclosed private asks cannot dictate "
        "actual public equity valuation.",
    ),
]:
    add_para(tf, title, 10.5, CARD, bold=True, first=(title == "Private precedent framework"), space_after=2)
    add_para(tf, body, 9.5, INK, space_after=5)

box = rect(s, Inches(8.5), Inches(3.55), Inches(4.35), Inches(1.35), fill=LGREY)
btf = box.text_frame
btf.word_wrap = True
add_para(btf, "Valuation anchors", 11, CARD, bold=True, first=True, space_after=4)
add_para(btf, f"DCF base: {_d(_base_px)}  |  Target: $140", 12, NAVY, bold=True, space_after=3)
add_para(
    btf,
    "Private precedents inform strategic context only; public comps on the prior slide.",
    9, INK, italic=True, space_after=0,
)

def _strip_terminal_periods(presentation):
    """Drop the full stop that closes a paragraph, deck-wide.

    Sentence punctuation inside a paragraph is left alone; only the very last
    character goes, so "9.0% WACC. Grid brackets..." keeps its interior period.
    Ellipses are skipped.
    """
    frames = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                frames.append(shape.text_frame)
            if shape.has_table:
                frames += [c.text_frame for r in shape.table.rows for c in r.cells]
    for frame in frames:
        for para in frame.paragraphs:
            runs = [r for r in para.runs if r.text]
            if not runs:
                continue
            text = "".join(r.text for r in runs).rstrip()
            if not text.endswith(".") or text.endswith(".."):
                continue
            last = runs[-1]
            last.text = last.text.rstrip().removesuffix(".")


_strip_terminal_periods(prs)
n = deck.save(OUT)
print("Saved", os.path.abspath(OUT), "with", n, "slides")
try:
    from gis_pitch import export_pdf
    pdf = export_pdf(OUT)
    print("Saved", pdf)
except Exception as e:
    print("PDF export skipped:", e)
