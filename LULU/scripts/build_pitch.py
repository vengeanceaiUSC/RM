"""Builds LULU_Investment_Pitch_Deck.pptx using the shared GIS pitch template.

Recommendation: LONG / OVERWEIGHT on lululemon athletica (NASDAQ: LULU).

Figures are sourced from data/pitch_values.json (run pitch_values.py after model changes).
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "GIS"))
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from gis_pitch import (
    PitchDeck,
    add_para,
    textbox,
    rect,
    _set_font,
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
IS_BULL = PV["income_statement"]["bull"]
BS_BASE = PV["balance_sheet"]["base"]
BS_BULL = PV["balance_sheet"]["bull"]
CF_BASE = PV["cash_flow"]["base"]
CF_BULL = PV["cash_flow"]["bull"]
PROJ_YEARS = PV["proj_years"]
FF = PV["football_field"]


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

deck = PitchDeck(company_header="lululemon athletica (NASDAQ: LULU)")
deck.default_source = "Source: company SEC filings (Form 10-K, CIK 0001397187)"
prs = deck.prs

slide_base = deck.slide_base
body_box = deck.body_box
pg = deck.pg
stmt_table = deck.stmt_table

# =====================================================================
# 1. TITLE
# =====================================================================
deck.title_slide(
    company_name="lululemon athletica inc.",
    ticker_line="NASDAQ: LULU",
    recommendation="OVERWEIGHT / LONG",
    current_price="$100.00",
    target_price="$133.64",
    upside_pct="+33.6% upside",
    descriptor="A net-cash, high-margin brand priced for terminal decline \u2014 we see a cyclical trough, not a broken business",
)

deck._page = 1  # so presentation slide 2 has page number 2

LULU_TOC = [
    "1.  Situation overview & current setup",
    "2.  Market narrative & analyst sentiment",
    "3.  Investment thesis summary & target price",
    "4.  Geographic segments & revenue breakdown",
    "5.  Business model & unit economics",
    "6.  Industry overview \u2014 trends & structure",
    "7.  Industry overview \u2014 barriers & profitability",
    "8.  Thesis I \u2014 Contrarian setup & valuation buffer",
    "9.  Thesis II \u2014 Partial margin recovery",
    "10. Thesis III \u2014 Geographic divergence",
    "11. Risks & mitigants",
    "12. Timeline of recovery",
    "13. Financials \u2014 income statement",
    "14. Financials \u2014 balance sheet",
    "15. Financials \u2014 cash flow",
    "16. Capital structure & WACC",
    "17. Valuation summary (football field)",
    "18. DCF valuation",
    "19. Comparable companies",
    "20. Appendix \u2014 bull / bear scenarios",
]

# Helper to create standard card slide for slides 2-12
def build_prompt_slide(
    slide_num,
    title,
    descriptor,
    tag,
    cards,
    bullets,
    sources_text,
    extra_heading=None,
):
    s = slide_base(title, descriptor, page=pg())

    # 3 Top highlight cards
    cx = Inches(0.5)
    cw = Inches(3.95)
    gap = Inches(0.25)
    top_y = Inches(1.12)
    card_h = Inches(1.10)
    for c_title, c_val, c_sub in cards:
        b = rect(s, cx, top_y, cw, card_h, fill=LGREY)
        btf = b.text_frame
        btf.word_wrap = True
        btf.margin_top = Pt(4)
        btf.margin_bottom = Pt(4)
        btf.margin_left = Pt(8)
        btf.margin_right = Pt(8)
        add_para(btf, c_title, 10.5, CARD, bold=True, first=True, space_after=1)
        add_para(btf, c_val, 19, NAVY, bold=True, space_after=1)
        add_para(btf, c_sub, 9.5, GREY, space_after=0)
        cx += cw + gap

    # Middle narrative box
    body_y = Inches(2.32)
    body_h = Inches(3.45)
    tb, tf = textbox(s, Inches(0.5), body_y, Inches(12.35), body_h)
    
    # Tag / section header
    add_para(tf, tag, 13, CARD, bold=True, first=True, space_after=4)
    if extra_heading:
        add_para(tf, extra_heading, 12, INK, bold=True, space_after=4)

    for i, bullet in enumerate(bullets):
        add_para(tf, bullet, 12.2, INK, bullet=True, space_after=6)

    # Bottom sources box
    src_y = Inches(5.85)
    src_h = Inches(1.22)
    sb = rect(s, Inches(0.5), src_y, Inches(12.35), src_h, fill=LGREY)
    stf = sb.text_frame
    stf.word_wrap = True
    stf.margin_top = Pt(4)
    stf.margin_bottom = Pt(4)
    stf.margin_left = Pt(8)
    stf.margin_right = Pt(8)

    if isinstance(sources_text, list):
        for j, s_line in enumerate(sources_text):
            is_hdr = s_line.startswith("Links & Sources")
            f_size = 9.5 if is_hdr else 8.5
            f_color = CARD if is_hdr else GREY
            add_para(stf, s_line, f_size, f_color, bold=is_hdr, first=(j == 0), space_after=1)
    else:
        add_para(stf, sources_text, 8.5, GREY, first=True, space_after=0)

    return s


# =====================================================================
# 2. SITUATION OVERVIEW (PROMPT SLIDE 2)
# =====================================================================
build_prompt_slide(
    2,
    "Situation Overview and Current Investment Setup",
    "This slide outlines why this investment opportunity exists and the historical financial context driving the current setup",
    "Slide 2: Situation Overview and Current Investment Setup",
    [
        ("CURRENT SHARE PRICE", "$100.00", "Down ~55% from peak; 17% single-day crash"),
        ("FY2026 GUIDANCE RESET", "−5% to −7%", "First revenue decline in corporate history"),
        ("CLEAN RUN-RATE MARGIN", "13.2%", "Durable operating cash flow through trough"),
    ],
    [
        "The recent guidance cut triggered a massive cyclical panic pushing Lululemon down to roughly 100 dollars [1]",
        "Despite historically compounding double digit growth the market capitulated over a guidance cut of 5 to 7 percent [2]",
        "Lululemon still retains durable cash flows with clean run rate operating margins of 13.2 percent [3]",
    ],
    [
        "Links & Sources:",
        "[1] TIKR LULU Stock Crashed 17%: https://www.tikr.com/blog/lululemon-stock-crashed-17-on-friday-the-guidance-cut-was-the-real-story",
        "[2] Lululemon Q2 FY2026 Guidance Release: https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
        "[3] Provided Valuation Model: LULUMODEL18.xlsx",
    ],
)

# =====================================================================
# 3. MARKET NARRATIVE (PROMPT SLIDE 3)
# =====================================================================
build_prompt_slide(
    3,
    "Market Narrative and Analyst Sentiment Surrounding the Stock",
    "This slide breaks down current market sentiment and exactly what analysts are saying about the recent guidance cut",
    "Slide 3: Market Narrative and Analyst Sentiment Surrounding the Stock",
    [
        ("REVENUE GUIDANCE CUT", "$10.35B", "Morgan Stanley pessimistic outlook"),
        ("CONSENSUS TARGET CUT", "$176 → $136", "Cyclical fears paralyze the Street"),
        ("RATING DISPERSION", "Neutral Bias", "BTIG ignores durable competitive moats"),
    ],
    [
        "Morgan Stanley issued a highly pessimistic forecast after management aggressively cut 2026 revenue guidance to 10.35 billion [1]",
        "Analysts are paralyzed by cyclical fears as the consensus price target was slashed from 176 dollars to 136 dollars [2]",
        "Firms like BTIG maintain neutral ratings due to short-term turbulence but ignore the durable competitive advantage we see [3]",
    ],
    "Links & Sources: [1] MarketBeat: Morgan Stanley Issues Pessimistic Forecast for lululemon athletica: https://www.marketbeat.com/instant-alerts/analyst-morgan-stanley-issues-pessimistic-forecast-for-lululemon-athletica-nasdaq-lulu-stock-price-2026-09-04/ [2] Simply Wall St: lululemon athletica Stock Analysis: https://simplywall.st/stocks/us/consumer-durables/nasdaq-lulu/lululemon-athletica [3] GuruFocus: LULU Reiterates by BTIG - Rating Maintained at Neutral: https://www.gurufocus.com/news/9068272/lulu-reiterates-by-btig-rating-maintained-at-neutral",
)

# =====================================================================
# 4. INVESTMENT THESIS SUMMARY (PROMPT SLIDE 4)
# =====================================================================
build_prompt_slide(
    4,
    "Investment Thesis Summary and Target Price",
    "This slide breaks down our actual 133 dollar base case target price and the three core pillars supporting our overweight recommendation",
    "Slide 4: Investment Thesis Summary and Target Price",
    [
        ("BASE CASE TARGET PRICE", "$133.64", "+33.6% upside vs $100.00 current price"),
        ("CLEAN RUN-RATE EBIT", "13.2%", "Pillar 1: Resilient profitability floor"),
        ("DISCOUNT RATE (WACC)", "9.0% / 2.3% g", "Pillar 2: Mathematically sound valuation"),
    ],
    [
        "Our discounted cash flow valuation generates a base case implied share price of 133.64 dollars representing a 33.6 percent upside from current levels [1]",
        "The first pillar is profitability because adjusting out the tariff refunds reveals LULU still maintains a highly resilient 13.2 percent clean run-rate operating margin [2]",
        "The second pillar is our mathematically sound 9.0 percent WACC which strictly bounds our 2.3 percent long-term revenue growth assumption [3]",
        "Finally international expansion remains the crucial growth engine as 4 percent growth in China Mainland easily offsets the temporary North American stagnation [4]",
    ],
    "Links & Sources: [1] Provided Valuation Model (Base Case Implied Value): LULUMODEL18_3.xlsx [2] Lululemon Q2 FY2026 Earnings Release (13.2% Margin Calc): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf [3] Provided Valuation Model (WACC & Revenue Drivers): LULUMODEL18_3.xlsx [4] Lululemon Q2 FY2026 Earnings Release (International Growth): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
)

# =====================================================================
# 5. GEOGRAPHIC SEGMENTS (PROMPT SLIDE 5)
# =====================================================================
build_prompt_slide(
    5,
    "Geographic Segments and Revenue Breakdown",
    "This slide outlines Lululemon's geographic segments and revenue breakdown while detailing why China growth offsets temporary US declines",
    "Slide 5",
    [
        ("TOTAL REVENUE (FY25)", "$11.1B", "10-K reported top-line foundation"),
        ("AMERICAS CONTRIBUTION", "$7.85B (70.68%)", "−3% comps reflect macro pressure"),
        ("CHINA MAINLAND SURGE", "+20% ($1.75B)", "High-growth international engine"),
    ],
    [
        "Lululemon generated 11.1 billion dollars in total revenue with Americas contributing 7.85 billion dollars or 70.68 percent [1]",
        "Americas comparable sales fell 3 percent due to temporary cyclical macro pressure rather than structural brand degradation [2]",
        "China Mainland surged 20 percent to 1.75 billion dollars proving high growth international expansion easily offsets US temporary weakness [3]",
    ],
    "Links & Sources: [1] Lululemon FY2025 Form 10-K (Segment Revenue and Percentages): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf [2] Lululemon FY2025 Form 10-K (Americas Comparable Sales): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf [3] Provided Valuation Model (China Mainland Growth & Revenue Driver): LULUMODEL18.xlsx",
)

# =====================================================================
# 6. BUSINESS MODEL & UNIT ECONOMICS (PROMPT SLIDE 6)
# =====================================================================
build_prompt_slide(
    6,
    "Business Model Unit Economics and Competitive Moats",
    "This slide analyzes Lululemon's unit economics and competitive moats that protect its long-term market leadership",
    "Slide 6: Business Model Unit Economics and Competitive Moats",
    [
        ("GROSS MARGIN", "56.6%", "Insulating moat against price competition"),
        ("DTC E-COMMERCE EBIT", "23.6%", "Zero retail occupancy drag economics"),
        ("STORE SALES PER SQ FT", "$1,426 / sq ft", "18.6% store-level EBIT margin"),
    ],
    [
        "Lululemon sustains a 56.6 percent gross margin providing a massive competitive moat against retail price competition [1]",
        "E-commerce unit economics remain elite with direct-to-consumer EBIT margins hitting 23.6 percent without retail occupancy drag [2]",
        "Company-operated stores generate 1,426 dollars per square foot and an 18.6 percent EBIT margin establishing high capital efficiency [3]",
    ],
    "Links & Sources: [1] Provided Valuation Model (Gross Margin): LULUMODEL18_6.xlsx [2] Provided Valuation Model (E-commerce EBIT Margin Driver): LULUMODEL18_6.xlsx [3] Provided Valuation Model (Store EBIT Margin & SPSF): LULUMODEL18_6.xlsx",
)

# =====================================================================
# 7. INDUSTRY OVERVIEW - PART 1 (PROMPT SLIDE 7)
# =====================================================================
build_prompt_slide(
    7,
    "Industry Overview - Trends and Structure",
    "This slide explores the ongoing athleisure industry trends including market fragmentation and the barriers to entry",
    "Slide 7: Industry Overview - Trends and Structure",
    [
        ("GLOBAL MARKET", "Fragmented", "Most competitors lack durable advantage"),
        ("PREMIUM SEGMENT", "High Barriers", "Scale, fabric R&D & brand equity protect"),
        ("LULULEMON ROE", "30.25%", "vs 6.81% retail industry median"),
    ],
    [
        "I argue the global athleisure market remains highly fragmented meaning most players lack a durable competitive advantage [1]",
        "The premium segment maintains high barriers to entry protecting global champions from temporary cyclical noise [2]",
        "Lululemon mathematically proves its moat by generating a 30.25 percent ROE far outpacing the 6.81 percent industry median [3]",
    ],
    "Links & Sources: [1] Market.us Media (Athleisure Market Fragmentation): https://media.market.us/athleisure-industry-statistics/ [2] Fortune Business Insights (Premium Athleisure Trends): https://www.fortunebusinessinsights.com/athleisure-market-110642 [3] FinanceCharts (Lululemon ROE & Retail Median): https://www.financecharts.com/stocks/LULU/growth/roe",
)

# =====================================================================
# 8. INDUSTRY OVERVIEW - PART 2 (PROMPT SLIDE 7 PART 2)
# =====================================================================
build_prompt_slide(
    8,
    "Industry Overview - Barriers to Entry and Profitability",
    "This slide dissects capital efficiency metrics comparing Lululemon's gross margin directly against legacy apparel competitors",
    "Slide 7 Part 2: Industry Overview - Barriers to Entry and Profitability",
    [
        ("LULULEMON GROSS MARGIN", "56.6%", "DTC scale prevents margin compression"),
        ("NIKE GROSS MARGIN", "43%", "Wholesale dependence & markdown drag"),
        ("UNDER ARMOUR GM", "46%", "Structural margin compression vs DTC"),
    ],
    [
        "Lululemon commands a 56.6 percent gross margin far exceeding traditional athletic apparel peers like Nike and Under Armour [1]",
        "Nike struggles to maintain a 43 percent margin while Under Armour hovers around 46 percent reflecting their wholesale dependence [2]",
        "Lululemon's direct-to-consumer scale prevents this structural margin compression and forms an impenetrable economic moat against industry price wars [3]",
    ],
    "Links & Sources: [1] Provided Valuation Model (Gross Margin Assumptions): LULUMODEL18.xlsx [2] Investing.com (Nike & Under Armour Historical Gross Margins): https://www.investing.com/pro/NYSE:NKE/explorer/gp_margin [3] ProAnalyst LULU Market Report (Competitor Margins & DTC Moat): https://lulu.proanalyst.ai/business",
)

# =====================================================================
# 9. THESIS I (PROMPT SLIDE 8)
# =====================================================================
build_prompt_slide(
    9,
    "Investment Thesis I",
    "This slide outlines the core contrarian investment thesis utilizing numerical evidence from the LULUMODEL18.xlsx file",
    "Slide 8 Investment Thesis I",
    [
        ("PRICE DRAWDOWN", ">50%", "Critically undervalued vs durable cash flow"),
        ("ADJ. GROSS MARGIN", "54.9%", "Protected by premium DTC channel mix"),
        ("TERMINAL GROWTH / UPSIDE", "2.25% g → $133.64", ">30% upside provides valuation buffer"),
    ],
    [
        "The 2026 price drop exceeding 50% has left Lululemon critically undervalued despite durable cash flows [1]",
        "Its premium Direct-to-Consumer revenue mix protects a massive 54.9% adjusted gross margin, mathematically justifying my intrinsic valuation thesis [2]",
        "Furthermore, even with Americas growth stagnating, sheer cash flow generation creates a robust intrinsic valuation buffer, where a conservative 2.25% terminal growth rate still yields over 30% upside to our $133.64 target price [3]",
    ],
    [
        "[1] https://everythingmoney.com/blog/lululemon-is-collapsing-burry-s-biggest-bet-4334",
        "[2] https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
        "[3] LULUMODEL18.xlsx, DCF Terminal Value & Valuation Summary Schedule is that ready",
    ],
)

# =====================================================================
# 10. THESIS II (PROMPT SLIDE 9)
# =====================================================================
build_prompt_slide(
    10,
    "Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor",
    "This slide details how partial margin recovery supported by core brand loyalty still drives a highly compelling valuation",
    "Slide 9: Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor",
    [
        ("PEAK EBIT MARGIN", "23.7%", "FY24 pandemic-era peak earnings power"),
        ("TROUGH MARGIN FLOOR", "13.2%", "Core brand loyalty sustains floor"),
        ("PARTIAL RECOVERY TO FY30", "15.5% EBIT", "Unlocks $133.64 base DCF share price"),
    ],
    [
        "Historical pandemic-era peak EBIT margins reached 23.7% in FY24 showing prior peak earnings power [1]",
        "Our valuation assumes a floor built on resilient baseline brand loyalty, proving Lululemon does not need to remain the hottest viral trend 100% of the time to sustain a 13.2% trough margin [2]",
        "A modest partial recovery to just 15.5% EBIT margin by FY30 still yields $133.64 per share [3]",
        "This proves returning to peak COVID profitability is completely unnecessary to unlock substantial market upside [4]",
    ],
    [
        "Links & Sources:",
        "[1] LULUMODEL18.xlsx / SEC EDGAR Form 10-K (FY24 Peak Margins)",
        "[2] LULUMODEL18.xlsx, Clean Run-Rate EBIT Margin & Brand Royalty Assumptions",
        "[3] LULUMODEL18.xlsx, Base Case DCF Valuation Summary",
        "[4] LULUMODEL18.xlsx, Discounted Cash Flow Valuation Summary",
    ],
)

# =====================================================================
# 11. THESIS III (PROMPT SLIDE 10)
# =====================================================================
build_prompt_slide(
    11,
    "Investment Thesis III: Geographic Growth Divergence",
    "This slide examines how Lululemon's top-line projections rely disproportionately on Chinese market expansion to conceal domestic North American stagnation.",
    "Slide 10:",
    [
        ("AMERICAS COMPS", "−4.0% → 2.0%", "Near-term contraction flatlining by FY30"),
        ("CHINA FOOTPRINT EXPANSION", "16 stores / yr", "Disproportionate vs 6 domestically"),
        ("CHINA COMPS AVERAGE", "+10.2%", "Double-digit engine offsets Americas drag"),
    ],
    [
        "The revenue build reveals Americas facing near-term contraction with -4.0% comps in FY26 flatlining at a terminal 2.0% growth rate by FY30 [1].",
        "To offset this domestic anchor, the model relies entirely on disproportionate FY26–FY30 Chinese footprint expansion, averaging 16 new stores annually versus just 6 domestically, and sustained double-digit (10.2% average) comp growth to overcome clear Americas expansion drawbacks [2].",
        "Consequently, if the Chinese consumer softens, this model's core top-line projections will not be optimal enough to meet our target [3].",
    ],
    "[1] LULUMODEL18.xlsx, Americas FY26-FY30 Comparable Sales Growth Assumptions [2] LULUMODEL18.xlsx, FY26 Store Openings & Mainland China Revenue Build [3] LULUMODEL18.xlsx, Revenue Drivers Schedule",
    extra_heading="Investment Thesis III: Geographic Growth Divergence This slide examines how Lululemon's top-line projections rely disproportionately on Chinese market expansion to conceal domestic North American stagnation.",
)

# =====================================================================
# 12. RISK & MITIGANTS (PROMPT SLIDE 11)
# =====================================================================
build_prompt_slide(
    12,
    "Risk & Mitigants",
    "This slide evaluates core downside risks and demonstrates how share repurchases compound EPS to drive share price recovery",
    "Slide 11 Risk & Mitigants",
    [
        ("BEAR CASE VALUATION", "$56.00", "Recovers 45% of entry price in cash ($45.64)"),
        ("CAPEX & INVENTORY", "$360M / yr", "Slashed CapEx + $101.5M inventory release"),
        ("SHARE REPURCHASES", "18.7M shares", "$500M/yr buybacks compound EPS to $13.36"),
    ],
    [
        "China deceleration risks a $56.00 bear floor, but $45.64 cumulative cash per share recovers 45% of entry price [1]",
        "Slashed CapEx saves $360M annually by relying on online e-commerce's 23.6% EBIT margin plus $101.5M inventory releases [2]",
        "Deploying $500M annually into buybacks retires 18.7 million shares, compounding EPS to $13.36 to support our target price of $133.64.",
    ],
    "[1] LULUMODEL18.xlsx, Bear Case DCF Valuation Summary / https://www.barrons.com/articles/lululemon-stock-earnings-guidance-a7a7c5c0 [2] LULUMODEL18.xlsx, CapEx & E-Commerce Channel EBIT Assumptions / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733 [3] LULUMODEL18.xlsx, Share Repurchase & EPS Accretion Schedule / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
)

# =====================================================================
# 13. TIMELINE OF RECOVERY (PROMPT SLIDE 12)
# =====================================================================
def build_timeline_slide():
    s = slide_base("Timeline of Recovery", "Slide 12: Timeline of Recovery", page=pg())

    cats = [
        ("2026 Q3 Trough", "2026 (Q3 FY2026 Trough): Q3 FY2026 revenue laps guidance trough while China Double 11 sales confirm holiday store traffic floor stabilization across key markets [1]"),
        ("2027 Year-End", "2027 (FY2026 Year-End): First full-year reset absorbs steep prior declines while $134.5M tariff refunds and targeted SG&A cost actions protect earnings per share [2]"),
        ("2027 Inflection", "2027 (FY2027 Margin Inflection): Operating margins expand to 13.8% (+60 bps) as promotional headwinds anniversary, while 75% UFCF buybacks compound EPS to boost sentiment [3]"),
        ("2028 Re-Rating", "2028 (FY2027–FY2028 Multiple Re-Rating): Accelerating international store scaling (+12% China comps) offsets Americas softness (-4%), driving overall revenue recovery and valuation re-rating toward $133.64 [4]"),
    ]

    top = Inches(1.22)
    row_h = Inches(1.00)
    gap = Inches(0.12)
    for date_lbl, full_text in cats:
        b1 = rect(s, Inches(0.5), top, Inches(2.2), row_h, fill=NAVY)
        bt1 = b1.text_frame
        bt1.word_wrap = True
        bt1.vertical_anchor = MSO_ANCHOR.MIDDLE
        add_para(bt1, date_lbl, 12, GOLD, bold=True, first=True, align=PP_ALIGN.CENTER, space_after=0)

        b2 = rect(s, Inches(2.8), top, Inches(10.05), row_h, fill=LGREY)
        bt2 = b2.text_frame
        bt2.word_wrap = True
        bt2.vertical_anchor = MSO_ANCHOR.MIDDLE
        bt2.margin_left = Pt(10)
        bt2.margin_right = Pt(10)
        add_para(bt2, full_text, 11.5, INK, first=True, space_after=0)
        top += row_h + gap

    # Sources box
    src_y = Inches(5.85)
    src_h = Inches(1.22)
    sb = rect(s, Inches(0.5), src_y, Inches(12.35), src_h, fill=LGREY)
    stf = sb.text_frame
    stf.word_wrap = True
    stf.margin_top = Pt(4)
    stf.margin_bottom = Pt(4)
    stf.margin_left = Pt(8)
    stf.margin_right = Pt(8)
    src_txt = "[1] LULUMODEL18.xlsx, Scenarios & Revenue Drivers / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733 [2] LULUMODEL18.xlsx, NOPAT Bridge & Scenarios / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733 [3] LULUMODEL18.xlsx, Scenarios & Unlevered Free Cash Flow Schedule / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733 [4] LULUMODEL18.xlsx, Revenue Drivers & DCF Valuation Summary / https://stockanalysis.com/stocks/lulu/forecast/"
    add_para(stf, src_txt, 8.5, GREY, first=True, space_after=0)
    return s

build_timeline_slide()

def m(v):
    return f"{v:,.0f}"


HY = ["FY2022", "FY2023", "FY2024", "FY2025"]
FIN_FOOTNOTE = (
    "Forecast: base = Scenarios col G (3-statement); bull = Scenarios col H. "
    "Historicals per FY2025 10-K. Each row uses the same metric definition across periods."
)


def _hist_m(section, key, years=HY):
    """Historical line item in US$M — same units as forecast JSON."""
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
    """FCF = CFO − capex — same definition as forecast (CFO + negative capex)."""
    return [m((D.CF["cfo"][y] - D.CF["capex"][y]) / 1000) for y in years]


def _hist_buybacks_out(years=HY):
    return [f"({m(D.CF['buybacks'][y] / 1000)})" for y in years]


def _fin_headers():
    hdr = ["US$ M", "22", "23", "24", "25"]
    for fy in PROJ_YEARS:
        yr = fy[4:6]
        hdr.extend([f"{yr}B", f"{yr}U"])
    return hdr


def _fin_explainer_boxes(slide, hist_title, hist_bullets, fcst_title, fcst_bullets):
    """Side-by-side Historical vs forecast explainer panels above the financial table."""
    for x, title, bullets, fill in (
        (0.5, hist_title, hist_bullets, LGREY),
        (6.7, fcst_title, fcst_bullets, NAVY),
    ):
        box = rect(slide, Inches(x), Inches(1.12), Inches(6.05), Inches(1.55), fill=fill)
        btf = box.text_frame
        btf.word_wrap = True
        title_color = CARD if fill == LGREY else GOLD
        body_color = INK if fill == LGREY else WHITE
        add_para(btf, title, 11.5, title_color, bold=True, first=True, space_after=3)
        for i, bullet in enumerate(bullets):
            add_para(btf, bullet, 10.5, body_color, bullet=True, first=(i == 0), space_after=2)


def _fmt_cell(v, fmt="num"):
    if fmt == "pct":
        return v
    if fmt == "eps":
        return f"{v:.2f}"
    if fmt == "neg_paren":
        return f"({abs(v):,})"
    return _m(v)


def _dual_year_vals(base_sec, bull_sec, key, by_year=True, fmt="num"):
    vals = []
    for fy in PROJ_YEARS:
        if by_year:
            b, u = base_sec[fy][key], bull_sec[fy][key]
        else:
            b, u = base_sec[key][fy], bull_sec[key][fy]
        vals.append(_fmt_cell(b, fmt))
        vals.append(_fmt_cell(u, fmt))
    return vals


def pitch_financial_slide(
    title,
    subtitle,
    rows,
    hist_title,
    hist_bullets,
    fcst_title,
    fcst_bullets,
    bold_rows=(),
    italic_note=None,
    col0w=2.35,
    table_top=2.82,
    table_height=3.55,
):
    """Build one financial slide: historicals + 5yr base/bull forecast table."""
    hdr = _fin_headers()
    s = slide_base(title, subtitle, page=pg(), sources=FIN_FOOTNOTE)
    _fin_explainer_boxes(s, hist_title, hist_bullets, fcst_title, fcst_bullets)
    stmt_table(s, rows, hdr, col0w=col0w, top=table_top, height=table_height,
               bold_rows=bold_rows, font_size=8.5, header_font_size=8)
    if italic_note:
        tb, tf = textbox(s, Inches(0.5), Inches(6.45), Inches(12.35), Inches(0.55))
        add_para(tf, italic_note, 10, GREY, italic=True, first=True, space_after=0)
    return s


# =====================================================================
# 14–16. FINANCIALS (historicals + base & bull, all five forecast years)
# =====================================================================
_IS26B, _IS30B = IS_BASE["FY2026E"], IS_BASE["FY2030E"]
_IS26U, _IS30U = IS_BULL["FY2026E"], IS_BULL["FY2030E"]

pitch_financial_slide(
    "Financials \u2014 Income Statement",
    "Reported history (10-K) vs base & bull operating forecast \u2014 all FY26\u201330 (US$ M)",
    [
        ["Net revenue"] + _hist_m(D.IS, "revenue") + _dual_year_vals(IS_BASE, IS_BULL, "revenue"),
        ["Gross profit"] + _hist_m(D.IS, "gross_profit") + _dual_year_vals(IS_BASE, IS_BULL, "gross_profit"),
        ["Operating income"] + _hist_m(D.IS, "operating_income") + _dual_year_vals(IS_BASE, IS_BULL, "operating_income"),
        ["Operating margin %"] + _hist_om() + _dual_year_vals(IS_BASE, IS_BULL, "operating_margin", fmt="pct"),
        ["Net income"] + _hist_m(D.IS, "net_income") + _dual_year_vals(IS_BASE, IS_BULL, "net_income"),
        ["Diluted EPS ($)"] + _hist_eps() + _dual_year_vals(IS_BASE, IS_BULL, "eps", fmt="eps"),
    ],
    hist_title="HISTORICALS (FY2022\u2013FY2025) \u2014 SEC 10-K",
    hist_bullets=[
        "Reported figures from Forms 10-K; FY2025 year ended February 1, 2026",
        "Revenue $8,111M \u2192 $11,103M; OM peaked 23.7% (FY24); EPS $6.68 \u2192 $13.26",
        "Share count 128.0M \u2192 119.1M via buybacks",
    ],
    fcst_title="FORECAST (FY2026E\u2013FY2030E) \u2014 base (G) vs bull (H)",
    fcst_bullets=[
        f"Base: FY26 {m(_IS26B['revenue'])}M rev / {_IS26B['operating_margin']} OM \u2192 FY30 {m(_IS30B['revenue'])}M / {_IS30B['operating_margin']}",
        f"Bull: FY26 {m(_IS26U['revenue'])}M (\u22124% guide) \u2192 FY30 {m(_IS30U['revenue'])}M; OM {_IS26U['operating_margin']} \u2192 {_IS30U['operating_margin']}",
        f"Bull EPS ${_IS26U['eps']:.2f} \u2192 ${_IS30U['eps']:.2f} (DCF implied price {_d(V['bull'])})",
    ],
    bold_rows=(0, 2, 5),
    italic_note="22\u201325 = FY2022\u201325 historical; 26B/26U = FY2026 base (G) / bull (H); same for 27\u201330. Bear in Appendix.",
)

pitch_financial_slide(
    "Financials \u2014 Balance Sheet",
    "Net-cash history vs base & bull funded growth \u2014 all FY26\u201330 (US$ M)",
    [
        ["Cash & equivalents"] + _hist_m(D.BS, "cash") + _dual_year_vals(BS_BASE, BS_BULL, "cash", by_year=False),
        ["Inventories"] + _hist_m(D.BS, "inventories") + _dual_year_vals(BS_BASE, BS_BULL, "inventories", by_year=False),
        ["Total assets"] + _hist_m(D.BS, "total_assets") + _dual_year_vals(BS_BASE, BS_BULL, "total_assets", by_year=False),
        ["Total liabilities"] + _hist_m(D.BS, "total_liab") + _dual_year_vals(BS_BASE, BS_BULL, "total_liab", by_year=False),
        ["Total equity"] + _hist_m(D.BS, "total_equity") + _dual_year_vals(BS_BASE, BS_BULL, "total_equity", by_year=False),
        ["Funded debt"] + ["0"] * 4 + ["0"] * (len(PROJ_YEARS) * 2),
    ],
    hist_title="HISTORICALS (FY2022\u2013FY2025) \u2014 SEC 10-K",
    hist_bullets=[
        "No funded debt; FY25 cash $1,807M; total assets $8,457M",
        "Net-cash balance sheet underpins downside at ~10x earnings",
        "Lease-backed ROU assets ~$1.6B \u2014 not traditional debt",
    ],
    fcst_title="FORECAST (FY2026E\u2013FY2030E) \u2014 base vs bull",
    fcst_bullets=[
        f"Base cash {m(BS_BASE['cash']['FY2026E'])}M \u2192 {m(BS_BASE['cash']['FY2030E'])}M ($500M/yr buybacks)",
        f"Bull cash {m(BS_BULL['cash']['FY2026E'])}M \u2192 {m(BS_BULL['cash']['FY2030E'])}M (75% FCF to buybacks)",
        "Bull BS scaled from base WC drivers at bull revenue; cash from FCF waterfall",
    ],
    bold_rows=(2, 4, 5),
    italic_note="Same BS line items as 10-K. Bull cash = FY25 cash + cumulative (FCF + buybacks).",
)

pitch_financial_slide(
    "Financials \u2014 Cash Flow",
    "Operating FCF (DCF input) vs financing \u2014 historical 10-K vs base & bull (US$ M)",
    [
        ["Cash from operations"] + _hist_m(D.CF, "cfo") + _dual_year_vals(CF_BASE, CF_BULL, "cfo", by_year=False),
        ["D&A (add-back)"] + _hist_m(D.CF, "d_and_a") + _dual_year_vals(CF_BASE, CF_BULL, "dna", by_year=False),
        ["Capital expenditures"] + _hist_capex_out() + _dual_year_vals(CF_BASE, CF_BULL, "capex", by_year=False, fmt="neg_paren"),
        ["Free cash flow"] + _hist_fcf() + _dual_year_vals(CF_BASE, CF_BULL, "fcf", by_year=False),
        ["Share repurchases (CFF)"] + _hist_buybacks_out() + _dual_year_vals(CF_BASE, CF_BULL, "buybacks", by_year=False, fmt="neg_paren"),
    ],
    hist_title="HISTORICALS (FY2022\u2013FY2025) \u2014 SEC 10-K",
    hist_bullets=[
        "FY25 CFO $1,603M; capex ($681M) \u2192 ~$922M FCF",
        "Buybacks ($1,178M) FY25 are CFF \u2014 reported on 10-K, not in DCF UFCF",
        "FCF = CFO \u2212 capex; same operating definition as forecast rows above buybacks",
    ],
    fcst_title="FORECAST (FY2026E\u2013FY2030E) \u2014 base vs bull",
    fcst_bullets=[
        f"FCF (DCF input): base ${m(CF_BASE['fcf']['FY2026E'])}M \u2192 ${m(CF_BASE['fcf']['FY2030E'])}M",
        f"Bull FCF ${m(CF_BULL['fcf']['FY2026E'])}M \u2192 ${m(CF_BULL['fcf']['FY2030E'])}M (Scenarios col H)",
        "Buybacks below = 3-statement CFF only; DCF IV uses basic 111.4M shares (no retirements)",
    ],
    bold_rows=(3,),
    italic_note=(
        "FCF rows tie to DCF unlevered FCF. Share repurchases are financing (CFF) \u2014 shown for 10-K / "
        "3-statement comparability and cash bridge; they do not change DCF implied price (basic share count)."
    ),
)

# =====================================================================
# 17. CAPITAL STRUCTURE & WACC
# =====================================================================
s = slide_base("Capital Structure & WACC", f"A lease-adjusted capital structure drives a ~{_pct(V['wacc'])} discount rate", page=pg())
tb, tf = textbox(s, Inches(0.5), Inches(1.2), Inches(6.2), Inches(5.6))
add_para(tf, "Capital structure", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "No funded debt; an undrawn revolving credit facility provides liquidity",
    "$1,807.2M cash and equivalents \u2014 a net-cash position",
    "Market EV at ~$100 is ~$9.3B (~3.5\u00d7 FY25 EBITDA); intrinsic DCF enterprise value is ~$15B",
    "We relever Yahoo 5Y \u03b2 (0.86) at FY25 ASC 842 lease debt \u00f7 market cap (~0.16 D/E) \u2192 \u03b2 used \u2248 0.84. Damodaran Retail \u03b2 (0.95) is shown for reference only.",
    "Capital returned via buybacks (no dividend); FY2025 repurchases $1,178.3M",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)
box = rect(s, Inches(7.0), Inches(1.2), Inches(5.85), Inches(4.6), fill=LGREY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "WACC BUILD (CAPM)", 13, CARD, bold=True, first=True, space_after=8)
for t, v in [
    ("Risk-free rate (10-yr UST)", _pct(V["rf"])),
    ("Equity risk premium", _pct(V["erp"])),
    ("Beta (relevered; lease adj.)", f"{V['beta']:.2f}"),
    ("Cost of equity", _pct(V["coe"])),
    ("Lease-debt weight", "14%"),
    ("Equity weight", "86%"),
]:
    p = btf.add_paragraph(); p.space_after = Pt(6)
    r = p.add_run(); r.text = f"{t}"; _set_font(r, 13, INK, bold=(t in ("Cost of equity",)))
    r2 = p.add_run(); r2.text = f"      {v}"; _set_font(r2, 13, NAVY, bold=True)
p = btf.add_paragraph(); p.space_before = Pt(4)
r = p.add_run(); r.text = f"WACC = {_pct(V['wacc'])}"; _set_font(r, 16, GREEN, bold=True)

# =====================================================================
# 18. VALUATION SUMMARY (FOOTBALL FIELD)
# =====================================================================
s = slide_base("Valuation Summary", "Multiple methods converge above the current price \u2014 target $133.64", page=pg(),
               sources="Source: GIS DCF and comps models; multiples are analyst ranges")
# football field: horizontal floating bars
methods = [
    ("P / E (10\u201318x FY2026E)", round(FF["P / E"]["low"]), round(FF["P / E"]["high"])),
    ("EV / EBITDA (4.5\u20137.5x)", round(FF["EV / EBITDA"]["low"]), round(FF["EV / EBITDA"]["high"])),
    ("DCF (bear \u2013 bull)", round(FF["DCF"]["low"]), round(FF["DCF"]["high"])),
    ("52-week range", 100, 226),
]
chart_l, chart_r = 3.2, 12.6
vmin, vmax = 60, 250
def xpos(v):
    return chart_l + (v - vmin) / (vmax - vmin) * (chart_r - chart_l)
top = 1.7
for name, lo, hi in methods:
    tb, tf = textbox(s, Inches(0.5), Inches(top-0.02), Inches(2.6), Inches(0.5), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, name, 12, NAVY, bold=True, first=True, space_after=0)
    bar = rect(s, Inches(xpos(lo)), Inches(top), Inches(xpos(hi)-xpos(lo)), Inches(0.45), fill=GOLD)
    bt = bar.text_frame; bt.word_wrap = False; bt.vertical_anchor = MSO_ANCHOR.MIDDLE
    bt.margin_left = Pt(4); bt.margin_right = Pt(4)
    p = bt.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = f"${lo}"; _set_font(r, 10.5, NAVY, bold=True)
    tb2, tf2 = textbox(s, Inches(xpos(hi)+0.02), Inches(top), Inches(0.9), Inches(0.45), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf2, f"${hi}", 10.5, NAVY, bold=True, first=True, space_after=0)
    top += 0.75
# current price line and target line
cp_x = xpos(100); pt_x = xpos(133.64)
ln = rect(s, Inches(cp_x), Inches(1.55), Pt(2), Inches(3.4), fill=INK)
ln2 = rect(s, Inches(pt_x), Inches(1.55), Pt(2), Inches(3.4), fill=CARD)
tb, tf = textbox(s, Inches(cp_x-0.7), Inches(4.95), Inches(1.6), Inches(0.3))
add_para(tf, "Current $100", 10.5, INK, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
tb, tf = textbox(s, Inches(pt_x-0.7), Inches(5.2), Inches(1.6), Inches(0.3))
add_para(tf, "Target $133.64", 10.5, CARD, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
tb, tf = textbox(s, Inches(0.5), Inches(5.7), Inches(12.35), Inches(1.2))
add_para(tf, f"We set a 12-month target of $133.64 \u2014 based on the DCF base case ({_d(V['base_dcf'])}) \u2014 implying +33.6% upside", 13, NAVY, bold=True, first=True, space_after=5)
add_para(tf, "The P/E low end (~$96) sits near today's price, showing how little recovery is required for the stock to work", 12.5, INK, italic=True, space_after=0)

# =====================================================================
# 19. DCF VALUATION
# =====================================================================
s = slide_base("DCF Valuation", f"Base-case unlevered DCF yields ~{_d(V['base_dcf'])} per share", page=pg(),
               sources="Source: GIS DCF model (from scratch); FY2025 cash & share count per 10-K")
# assumptions table
tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(6.1), Inches(0.4))
add_para(tf, "Base-case assumptions", 14, CARD, bold=True, first=True, space_after=0)
arows = [
    ["Assumption", "Value"],
    ["Revenue growth (FY26 \u2192 FY30)", "\u22126.1% \u2192 +2.3%"],
    ["EBIT margin (FY26 \u2192 FY30)", "13.2% clean \u2192 15.5% (+$134.5M FY26)"],
    ["Tax rate", "30%"],
    ["Capex % of revenue", "5.5% blend"],
    ["WACC", _pct(V["wacc"])],
    ["Terminal growth", "2.25%"],
]
t = s.shapes.add_table(len(arows), 2, Inches(0.5), Inches(1.6), Inches(6.0), Inches(3.0)).table
t.columns[0].width = Inches(4.0); t.columns[1].width = Inches(2.0)
for ri, row in enumerate(arows):
    for ci, val in enumerate(row):
        cell = t.cell(ri, ci); cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if ri == 0 else (WHITE if ri % 2 else LGREY)
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT
        r = p.add_run(); r.text = val
        _set_font(r, 11.5, WHITE if ri == 0 else (CARD if ci == 1 and ri > 0 else INK), bold=(ri == 0))
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
# output box
box = rect(s, Inches(7.0), Inches(1.6), Inches(5.85), Inches(3.0), fill=LGREY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "VALUATION OUTPUT (US$ M)", 12.5, CARD, bold=True, first=True, space_after=6)
for t2, v in [
    ("PV of explicit FCF (FY26\u2013FY30)", _m(round(V["pv_fcf_m"]))),
    ("PV of terminal value", _m(round(V["pv_tv_m"]))),
    ("Enterprise value", _m(round(V["ev_m"]))),
    ("Plus: cash", _m(round(V["cash_m"]))),
    ("Equity value", _m(round(V["equity_m"]))),
    ("\u00f7 Diluted shares (M)", f"{V['shares_m']:.1f}"),
]:
    p = btf.add_paragraph(); p.space_after = Pt(5)
    r = p.add_run(); r.text = t2; _set_font(r, 12.5, INK, bold=("Enterprise" in t2 or "Equity" in t2))
    r2 = p.add_run(); r2.text = f"      {v}"; _set_font(r2, 12.5, NAVY, bold=True)
p = btf.add_paragraph(); p.space_before = Pt(6)
r = p.add_run(); r.text = f"Implied value:  {_d(V['base_dcf'])} / share"; _set_font(r, 16, GREEN, bold=True)
# sensitivity mini
tb, tf = textbox(s, Inches(0.5), Inches(4.85), Inches(12.35), Inches(2.0))
add_para(tf, "Sensitivity \u2014 implied share price (WACC vs terminal growth)", 13, CARD, bold=True, first=True, space_after=4)
_g_cols = ["1.5%", "2.0%", "2.25%", "2.5%", "3.0%"]
sens = [["WACC \\ g"] + _g_cols]
_base_wacc_idx = min(range(len(PV["sensitivity"])), key=lambda i: abs(float(PV["sensitivity"][i]["wacc"].rstrip("%")) - V["wacc"] * 100))
for row in PV["sensitivity"]:
    sens.append([row["wacc"]] + [_d(p) for p in row["prices"]])
st = s.shapes.add_table(len(sens), 6, Inches(0.5), Inches(5.35), Inches(7.6), Inches(1.5)).table
for ri, row in enumerate(sens):
    for ci, val in enumerate(row):
        cell = st.cell(ri, ci); cell.fill.solid()
        hot = (ri == _base_wacc_idx + 1 and ci == 3)
        cell.fill.fore_color.rgb = NAVY if (ri == 0 or ci == 0) else (GOLD if hot else (WHITE if ri % 2 else LGREY))
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = val
        _set_font(r, 10.5, WHITE if (ri == 0 or ci == 0) else (NAVY if hot else INK), bold=(ri == 0 or ci == 0 or hot))
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

# =====================================================================
# 20. COMPS
# =====================================================================
s = slide_base("Comparable Companies", "PitchBook pubcomps are the current tape; Alo still has no EV/EBITDA", page=pg(),
               sources="Source: PitchBook Comps Set 04-Sep-2026 (EV/EBITDA = daily EV / TTM EBITDA); Alo ask $10bn / Forbes ~$2bn")
hdr = ["Company", "EV/EBITDA", "P/E", "EV ($000)", "EBITDA ($000)"]
rows = [
    ["lululemon (LULU)", "4.7x", "8.3x", "11,890,440", "2,548,084"],
    ["Nike (NKE)", "12.7x", "18.3x", "58,972,350", "4,647,000"],
    ["adidas (ADS)", "9.2x", "19.1x", "35,661,944", "3,857,684"],
    ["Deckers (DECK)", "7.9x", "12.1x", "10,555,510", "1,329,439"],
    ["Crocs (CROX)", "7.8x", "10.3x", "7,153,911", "922,278"],
    ["Levi Strauss (LEVI)", "9.6x", "15.0x", "9,422,753", "976,700"],
    ["Kontoor (KTB)", "12.9x", "18.0x", "5,236,269", "407,521"],
    ["Alo Yoga (private)", "n.a.", "n.m.", "\u2014", "\u2014"],
]
stmt_table(s, rows, hdr, col0w=3.4, top=1.20, height=3.90, bold_rows=(0,))
tb, tf = textbox(s, Inches(0.5), Inches(5.20), Inches(12.35), Inches(1.75))
add_para(tf, "Read-through \u2014 we do not average these for terminal value", 13.5, CARD, bold=True, first=True, space_after=3)
add_para(tf, "PitchBook core mean (ex-UAA, ex-WSM/MOV/LZB) is the current athletic/apparel tape. UAA has negative EBITDA so it is out of the average", 12.5, INK, bullet=True, space_after=3)
add_para(tf, "Selected TV is the Gordon identity (~7.4x FY30 EBITDA at base WACC/g). A 2.25% g year is not today\u2019s NKE 12.7x or WSM 15.4x", 12.5, INK, bullet=True, space_after=3)
add_para(tf, "Alo EV/EBITDA is still n.a. Implied 5.0x is EV/Sales on an unclosed $10bn ask \u2014 not in this PitchBook set and not the TV", 12.5, INK, bullet=True, space_after=0)

# =====================================================================
# 21. APPENDIX: SCENARIOS
# =====================================================================
s = slide_base("Appendix \u2014 Bull / Bear Scenarios", "Asymmetric payoff: limited downside, substantial upside", page=pg(),
               sources="Source: GIS DCF model (scenario tab)")
cols = [
    ("BEAR", _d(V["bear"]), _up(V["bear"]), CARD, [
        "FY2026 revenue \u22129%; growth stays negative (\u22121% avg FY28\u201330)",
        "Terminal EBIT margin 12.0%",
        "WACC 11.0%; terminal growth 1.5%",
        "Americas decline persists; share loss continues",
    ]),
    ("BASE", _d(V["base_dcf"]), _up(V["base_dcf"]), GREEN, [
        "FY2026 revenue \u22126.1%; then +2.3% (Street 3Y forecast)",
        "Clean EBIT margin 13.2% \u2192 15.5%; +$134.5M refund in FY26 only",
        f"WACC {_pct(V['wacc'])}; terminal growth 2.25%",
        "International offsets a stabilizing Americas",
    ]),
    ("BULL", _d(V["bull"]), _up(V["bull"]), NAVY, [
        "FY2026 revenue \u22124%; +6% avg FY27\u201330",
        "Terminal EBIT margin 19.0%",
        "WACC 9.5%; terminal growth 3.0%",
        "Margin recovery toward peak; brand re-accelerates",
    ]),
]
x = 0.5
for name, px, up, color, bullets in cols:
    head = rect(s, Inches(x), Inches(1.35), Inches(4.05), Inches(0.95), fill=color)
    ht = head.text_frame; ht.word_wrap = True; ht.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(ht, f"{name}   {px}", 20, WHITE, bold=True, first=True, space_after=0, align=PP_ALIGN.CENTER)
    add_para(ht, f"{up} vs $100", 12.5, WHITE, align=PP_ALIGN.CENTER, space_after=0)
    b = rect(s, Inches(x), Inches(2.4), Inches(4.05), Inches(3.6), fill=LGREY)
    bt = b.text_frame; bt.word_wrap = True
    for i, blt in enumerate(bullets):
        add_para(bt, blt, 12.5, INK, bullet=True, first=(i == 0), space_after=8)
    x += 4.25
tb, tf = textbox(s, Inches(0.5), Inches(6.2), Inches(12.35), Inches(0.8))
add_para(tf, f"Probability-weighted value (25% / 50% / 25%) \u2248 {_d(V['prob_weighted'])} \u2014 the risk/reward skews decisively to the upside", 13.5, NAVY, bold=True, italic=True, first=True, space_after=0)

# =====================================================================
# 22. DISCLAIMER / SOURCES
# =====================================================================
s = slide_base("Sources & Disclaimer", "Data provenance and standard research disclaimer", page=pg())
tb, tf = body_box(s)
add_para(tf, "Sources", 14, CARD, bold=True, first=True, space_after=5)
for t in [
    "Historical financials: lululemon athletica inc. Forms 10-K via SEC EDGAR (CIK 0001397187); FY2025 fiscal year ended February 1, 2026",
    "Q2 FY2026 results and FY2026 guidance: company earnings release dated September 3, 2026",
    "Market data (price, shares, beta): public market sources as of early September 2026",
    "Projections, DCF and comparable-company analysis: GIS Investment Research models, built from scratch for this assignment",
    "Alo Yoga: Reuters 2023 Moelis ask ~$10bn (no deal announced, Reuters 2026) and Forbes Color Image sales nearly $2bn. No page prints EV/EBITDA; implied 5.0x is EV/Sales, not the FY30 exit",
]:
    add_para(tf, t, 12.5, INK, bullet=True, space_after=5)
add_para(tf, "Disclaimer", 14, CARD, bold=True, space_after=5)
add_para(tf, "This presentation is prepared for educational purposes as part of the Global Investment Society selection process and does not constitute investment advice or a recommendation to buy or sell any security", 12, GREY, italic=True, space_after=0)

n = deck.save(OUT)
print("Saved", os.path.abspath(OUT), "with", n, "slides")
try:
    from gis_pitch import export_pdf
    pdf = export_pdf(OUT)
    print("Saved", pdf)
except Exception as e:
    print("PDF export skipped:", e)
