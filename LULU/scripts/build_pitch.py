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
    target_price="$140",
    upside_pct="+40% upside",
    descriptor="A net-cash, high-margin brand priced for terminal decline \u2014 we see a cyclical trough, not a broken business",
)

# =====================================================================
# 2. TABLE OF CONTENTS (standard GIS + optional macro slide)
# =====================================================================
LULU_TOC = [
    "1.  Investment thesis summary", "2.  Situation overview", "3.  Macro \u00d7 micro overlap",
    "4.  Market narrative", "5.  Company overview", "6.  Business model & unit economics",
    "7.  Industry overview", "8.  Thesis I \u2014 Priced for terminal decline",
    "9.  Thesis II \u2014 International growth engine", "10. Thesis III \u2014 Elite economics & capital return",
    "11. Risks & mitigants", "12. Catalyst timeline", "13. Financials \u2014 income statement",
    "14. Financials \u2014 balance sheet", "15. Financials \u2014 cash flow",
    "16. Capital structure & WACC", "17. Valuation summary (football field)",
    "18. DCF valuation", "19. Comparable companies", "20. Appendix \u2014 bull / bear scenarios",
]
deck.toc_slide(LULU_TOC)

# =====================================================================
# 3. INVESTMENT THESIS SUMMARY
# =====================================================================
s = slide_base("Investment Thesis Summary", "Overweight LULU with a $140 target on ~40% upside and an asymmetric payoff", page=pg())
# target box
box = rect(s, Inches(0.5), Inches(1.2), Inches(3.9), Inches(1.5), fill=LGREY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "PRICE TARGET", 12, CARD, bold=True, first=True, space_after=2)
add_para(btf, "$140", 30, NAVY, bold=True, space_after=0)
add_para(btf, "+40% vs $100.00 today", 12, GREEN, bold=True, space_after=0)
# rating box
box2 = rect(s, Inches(4.6), Inches(1.2), Inches(3.9), Inches(1.5), fill=LGREY)
b2 = box2.text_frame; b2.word_wrap = True
add_para(b2, "RATING", 12, CARD, bold=True, first=True, space_after=2)
add_para(b2, "OVERWEIGHT", 24, NAVY, bold=True, space_after=0)
add_para(b2, "12-month horizon", 12, GREY, space_after=0)
# setup box
box3 = rect(s, Inches(8.7), Inches(1.2), Inches(4.15), Inches(1.5), fill=LGREY)
b3 = box3.text_frame; b3.word_wrap = True
add_para(b3, "WHY NOW", 12, CARD, bold=True, first=True, space_after=2)
add_para(b3, "Shares \u201355% off highs; ~18% single-day drop after Q2 FY2026 print (Sep 3, 2026) overshoots the fundamentals", 12.5, INK, space_after=0)
tb, tf = textbox(s, Inches(0.5), Inches(2.95), Inches(12.35), Inches(4.0))
add_para(tf, "Three reasons to be long", 15, CARD, bold=True, first=True, space_after=6)
add_para(tf, "1.  Priced for terminal decline \u2014 at ~3.5x EV/EBITDA and ~10x FY2026E EPS with a net-cash balance sheet, the stock embeds a permanent impairment that the business does not support", 13.5, INK, bold=False, space_after=7)
add_para(tf, "2.  International is a multi-year growth engine \u2014 China Mainland and Rest-of-World more than offset a maturing Americas and can return total revenue to low-single-digit growth (~2.3% FY27\u201330 in our model)", 13.5, INK, space_after=7)
add_para(tf, "3.  Elite economics + accretive buybacks \u2014 FY2025 operating margin was 19.9% (last full year, not the trough); we model a 13.2% clean FY2026 run-rate (Q2 18.8% minus 560bps of tariff refunds), then add the $134.5M refund once (~15.2% reported in our IS). Diluted shares are down ~7% (128.0M \u2192 119.1M)", 13.5, INK, space_after=7)
add_para(tf, f"Base-case DCF {_d(V['base_dcf'])}; bear {_d(V['bear'])} ({_up(V['bear'])}) vs bull {_d(V['bull'])} ({_up(V['bull'])}) \u2014 downside is protected by net cash and an ~8\u20139% FCF yield", 13, NAVY, bold=True, italic=True, space_after=0)

# =====================================================================
# 4. SITUATION OVERVIEW
# =====================================================================
s = slide_base("Situation Overview", "A high-quality compounder has de-rated to value multiples after a guidance reset", page=pg(),
               sources="Source: company 10-K (FY2025, ended Feb 1, 2026) and Q2 FY2026 release (Sep 3, 2026)")
tb, tf = body_box(s)
add_para(tf, "What happened", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "LULU shares fell from a 52-week high of ~$226 to ~$100, including an ~18% single-day decline following the Q2 FY2026 earnings release on September 3, 2026",
    "Management guided FY2026 to the first annual revenue decline in company history: net revenue of $10.35\u2013$10.50B (\u22125% to \u22127%) and diluted EPS of $9.48\u2013$9.73",
    "The Americas (the majority of sales) has decelerated to negative comparable sales, while tariffs and higher promotions pressure gross margin and SG&A deleverages on lower volume",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=5)
add_para(tf, "Where the business stands (FY2025, ended Feb 1, 2026)", 14.5, CARD, bold=True, space_after=5)
for t in [
    "Net revenue $11,102.6M (+4.9% y/y); gross profit $6,284.1M (56.6% margin); operating income $2,210.6M (19.9% margin)",
    "Net income $1,579.2M; diluted EPS $13.26; diluted shares 119.1M (down from 128.0M in FY2022)",
    "Balance sheet: $1,807.2M cash and no funded debt \u2014 a net-cash position that funds buybacks and international expansion",
    "Cash from operations $1,602.5M; capital expenditures $680.8M \u2192 ~$0.9B free cash flow even in a decelerating year",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=5)

# =====================================================================
# 5. MACRO × MICRO OVERLAP (Firecrawl-sourced context)
# =====================================================================
s = slide_base("Macro \u00d7 Micro Overlap", "A discretionary macro air pocket collided with an Americas-specific slowdown \u2014 the stock prices both at once", page=pg(),
               sources="Source: FY2025 10-K (SEC EDGAR); Q2 FY2026 release (Sep 3, 2026); FRED GDPC1 (Aug 2026)")
tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(6.1), Inches(5.9))
add_para(tf, "Macro (external)", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "Real GDP still grows ~2.1% y/y (FRED GDPC1 Q2\u201926 vs Q2\u201925) \u2014 not recession, but selective spending away from premium discretionary",
    "Inflation and economic uncertainty weighed on Americas store traffic (10-K) \u2014 macro consumer caution shows up in micro footfall",
    "Tariff regime: Q2 recognized $134.5M IEEPA refunds (+560 bps GM) while new tariff layers remain an overhang on sourcing costs",
    "Athleisure competition intensifying at macro category level (Nike/adidas incumbents + Alo/Vuori/On challengers)",
]:
    add_para(tf, t, 12.8, INK, bullet=True, space_after=5)
add_para(tf, "Micro (company-specific)", 14.5, CARD, bold=True, space_after=5)
for t in [
    "Americas comps: \u22123% FY2025 \u2192 \u221212% in Q2 FY2026; China +20% FY2025 but \u22126% in Q2 as the base scales",
    "811 stores (+44 net FY2025) at ~$1,426/sq ft \u2014 still opening doors while mature-region comps turn negative",
    "FY2026 guide: first annual revenue decline (\u22125% to \u22127%) and EPS $9.48\u2013$9.73; Q3 guide \u221210% to \u221211%",
    "Clean run-rate EBIT margin ~13.2% (Q2 18.8% OM minus 560 bps tariff boost) \u2014 investors who miss this overstate sustainable margin",
]:
    add_para(tf, t, 12.8, INK, bullet=True, space_after=5)
box = rect(s, Inches(6.85), Inches(1.15), Inches(6.0), Inches(5.5), fill=NAVY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "OVERLAP \u2014 WHY $100", 12.5, GOLD, bold=True, first=True, space_after=8)
for t in [
    "Macro weakness maps onto LULU\u2019s largest profit pool: Americas ~71% of FY25 revenue (10-K: 70.7%)",
    "Tariff is simultaneously macro policy and micro P&L \u2014 Q2 optics mask run-rate margin",
    "The de-rating blends macro fear with a micro guide miss; both are largely in the price at ~10x FY26E EPS",
    "Our edge: underwrite trough economics (13.2% \u2192 15.5% OM), not a return to 30x earnings",
    "International micro growth (China/RoW) is the bridge \u2014 but needs macro stability, not a free option",
]:
    add_para(btf, t, 12.5, WHITE, bullet=True, space_after=6)

# =====================================================================
# 6. MARKET NARRATIVE
# =====================================================================
s = slide_base("Market Narrative", "Sentiment has capitulated \u2014 the sell-side is cutting targets into the print", page=pg(),
               sources="Source: sell-side research notes (Sep 2026); company guidance")
tb, tf = body_box(s)
add_para(tf, "What the street is saying", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "Consensus rating has drifted to \u201cHold / Reduce\u201d with a rapidly falling target; recent cuts include Goldman Sachs and JPMorgan to $95, Morgan Stanley $93 (Underweight), Wells Fargo $95, and Bank of America to $122",
    "Bear case: Americas saturation and negative comps, tariff-driven gross-margin pressure, SG&A deleverage, and share loss to emerging brands (Alo, Vuori, On)",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=5)
add_para(tf, "Our differentiated view", 14.5, CARD, bold=True, space_after=5)
for t in [
    "The tape has moved from euphoria (30\u201340x earnings in 2023\u201324) to capitulation (~10x today) \u2014 expectations are now low enough that even a conservative recovery re-rates the stock",
    "At ~3.5x EV/EBITDA the market implicitly assumes revenue and margins decline in perpetuity; our base case only needs stabilization, not a return to peak growth",
    "Consensus is extrapolating one weak fiscal year; we underwrite the international runway and structural margins the street is discounting to zero",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=5)

# =====================================================================
# 6. COMPANY OVERVIEW
# =====================================================================
s = slide_base("Company Overview", "A vertically integrated, direct-to-consumer technical apparel brand", page=pg())
tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(7.2), Inches(5.9))
add_para(tf, "Business", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "Founded 1998 in Vancouver; designs technical athletic apparel and accessories for yoga, training, running and everyday wear",
    "Sells through company-operated stores, a direct e-commerce channel, and select wholesale \u2014 a predominantly DTC model that captures full retail economics and first-party data",
    "Reports revenue by geography: Americas (majority of sales), China Mainland (fastest-growing), and Rest of World",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)
add_para(tf, "Revenue trajectory (US$ M)", 14.5, CARD, bold=True, space_after=5)
for y in ["FY2022", "FY2023", "FY2024", "FY2025"]:
    add_para(tf, f"{y}:  {D.IS['revenue'][y]:,}   \u2192  op. margin {D.IS['operating_income'][y]/D.IS['revenue'][y]*100:.1f}%",
             13, INK, bullet=True, space_after=4)
# right: segment mix note box
box = rect(s, Inches(8.0), Inches(1.15), Inches(4.85), Inches(5.5), fill=LGREY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "GROWTH ARCHITECTURE", 12.5, CARD, bold=True, first=True, space_after=6)
add_para(btf, "Americas", 13.5, NAVY, bold=True, space_after=1)
add_para(btf, "Mature, high-productivity base; now negative comps \u2014 the source of the market's concern", 12, INK, space_after=8)
add_para(btf, "China Mainland", 13.5, NAVY, bold=True, space_after=1)
add_para(btf, "Fastest-growing region (~20%+); large store and brand-awareness runway", 12, INK, space_after=8)
add_para(btf, "Rest of World", 13.5, NAVY, bold=True, space_after=1)
add_para(btf, "Early-stage in Europe and APAC; men's and international whitespace underpenetrated", 12, INK, space_after=0)

# =====================================================================
# 7. BUSINESS MODEL
# =====================================================================
s = slide_base("Business Model & Unit Economics", "Premium pricing plus DTC scale drives sector-leading margins", page=pg())
tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(7.2), Inches(5.9))
add_para(tf, "Unit economics", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "Gross margin ~57\u201359% \u2014 premium, full-price selling with historically limited promotion",
    "Operating margin ~20% (peaked ~24% in FY2024) \u2014 among the best in branded apparel",
    "Asset-efficient store fleet: leased footprint (right-of-use assets ~$1.6B) with high sales productivity per square foot",
    "Cash-generative: FY2025 CFO $1,602.5M on $680.8M capex \u2192 ~$0.9B free cash flow",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)
add_para(tf, "Competitive moats", 14.5, CARD, bold=True, space_after=5)
for t in [
    "Brand and community (ambassadors, in-store events) supporting pricing power",
    "Product innovation and proprietary fabrics driving repeat purchase",
    "First-party DTC data enabling assortment, pricing and inventory discipline",
    "Scale in sourcing and a growing men's and international opportunity",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)
box = rect(s, Inches(8.0), Inches(1.15), Inches(4.85), Inches(3.0), fill=NAVY)
btf = box.text_frame; btf.word_wrap = True
add_para(btf, "MARGIN PROFILE (FY2025)", 12.5, GOLD, bold=True, first=True, space_after=8)
add_para(btf, "Gross margin        56.6%", 15, WHITE, space_after=6)
add_para(btf, "Operating margin   19.9%", 15, WHITE, space_after=6)
add_para(btf, "Net margin            14.2%", 15, WHITE, space_after=6)
add_para(btf, "FCF (approx.)         ~$0.9B", 15, WHITE, space_after=0)

# =====================================================================
# 8. INDUSTRY OVERVIEW
# =====================================================================
s = slide_base("Industry Overview", "Structural tailwinds favor scaled brands in a consolidating category", page=pg(),
               sources="Source: company filings; industry estimates (analyst)")
tb, tf = body_box(s)
add_para(tf, "Category dynamics", 14.5, CARD, bold=True, first=True, space_after=5)
for t in [
    "Global activewear / athleisure is a large (~$400B) market growing at a mid-single-digit rate, supported by health, wellness and the ongoing casualization of apparel",
    "The category is fragmented but consolidating toward brands with scale, technical product and omnichannel reach; Nike and adidas are incumbents, with Vuori, Alo and On emerging",
    "Barriers to entry are rising: brand equity, technical fabric development, supply-chain scale and a productive retail footprint are difficult to replicate quickly",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)
add_para(tf, "Where LULU fits", 14.5, CARD, bold=True, space_after=5)
for t in [
    "A premium, category-defining brand in technical apparel with structurally higher margins than most peers",
    "Under-indexed internationally versus Nike/adidas \u2014 the primary multi-year growth lever",
    "Competition is real but LULU's scale, innovation cadence and community remain differentiated",
]:
    add_para(tf, t, 13, INK, bullet=True, space_after=6)

# =====================================================================
# 9-11. THESIS I / II / III
# =====================================================================
def thesis_slide(num, title, desc, headline, bullets, metric_title, metric_lines):
    s = slide_base(f"Investment Thesis {num}", desc, page=pg())
    tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(7.4), Inches(5.9))
    add_para(tf, headline, 15, CARD, bold=True, first=True, space_after=7)
    for t in bullets:
        add_para(tf, t, 13.3, INK, bullet=True, space_after=7)
    box = rect(s, Inches(8.15), Inches(1.15), Inches(4.7), Inches(len(metric_lines)*0.55 + 0.9), fill=LGREY)
    btf = box.text_frame; btf.word_wrap = True
    add_para(btf, metric_title, 12.5, CARD, bold=True, first=True, space_after=8)
    for ml in metric_lines:
        add_para(btf, ml, 13.5, NAVY, bold=True, space_after=6)
    return s

thesis_slide(
    "I", "Priced for terminal decline",
    "The valuation already discounts a permanently shrinking business",
    "The market is paying trough multiples for a brand that still earned 19.9% operating margin in FY2025",
    [
        "At ~$100 the stock trades at ~3.5x EV/EBITDA and ~10.4x FY2026E EPS \u2014 versus a 5-year history of ~20\u201330x earnings",
        "The balance sheet holds $1.8B of cash and no funded debt, so nearly the entire enterprise value is covered by the operating business at a very low multiple",
        f"Our base-case DCF (WACC {_pct(V['wacc'])}, terminal growth 2.25%, terminal EBIT margin 15.5%) is in the model \u2014 bear is ~{_d(V['bear'])}; downside is cushioned by net cash and FCF yield, not because bear brackets $100",
        "To justify $100 you must assume revenue and margins fall in perpetuity; that is inconsistent with international growth and the FY2026 tariff-refund tailwind",
    ],
    "VALUATION SNAPSHOT",
    ["EV / EBITDA:  3.5x", "FY2026E P/E:  10.4x", "FCF yield:  ~8\u20139%", "Net cash:  $1.8B", f"DCF base:  {_d(V['base_dcf'])}"],
)

thesis_slide(
    "II", "International growth engine",
    "China and Rest-of-World offset a maturing Americas and restore growth",
    "Geographic mix shift is the bridge back to low-single-digit revenue growth (~2.3% FY27\u201330)",
    [
        "China Mainland has been compounding at ~20%+ with a long runway in store count and brand awareness relative to the Americas base",
        "Rest of World (Europe, APAC) is early-stage and under-penetrated versus global peers \u2014 an incremental multi-year contributor",
        "As international scales, it dilutes the drag from negative Americas comps; our model returns total revenue to ~2.3% growth (Street 3Y forecast 2.26%) after the FY2026 reset",
        "Men's and the digital channel add further optionality that the market is not paying for today",
    ],
    "REVENUE PATH (model)",
    ["FY2025A:  $11,102.6M", f"FY2026E:  ${IS_BASE['FY2026E']['revenue']:,}M (\u22126.1%)", f"FY2028E:  ${IS_BASE['FY2028E']['revenue']:,}M", f"FY2030E:  ${IS_BASE['FY2030E']['revenue']:,}M"],
)

thesis_slide(
    "III", "Elite economics & capital return",
    "Best-in-class margins and buybacks compound value through the trough",
    "Even a down year throws off ~$1B of FCF that is being returned aggressively",
    [
        "Operating margin remains ~20% in FY2025 \u2014 far above most branded-apparel peers \u2014 with room to recover as tariffs and promotions normalize",
        "FY2025 cash from operations was $1,602.5M; after ~$681M capex the business generated roughly $0.9B of free cash flow",
        "The company repurchased $1.6B (FY2024) and $1.2B (FY2025) of stock; share count has fallen from 128.0M (FY2022) to 119.1M (FY2025)",
        "At ~$100 per share, every dollar of buyback retires far more shares than at prior highs \u2014 highly accretive to per-share value",
    ],
    "CAPITAL RETURN",
    ["FY2025 CFO:  $1,602.5M", "FY2025 capex:  $680.8M", "FY2025 buyback:  $1,178.3M", "Shares:  128.0M \u2192 119.1M"],
)

# =====================================================================
# 12. RISKS & MITIGANTS
# =====================================================================
s = slide_base("Risks & Mitigants", "The bear points are real but largely discounted at today's multiple", page=pg())
risks = [
    ("Brand fatigue / fashion risk in a core-heavy assortment",
     "Deep product pipeline, community engagement and international whitespace diversify demand"),
    ("Americas comparable sales stay negative",
     "Expectations are already low; ~10x earnings provides a valuation cushion if declines moderate"),
    ("Tariffs compress gross margin",
     "Pricing power, sourcing diversification, and $0.86/share of FY2026 tariff refunds offset part of the hit"),
    ("Competition from Alo, Vuori, On and incumbents",
     "Scale, fabric innovation, and under-penetrated men's / international segments defend share"),
    ("FX translation on a growing international mix",
     "Natural operational hedges and an active hedging program limit earnings volatility"),
]
tb, tf = textbox(s, Inches(0.5), Inches(1.15), Inches(12.35), Inches(0.3))
add_para(tf, "Risk  \u2192  Mitigant", 14, CARD, bold=True, first=True, space_after=0)
top = 1.55
for rk, mg in risks:
    b = rect(s, Inches(0.5), Inches(top), Inches(6.0), Inches(0.95), fill=LGREY)
    bt = b.text_frame; bt.word_wrap = True; bt.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt, rk, 12.5, CARD, bold=True, first=True, space_after=0)
    b2 = rect(s, Inches(6.7), Inches(top), Inches(6.15), Inches(0.95), fill=WHITE, line=GREY)
    bt2 = b2.text_frame; bt2.word_wrap = True; bt2.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt2, mg, 12.5, INK, first=True, space_after=0)
    top += 1.06

# =====================================================================
# 13. CATALYST TIMELINE
# =====================================================================
s = slide_base("Catalyst Timeline", "A sequence of events that closes the gap to intrinsic value", page=pg())
cats = [
    ("Q3 FY2026 (Dec 2026)", "Company guided revenue \u221210% to \u221211%; watch for stabilization signals and holiday traffic / China 11.11 read-through"),
    ("FY2026 year-end (early 2027)", "First full year lapping the reset; tariff refunds and cost actions support EPS versus lowered expectations"),
    ("FY2027", "Margin inflection as promotions/tariffs anniversary; continued accretive buybacks compound per-share value"),
    ("FY2027\u2013FY2028", "Total revenue returns to growth as international scales \u2014 the trigger for multiple re-rating toward peers"),
]
top = 1.5
for i, (when, what) in enumerate(cats):
    dot = rect(s, Inches(0.6), Inches(top+0.05), Inches(0.22), Inches(0.22), fill=CARD)
    b = rect(s, Inches(1.1), Inches(top), Inches(2.9), Inches(1.05), fill=NAVY)
    bt = b.text_frame; bt.word_wrap = True; bt.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt, when, 13, GOLD, bold=True, first=True, space_after=0)
    b2 = rect(s, Inches(4.2), Inches(top), Inches(8.6), Inches(1.05), fill=LGREY)
    bt2 = b2.text_frame; bt2.word_wrap = True; bt2.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_para(bt2, what, 13, INK, first=True, space_after=0)
    top += 1.25

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
# 13–15. FINANCIALS (historicals + base & bull, all five forecast years)
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
s = slide_base("Valuation Summary", "Multiple methods converge above the current price \u2014 target $140", page=pg(),
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
cp_x = xpos(100); pt_x = xpos(140)
ln = rect(s, Inches(cp_x), Inches(1.55), Pt(2), Inches(3.4), fill=INK)
ln2 = rect(s, Inches(pt_x), Inches(1.55), Pt(2), Inches(3.4), fill=CARD)
tb, tf = textbox(s, Inches(cp_x-0.7), Inches(4.95), Inches(1.6), Inches(0.3))
add_para(tf, "Current $100", 10.5, INK, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
tb, tf = textbox(s, Inches(pt_x-0.7), Inches(5.2), Inches(1.6), Inches(0.3))
add_para(tf, "Target $140", 10.5, CARD, bold=True, align=PP_ALIGN.CENTER, first=True, space_after=0)
tb, tf = textbox(s, Inches(0.5), Inches(5.7), Inches(12.35), Inches(1.2))
add_para(tf, f"We set a 12-month target of $140 \u2014 above the Gordon DCF of {_d(V['base_dcf'])}, for a partial re-rating toward historical multiples \u2014 implying ~40% upside", 13, NAVY, bold=True, first=True, space_after=5)
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
