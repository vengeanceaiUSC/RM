"""Humanized pitch-deck narrative copy (slides 2–13).

Body bullets carry no inline [n] tags; sources are footnoted per slide.
"""
from __future__ import annotations

MODEL = "unbeiesgbar_final.xlsx"

# (title, descriptor/subtitle, bullets, footnotes)
# Footnotes: list of (num, clean_source_line) — no raw URLs in deck body/footer.
NARRATIVE_SLIDES: list[tuple[str, str, list[str], list[tuple[int, str]]]] = [
    (
        "Situation Overview and Current Investment Setup",
        "Guidance reset compresses the multiple; operating base case still intact",
        [
            "Shares trade near $100 following the Q2 FY26 guidance revision (~$100.61 last sale).",
            "FY26 revenue guide: -5% to -7% vs. prior double-digit growth trajectory.",
            "Americas: ~flat comps and ~6 net store openings/yr (FY26–FY30 base); China: ~10.2% comps and ~16 openings/yr.",
            "Clean run-rate EBIT margin: 13.2% (FY25 10-K baseline).",
        ],
        [
            (1, "TIKR — post-earnings price action (Sep 2026)"),
            (2, "Lululemon Q2 FY26 earnings supplement"),
            (3, f"{MODEL}, Revenue Drivers & Scenarios (base case)"),
            (4, f"{MODEL}, Scenarios tab; FY2025 Form 10-K"),
        ],
    ),
    (
        "Market Narrative and Analyst Sentiment Surrounding the Stock",
        "Consensus de-rated; estimates lag revised FY26 outlook",
        [
            "Morgan Stanley cut estimates after FY26 revenue guide of ~$10.35B.",
            "Consensus target moved from ~$176 to ~$136 (Simply Wall St aggregate).",
            "BTIG maintained Neutral; near-term caution, limited change to long-run margin view.",
        ],
        [
            (1, "MarketBeat — Morgan Stanley estimate revision (Sep 2026)"),
            (2, "Simply Wall St — LULU consensus target history"),
            (3, "GuruFocus — BTIG rating reiteration"),
        ],
    ),
    (
        "Investment Thesis Summary and Target Price",
        "Base DCF $133.64; 12-month target $140 on modest multiple re-rating",
        [
            "Base-case DCF: $133.64/sh (+33.6% vs. ~$100 spot); 12-mo target $140.00.",
            "FY26–FY30 avg EBIT margin ~14.4%; path from 13.2% clean margin to 15.5% by FY30.",
            "Lease-adjusted WACC 9.0%; revenue CAGR 2.3% in base case.",
            "FY25 China segment +29% y/y vs. flat Americas; model offsets US softness via China comps/openings.",
        ],
        [
            (1, f"{MODEL}, DCF valuation summary (base case)"),
            (2, f"{MODEL}, Scenarios — EBIT margin path"),
            (3, f"{MODEL}, WACC & Revenue Drivers tabs"),
            (4, "Lululemon Q2 FY26 earnings release (segment growth)"),
        ],
    ),
    (
        "Geographic Segments and Revenue Breakdown",
        "China segment growth offsets Americas cyclical stagnation",
        [
            "FY25 revenue $11.1B; Americas $7.85B (70.7% of total).",
            "Americas comparable sales -3% (FY25 10-K); cyclical, not structural share loss.",
            "China Mainland $1.75B segment revenue; +20% FY25 comparable sales.",
        ],
        [
            (1, "FY2025 Form 10-K — segment revenue"),
            (2, "FY2025 Form 10-K — Americas comparable sales"),
            (3, f"{MODEL}, Revenue Drivers — China Mainland build"),
        ],
    ),
    (
        "Business Model Unit Economics and Competitive Moats",
        "DTC mix and store productivity support margin structure",
        [
            "Gross margin 56.6% (FY25 base case).",
            "E-commerce EBIT margin 23.6% (no retail occupancy load).",
            "Company-operated stores: $1,426/sq ft; 18.6% EBIT margin.",
        ],
        [
            (1, f"{MODEL}, Scenarios — gross margin"),
            (2, f"{MODEL}, Revenue Drivers — e-commerce margin"),
            (3, f"{MODEL}, Revenue Drivers — store productivity"),
        ],
    ),
    (
        "Industry Overview - Trends and Structure",
        "Fragmented athleisure; premium segment retains entry barriers",
        [
            "Global athleisure remains fragmented; scale and brand equity concentrated in leaders.",
            "Premium segment: higher barriers vs. mass athletic wholesale.",
            "LULU ROE 30.25% vs. ~6.81% apparel median (FinanceCharts).",
        ],
        [
            (1, "Market.us — athleisure market structure"),
            (2, "Fortune Business Insights — premium athleisure"),
            (3, "FinanceCharts — LULU ROE vs. retail median"),
        ],
    ),
    (
        "Industry Overview - Barriers to Entry and Profitability",
        "LULU gross margin premium vs. wholesale-heavy peers",
        [
            "LULU gross margin 56.6% vs. Nike ~43% and Under Armour ~46%.",
            "Wholesale mix structurally caps peer gross margin.",
            "DTC scale limits promotional margin erosion vs. mass-market peers.",
        ],
        [
            (1, f"{MODEL}, gross margin assumption"),
            (2, "Investing.com — NKE / UAA historical gross margin"),
            (3, "ProAnalyst — LULU competitive positioning"),
        ],
    ),
    (
        "Investment Thesis I",
        "Share price dislocation vs. normalized cash generation",
        [
            "2026 drawdown >50% from 52-wk high; spot ~$100 vs. base intrinsic ~$134.",
            "DTC mix supports ~54.9% adjusted gross margin in base case.",
            "2.25% terminal growth still implies >30% upside to $133.64 base value.",
        ],
        [
            (1, "Everything Money — 2026 price dislocation (context)"),
            (2, "Lululemon Q2 FY26 earnings release"),
            (3, f"{MODEL}, DCF terminal value & valuation summary"),
        ],
    ),
    (
        "Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor",
        "Recovery to 15.5% FY30 EBIT not required to reach base value",
        [
            "Peak EBIT margin 23.7% (FY24 10-K); not required for base case.",
            "Base case floors at 13.2% clean EBIT; brand retention supports trough margin.",
            "15.5% FY30 EBIT margin → $133.64/sh in base DCF.",
            "Full reversion to COVID-era peak margins not in base case.",
        ],
        [
            (1, f"{MODEL} & FY2024 Form 10-K — peak margins"),
            (2, f"{MODEL}, Scenarios — clean run-rate EBIT"),
            (3, f"{MODEL}, DCF valuation summary"),
            (4, f"{MODEL}, DCF — margin sensitivity"),
        ],
    ),
    (
        "Investment Thesis III: Geographic Growth Divergence",
        "Top-line bridge relies on China offsetting flat Americas comps",
        [
            "Americas: -4.0% comps FY26; terminal 2.0% by FY30 (base case).",
            "China: ~16 gross openings/yr vs. ~6 Americas; ~10.2% avg comps (14% → 7% FY26–FY30).",
            "China slowdown is primary risk to base-case revenue path.",
        ],
        [
            (1, f"{MODEL}, Revenue Drivers — Americas comps"),
            (2, f"{MODEL}, Revenue Drivers — China store & comp assumptions"),
            (3, f"{MODEL}, Scenarios — consolidated revenue"),
        ],
    ),
    (
        "Risk & Mitigants",
        "Bear floor $56; FCF and buybacks support per-share recovery",
        [
            "Bear-case DCF floor ~$56/sh; cumulative cash ~$45.64/sh (~45% of entry).",
            "CapEx fades 7.0% → 5.5% of revenue; 23.6% e-commerce EBIT supports FCF if China comps soften.",
            "$750M/yr repurchases → ~32.9M shares retired; base EPS ~$14.57 by FY30 schedule.",
        ],
        [
            (1, f"{MODEL}, Scenarios bear case; Barron's — guide context"),
            (2, f"{MODEL}, DCF capex & NWC; Q2 FY26 supplement"),
            (3, f"{MODEL}, DCF share repurchase schedule"),
        ],
    ),
]

TIMELINE_DESCRIPTOR = "Recovery path: guide trough → margin stabilization → re-rating"

TIMELINE_EVENTS: list[tuple[str, str]] = [
    (
        "2026 (Q3 FY2026 Trough)",
        "Q3 laps guidance trough; China Double 11 supports holiday traffic baseline.",
    ),
    (
        "2027 (FY2026 Year-End)",
        "First full-year post-reset; $134.5M IEEPA tariff refunds and SG&A actions support EPS.",
    ),
    (
        "2027 (FY2027 Margin Inflection)",
        "EBIT margin ~13.8% (+60 bps); $750M buybacks → EPS ~$10.09 (base schedule).",
    ),
    (
        "2028 (FY2027–FY2028 Multiple Re-Rating)",
        "China ~17 net stores FY27; +12%/+10% comps FY27–FY28 offset Americas -4% FY26 comp.",
    ),
]

TIMELINE_FOOTNOTES: list[tuple[int, str]] = [
    (1, f"{MODEL}, Scenarios & Revenue Drivers; Q2 FY26 supplement"),
    (2, f"{MODEL}, Scenarios & NOPAT reconciliation; Q2 FY26 supplement"),
    (3, f"{MODEL}, Scenarios & DCF FCF bridge"),
    (4, f"{MODEL}, Revenue Drivers & DCF summary; StockAnalysis consensus"),
]
