#!/usr/bin/env python3
"""Generate the GIS Investment Research pitch deck template.

Produces a fully laid-out example deck (Summit Outdoor Co.) with charts, tables,
KPI boxes, football field, DCF sensitivity, and comps — not generic placeholders.
Replace company name, ticker, and numbers; keep the slide structure.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from gis_pitch import PitchDeck, STANDARD_TOC, add_para, export_pdf, CARD, GREY, INK, GREEN, NAVY
from slide_layouts import (
    thesis_summary,
    bullet_sections,
    company_overview,
    business_model_slide,
    industry_slide,
    thesis_slide,
    risks_mitigants,
    catalyst_timeline,
    financials_slide,
    wacc_slide,
    football_field,
    dcf_slide,
    comps_slide,
    scenario_columns,
)
import template_data as T

OUT = os.path.join(os.path.dirname(__file__), "GIS_Investment_Pitch_Template.pptx")


def m(v):
    return f"{v:,.0f}"


def main():
    deck = PitchDeck(company_header=T.HEADER)
    deck.default_source = f"Source: company SEC filings (Form 10-K, CIK {T.CIK})"

    deck.title_slide(
        company_name=T.COMPANY_INC,
        ticker_line=f"{T.EXCHANGE}: {T.TICK}",
        recommendation=T.RECOMMENDATION,
        current_price=f"${T.PRICE:.2f}",
        target_price=f"${T.TARGET:.0f}",
        upside_pct=T.UPSIDE + " upside",
        descriptor=T.DESCRIPTOR,
        date_line="September 2026        GIS template — replace Summit / SUMM with your company",
    )

    deck.toc_slide(STANDARD_TOC)

    thesis_summary(
        deck,
        f"Overweight {T.TICK} with a ${T.TARGET:.0f} target on {T.UPSIDE} upside and an asymmetric payoff",
        f"${T.TARGET:.0f}",
        f"{T.UPSIDE} vs ${T.PRICE:.2f} today",
        "OVERWEIGHT",
        "Shares −38% from highs after FY2026 guide reset; sentiment overshoots fundamentals at ~8x forward EPS",
        [
            "Priced for permanent decline — at ~5.2x EV/EBITDA and ~8.4x FY2026E EPS with $445M net cash, the stock embeds impairment the P&L does not support",
            "International whitespace — EMEA and APAC are <35% of sales vs ~50% at global peers; store rollout can restore mid-single-digit growth post-FY2026",
            "Durable FCF + buybacks — 14% operating margin in FY2025, ~$270M FCF, and shares down 9% over three years of repurchases",
        ],
        f"Base-case DCF ${T.DCF_BASE}; bear ${T.DCF_BEAR} (−19%) vs bull ${T.DCF_BULL} (+86%) — downside cushioned by net cash and ~7% FCF yield",
    )

    bullet_sections(
        deck,
        "Situation Overview",
        "A profitable DTC compounder has de-rated to value multiples after a guidance reset",
        [
            ("What happened", [
                f"{T.TICK} shares fell from a 52-week high of ~$68 to ~${T.PRICE:.0f} after Q2 FY2026 earnings (July 2026)",
                "Management guided FY2026 revenue $3.28–$3.34B (−4% to −6%) and diluted EPS $2.30–$2.45 vs $2.68 in FY2025",
                "Americas comps turned negative (−4% FY2025) while higher promo and freight pressure gross margin; SG&A deleverages on lower volume",
            ]),
            ("Where the business stands (FY2025)", [
                f"Net revenue ${T.REV[-1]:,}M (+3.9% y/y); gross profit ${T.GP[-1]:,}M (56.6% margin); operating income ${T.OI[-1]:,}M (14.0% margin)",
                f"Net income ${T.NI[-1]:,}M; diluted EPS ${T.EPS[-1]:.2f}; diluted shares 118.6M (down from 128.2M in FY2022)",
                f"Balance sheet: ${T.CASH[-1]:,}M cash, $285M term loan — net cash ~$160M funds buybacks and international stores",
                f"CFO ${T.CFO[-1]:,}M; capex ${T.CAPEX[-1]:,}M → ~${T.FCF[-1]:,}M free cash flow in a slowing year",
            ]),
        ],
        sources="Source: company 10-K (FY2025) and Q2 FY2026 earnings release",
    )

    bullet_sections(
        deck,
        "Market Narrative",
        "Sentiment has capitulated — sell-side targets cluster near the current price",
        [
            ("What the street is saying", [
                "Consensus drifted to Hold; recent cuts include Morgan Stanley $40 (EW), Barclays $41, and UBS $43",
                "Bear case: Americas saturation, promo-driven GM pressure, private-label share gain, and execution risk on international rollout",
            ]),
            ("Our differentiated view", [
                "Tape moved from ~22x earnings (FY2024) to ~8x today — expectations low enough that stabilization re-rates the stock",
                f"At ~5x EV/EBITDA the market implies perpetual decline; our base case needs only flat-to-low-single-digit growth post-FY2026",
                "Consensus extrapolates one reset year; we underwrite international runway and normalized 15% operating margins",
            ]),
        ],
        sources="Source: sell-side research (Aug 2026); company guidance",
    )

    company_overview(
        deck,
        "A vertically integrated, direct-to-consumer outdoor apparel brand",
        [
            "Founded 2004 in Denver; designs technical outerwear, trail apparel, and accessories for hiking, ski, and everyday outdoor",
            "Predominantly DTC — 412 company-operated stores plus e-commerce (~48% of revenue) and limited wholesale",
            "Reports Americas (~68%), EMEA (~18%), and APAC (~14%) — international under-indexed vs global outdoor peers",
        ],
        [f"{y}:  {rev:,}   →  op. margin {om:.1f}%" for y, rev, om in zip(T.HY, T.REV, T.OM)],
        [
            ("Americas", "Mature, high-productivity base; negative comps are the market concern"),
            ("EMEA", "Fastest store growth (+22% revenue FY2025); brand awareness still early"),
            ("APAC", "Small but accelerating — distributor transition to owned stores in FY2026–27"),
        ],
        T.REV_CHART_LABELS,
        T.REV_CHART_VALUES,
    )

    business_model_slide(
        deck,
        "Premium pricing plus DTC scale drives sector-leading unit economics",
        [
            "Gross margin ~56–57% — full-price positioning with disciplined promo (FY2025 GM 56.6%)",
            "Operating margin ~14% FY2025 (peak 17.0% FY2024) — top quartile in specialty apparel",
            "Store productivity ~$1,180/sq ft on ~3,200 sq ft average — above mall apparel median",
            f"Cash-generative: FY2025 CFO ${T.CFO[-1]:,}M on ${T.CAPEX[-1]:,}M capex → ~${T.FCF[-1]:,}M FCF",
        ],
        [
            "Brand community and athlete ambassadors support pricing power",
            "Technical fabric R&D (Gore-Tex, proprietary insulation) drives repeat purchase",
            "First-party DTC data informs assortment, pricing, and inventory",
            "Scale sourcing + international whitespace",
        ],
        [
            "Gross margin        56.6%",
            "Operating margin   14.0%",
            "Net margin            9.1%",
            f"FCF (approx.)         ~${T.FCF[-1]/1000:.2f}B",
        ],
        T.OM_CHART_LABELS,
        T.OM_CHART_VALUES,
    )

    industry_slide(
        deck,
        "Structural tailwinds favor scaled brands in a consolidating outdoor category",
        [
            "Global outdoor / performance apparel ~$180B, growing ~6% — wellness, adventure travel, and gorpcore casualization",
            "Fragmented market consolidating toward technical product, omnichannel, and brand equity; Patagonia, Arc'teryx, and Columbia are anchors",
            "Barriers rising: fabric partnerships, supply-chain scale, and productive retail footprints are hard to replicate",
        ],
        [
            "Premium technical positioning with structurally higher margins than mass outdoor",
            "Under-indexed internationally — primary multi-year growth lever vs Patagonia / Columbia",
            "Competition real but Summit's DTC scale and innovation cadence remain differentiated",
        ],
        T.GEO_MIX,
    )

    thesis_slide(
        deck, "I",
        "The valuation already discounts a permanently shrinking business",
        "The market pays trough multiples for a brand that still earned 14% operating margin in FY2025",
        [
            f"At ~${T.PRICE:.0f} the stock trades at ~5.2x EV/EBITDA and ~8.4x FY2026E EPS vs a 5-year range of ~15–25x",
            f"Net cash ~$160M after term debt; enterprise value ~${T.PRICE * 118.6 / 1000:.1f}B at current price",
            f"Base-case DCF ${T.DCF_BASE}/sh (WACC 10.2%, g 2.25%, terminal OM 15%) — bear ${T.DCF_BEAR} brackets downside",
            "To justify current price you need revenue and margins to shrink forever — inconsistent with international growth",
        ],
        "VALUATION SNAPSHOT",
        ["EV / EBITDA:  5.2x", "FY2026E P/E:  8.4x", "FCF yield:  ~7%", "Net cash:  ~$160M", f"DCF base:  ${T.DCF_BASE}"],
    )

    thesis_slide(
        deck, "II",
        "International expansion is the multi-year growth engine",
        "Geographic mix shift is the bridge back to mid-single-digit revenue growth",
        [
            "EMEA revenue +22% FY2025 with a long store-count runway vs a mature Americas base",
            "APAC transitioning from distributors to owned retail — incremental FY2026–28 contributor",
            "As international scales, it dilutes Americas drag; model returns total revenue to +4.2% by FY2028E",
            "Women's technical and footwear add optionality the market is not paying for",
        ],
        "REVENUE PATH (model)",
        [f"FY2025A:  ${T.REV[-1]:,}M", f"FY2026E:  ${T.FC_REV[0]:,}M (−5.8%)", f"FY2028E:  ${T.FC_REV[1]:,}M", f"FY2030E:  ${T.FC_REV[2]:,}M"],
    )

    thesis_slide(
        deck, "III",
        "Elite economics and buybacks compound value through the trough",
        "Even a reset year throws off ~$250M+ of FCF returned via repurchases",
        [
            "Operating margin 14.0% FY2025 — well above specialty retail median ~8%",
            f"FY2025 CFO ${T.CFO[-1]:,}M; after ${T.CAPEX[-1]:,}M capex → ~${T.FCF[-1]:,}M FCF",
            f"Repurchased ${T.BUYBACKS[-2]:,}M (FY2024) and ${T.BUYBACKS[-1]:,}M (FY2025); shares 128.2M → 118.6M",
            f"At ~${T.PRICE:.0f}/sh, buybacks retire ~2× the shares vs prior highs — accretive to per-share value",
        ],
        "CAPITAL RETURN",
        [f"FY2025 CFO:  ${T.CFO[-1]:,}M", f"FY2025 capex:  ${T.CAPEX[-1]:,}M", f"FY2025 buyback:  ${T.BUYBACKS[-1]:,}M", "Shares:  128.2M → 118.6M"],
    )

    risks_mitigants(
        deck,
        "The bear points are real but largely discounted at today's multiple",
        [
            ("Weather / fashion risk in core outerwear", "Broadened assortment, year-round technical layers, and international diversification"),
            ("Americas comps stay negative", "Expectations reset; ~8x EPS provides cushion if declines moderate"),
            ("Freight and input-cost inflation", "Pricing power, sourcing mix shift, and hedging program"),
            ("Competition from Patagonia, Arc'teryx, private label", "DTC data, fabric partnerships, under-penetrated EMEA/APAC"),
            ("FX on growing international mix", "Natural hedges and active FX program limit earnings volatility"),
        ],
    )

    catalyst_timeline(
        deck,
        "A sequence of events that closes the gap to intrinsic value",
        [
            ("Q3 FY2026 (Nov 2026)", "Guided revenue −6% to −7%; watch holiday traffic and EMEA store openings for stabilization"),
            ("FY2026 year-end (Mar 2027)", "First full year lapping the reset; cost actions vs lowered bar support EPS"),
            ("FY2027", "Margin inflection as promo/freight anniversaries; buybacks compound per-share value"),
            ("FY2027–FY2028", "Total revenue returns to growth — trigger for multiple re-rating toward peers"),
        ],
    )

    hdr = ["US$ M"] + T.HY + T.FC
    financials_slide(
        deck, "Financials — Income Statement",
        "Historical results with our base-case forecast (US$ M)",
        hdr,
        [
            ["Net revenue"] + [m(x) for x in T.REV] + [m(x) for x in T.FC_REV],
            ["Gross profit"] + [m(x) for x in T.GP] + ["1,848", "2,009", "2,247"],
            ["Operating income"] + [m(x) for x in T.OI] + [m(x) for x in T.FC_OI],
            ["Operating margin %"] + [f"{x:.1f}%" for x in T.OM] + T.FC_OM,
            ["Net income"] + [m(x) for x in T.NI] + [m(x) for x in T.FC_NI],
            ["Diluted EPS ($)"] + [f"{x:.2f}" for x in T.EPS] + T.FC_EPS,
        ],
        "FY2026E aligns with company guidance (revenue −5.8%, EPS ~$2.35); recovery driven by international growth and margin normalization",
        bold_rows=(0, 2, 5),
    )

    financials_slide(
        deck, "Financials — Balance Sheet",
        "A net-cash balance sheet underpins downside protection (US$ M)",
        hdr,
        [
            ["Cash & equivalents"] + [m(x) for x in T.CASH] + ["512", "628", "845"],
            ["Inventories"] + [m(x) for x in T.INV] + ["342", "335", "358"],
            ["Total assets"] + [m(x) for x in T.TA] + ["2,485", "2,712", "3,105"],
            ["Total liabilities"] + [m(x) for x in T.TL] + ["945", "982", "1,045"],
            ["Total equity"] + [m(x) for x in T.TE] + ["1,540", "1,730", "2,060"],
            ["Funded debt"] + ["320", "305", "285", "285"] + ["265", "220", "180"],
        ],
        "De-leveraging path and growing cash fund buybacks and international expansion — core of our margin of safety",
        bold_rows=(2, 4, 5),
    )

    financials_slide(
        deck, "Financials — Cash Flow",
        "Durable free cash flow through the trough (US$ M)",
        hdr,
        [
            ["Cash from operations"] + [m(x) for x in T.CFO] + ["365", "398", "445"],
            ["Capital expenditures"] + [f"({m(x)})" for x in T.CAPEX] + ["(132)", "(118)", "(112)"],
            ["Free cash flow"] + [m(x) for x in T.FCF] + ["233", "280", "333"],
            ["Share repurchases"] + [f"({m(x)})" for x in T.BUYBACKS] + ["(125)", "(125)", "(125)"],
            ["D&A"] + ["142", "158", "172", "185"] + ["192", "205", "218"],
        ],
        "FCF stays ~$230M+ even in the FY2026 reset year, comfortably funding continued repurchases",
        bold_rows=(2,),
    )

    wacc_slide(
        deck,
        "A modestly levered, net-cash-lean structure drives a ~10.2% discount rate",
        [
            "$285M term loan (3.8% fixed) plus undrawn revolver; net cash ~$160M after debt",
            f"Enterprise value ~${T.PRICE * 118.6 / 1000:.1f}B at ~${T.PRICE:.0f}/sh",
            "Relevered beta 1.05 on specialty retail comp set; Damodaran unlevered retail β cross-check",
            f"Capital returned via buybacks; FY2025 repurchases ${T.BUYBACKS[-1]:,}M (no dividend)",
        ],
        [
            ("Risk-free rate (10-yr UST)", "4.8%"),
            ("Equity risk premium", "6.0%"),
            ("Levered beta", "1.05"),
            ("Cost of equity", "11.1%"),
            ("After-tax cost of debt", "2.9%"),
            ("Equity / debt weight", "92% / 8%"),
        ],
        "WACC = 10.2%",
    )

    football_field(
        deck,
        f"Multiple methods converge above the current price — target ${T.TARGET:.0f}",
        [
            ("P / E (9–14x FY2026E)", 32, 52),
            ("EV / EBITDA (4.5–7.0x)", 38, 58),
            ("DCF (bear – bull)", T.DCF_BEAR, T.DCF_BULL),
            ("52-week range", 38, 68),
        ],
        T.PRICE,
        T.TARGET,
        vmin=25,
        vmax=85,
        summary=f"We set a 12-month target of ${T.TARGET:.0f} — above base DCF ${T.DCF_BASE}, allowing execution risk — implying {T.UPSIDE} upside",
        subtext="The P/E low end (~$32) sits below today's price, showing how little recovery is required for the stock to work",
    )

    dcf_slide(
        deck,
        f"Base-case unlevered DCF yields ~${T.DCF_BASE} per share",
        [
            ["Assumption", "Value"],
            ["Revenue growth (FY26 → FY30)", "−5.8% → +4.2%"],
            ["EBIT margin (FY26 → FY30)", "12.5% → 15.0%"],
            ["Tax rate", "24%"],
            ["Capex % of revenue", "4.0%"],
            ["WACC", "10.2%"],
            ["Terminal growth", "2.25%"],
        ],
        [
            ("PV of explicit FCF (FY26–FY30)", "892"),
            ("PV of terminal value", "2,148"),
            ("Enterprise value", "3,040"),
            ("Plus: net cash", "160"),
            ("Equity value", "3,200"),
            ("÷ Diluted shares (M)", "118.6"),
        ],
        f"Implied value:  ${T.DCF_BASE} / share",
        [
            ["WACC \\ g", "1.5%", "2.0%", "2.25%", "2.5%", "3.0%"],
            ["9.2%", "$54", "$56", "$58", "$60", "$64"],
            ["10.2%", "$48", "$50", f"${T.DCF_BASE}", "$53", "$56"],
            ["11.2%", "$43", "$45", "$46", "$48", "$50"],
        ],
    )

    comps_slide(
        deck,
        f"{T.TICK} screens cheap on absolute and relative multiples",
        ["Company", "EV/EBITDA", "P/E (fwd)", "Rev growth", "Op margin"],
        [
            [f"{T.COMPANY} ({T.TICK})", "5.2x", "8.4x", "−6% (FY26E)", "~14%"],
            ["Columbia (COLM)", "9.8x", "16.2x", "low-single", "~12%"],
            ["Deckers (DECK)", "7.9x", "12.1x", "mid-teens", "~22%"],
            ["VF Corp (VFC)", "10.1x", "14.0x", "flat/decl.", "~8%"],
            ["Canada Goose (GOOS)", "11.5x", "18.5x", "mid-single", "~18%"],
            ["Amer Sports (AS)", "14.2x", "22.0x", "high-single", "~15%"],
        ],
        [
            f"{T.TICK} trades at roughly half the EV/EBITDA of branded outdoor peers despite similar gross margins",
            "Discount reflects near-term Americas weakness, not structural inferiority",
            "Partial re-rating to 7x EV/EBITDA implies ~$55/sh — still below historical median",
        ],
    )

    scenario_columns(
        deck,
        "Asymmetric payoff: limited downside, substantial upside",
        [
            ("BEAR", f"${T.DCF_BEAR}", "−19%", CARD, [
                "FY2026 revenue −9%; growth stays flat FY28–30",
                "Terminal EBIT margin 11.5%",
                "WACC 11.2%; terminal growth 1.5%",
                "Americas decline persists; share loss continues",
            ]),
            ("BASE", f"${T.DCF_BASE}", "+24%", GREEN, [
                "FY2026 revenue −5.8%; returns to +4% by FY2028",
                "Terminal EBIT margin 15.0%",
                "WACC 10.2%; terminal growth 2.25%",
                "International offsets stabilizing Americas",
            ]),
            ("BULL", f"${T.DCF_BULL}", "+86%", NAVY, [
                "FY2026 revenue −3%; +7% avg FY27–30",
                "Terminal EBIT margin 16.5%",
                "WACC 9.2%; terminal growth 3.0%",
                "Margin recovery; brand re-accelerates",
            ]),
        ],
        "Probability-weighted value (25% / 50% / 25%) ≈ $56 — risk/reward skews to the upside",
    )

    s = deck.slide_base("Sources & Disclaimer", "Data provenance and standard research disclaimer", page=deck.pg())
    tb, tf = deck.body_box(s)
    add_para(tf, "Sources", 14, CARD, bold=True, first=True, space_after=5)
    for t in [
        f"Historical financials: {T.COMPANY} Form 10-K via SEC EDGAR (CIK {T.CIK}); FY2025 ended January 31, 2026",
        "Q2 FY2026 results and FY2026 guidance: company earnings release dated July 15, 2026",
        "Market data (price, shares, beta): public market sources as of September 2026",
        "Projections, DCF and comparable-company analysis: GIS Investment Research models — replace with your company data",
        "Summit Outdoor Co. is a fictional GIS template example; swap all figures for your assigned company",
    ]:
        add_para(tf, t, 12.5, INK, bullet=True, space_after=5)
    add_para(tf, "Disclaimer", 14, CARD, bold=True, space_after=5)
    add_para(
        tf,
        "This presentation is prepared for educational purposes as part of the Global Investment Society "
        "selection process and does not constitute investment advice or a recommendation to buy or sell any security",
        12, GREY, italic=True, space_after=0,
    )

    n = deck.save(OUT)
    print(f"Saved {os.path.abspath(OUT)} with {n} slides")
    try:
        pdf = export_pdf(OUT)
        print(f"Saved {pdf}")
    except Exception as e:
        print(f"PDF export skipped ({e})")


if __name__ == "__main__":
    main()
