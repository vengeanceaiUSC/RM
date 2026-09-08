#!/usr/bin/env python3
"""Generate a blank GIS Investment Research pitch deck template.

Output: GIS_Investment_Pitch_Template.pptx (22 slides + sources = 22 content slides)
Matches the standard GIS IR selection structure. Fill placeholders or copy slide
chrome into a company-specific build_pitch.py (see LULU/scripts/build_pitch.py).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from gis_pitch import PitchDeck, STANDARD_TOC, add_para, CARD, GREY, INK

OUT = os.path.join(os.path.dirname(__file__), "GIS_Investment_Pitch_Template.pptx")

PLACEHOLDER_DESCRIPTOR = "[One-sentence descriptor — no trailing period]"


def main():
    deck = PitchDeck(company_header="[Company Name] (NASDAQ: [TICK])")
    deck.default_source = "Source: [Primary filing or data source]"

    deck.title_slide(
        company_name="[Company Name] inc.",
        ticker_line="NASDAQ: [TICK]",
        recommendation="[OVERWEIGHT / LONG | NEUTRAL | UNDERWEIGHT / SHORT]",
        current_price="$[XX.XX]",
        target_price="$[XXX]",
        upside_pct="+[XX]% upside",
        descriptor="[Investment thesis in one sentence — cyclical trough, not broken business, etc.]",
    )

    deck.toc_slide(STANDARD_TOC)

    deck.placeholder_slide(
        "Investment Thesis Summary",
        "[Overweight TICK with a $XXX target on ~XX% upside and an asymmetric payoff]",
        [
            ("Three reasons to be long", ["[Thesis pillar 1]", "[Thesis pillar 2]", "[Thesis pillar 3]"]),
        ],
    )

    deck.placeholder_slide(
        "Situation Overview",
        "[A high-quality compounder has de-rated to value multiples after a guidance reset]",
        [
            ("What happened", ["[Tape / price action]", "[Guidance reset]", "[Regional or margin pressure]"]),
            (
                "Where the business stands (latest fiscal year)",
                ["[Revenue and margin snapshot]", "[Balance sheet]", "[Cash flow]"],
            ),
        ],
    )

    deck.placeholder_slide(
        "Market Narrative",
        "[Sentiment has capitulated — the sell-side is cutting targets into the print]",
        [
            ("What the street is saying", ["[Consensus rating and target cuts]", "[Bear narrative]"]),
            ("Our differentiated view", ["[Why expectations are too low]", "[What we underwrite instead]"]),
        ],
    )

    deck.placeholder_slide(
        "Company Overview",
        "[A vertically integrated, direct-to-consumer [category] brand]",
        [
            ("Business", ["[Founded / HQ / product]", "[Channel mix]", "[Geographic reporting]"]),
            ("Revenue trajectory (US$ M)", ["[FY-3]", "[FY-2]", "[FY-1]", "[Latest FY]"]),
        ],
    )

    deck.placeholder_slide(
        "Business Model & Unit Economics",
        "[Premium pricing plus DTC scale drives sector-leading margins]",
        [
            ("Unit economics", ["[Gross margin]", "[Operating margin]", "[Store / asset efficiency]", "[FCF]"]),
            ("Competitive moats", ["[Moat 1]", "[Moat 2]", "[Moat 3]"]),
        ],
    )

    deck.placeholder_slide(
        "Industry Overview",
        "[Structural tailwinds favor scaled brands in a consolidating category]",
        [
            ("Category dynamics", ["[TAM / growth]", "[Competitive landscape]", "[Barriers to entry]"]),
            ("Where [TICK] fits", ["[Positioning]", "[Growth lever]", "[Differentiation vs peers]"]),
        ],
    )

    for num, desc in [
        ("I", "[The valuation already discounts a permanently shrinking business]"),
        ("II", "[International / segment X is the multi-year growth engine]"),
        ("III", "[Elite economics and capital return compound through the trough]"),
    ]:
        deck.placeholder_slide(
            f"Investment Thesis {num}",
            desc,
            [("[Headline thesis statement]", ["[Supporting bullet 1]", "[Supporting bullet 2]", "[Supporting bullet 3]"])],
        )

    deck.placeholder_slide(
        "Risks & Mitigants",
        "[The bear points are real but largely discounted at today's multiple]",
        [("Risk  →  Mitigant", ["[Risk 1] → [Mitigant 1]", "[Risk 2] → [Mitigant 2]"])],
    )

    deck.placeholder_slide(
        "Catalyst Timeline",
        "[A sequence of events that closes the gap to intrinsic value]",
        [("[When]", ["[Catalyst 1 — what to watch]", "[Catalyst 2]", "[Catalyst 3]"])],
    )

    for title in [
        "Financials — Income Statement",
        "Financials — Balance Sheet",
        "Financials — Cash Flow",
    ]:
        s = deck.slide_base(title, "[Historical results with our base-case forecast (US$ M)]", page=deck.pg())
        tb, tf = deck.body_box(s)
        add_para(tf, "[Insert financial table from 3-statement model]", 13, INK, first=True, space_after=5)
        add_para(tf, "[Footnote on forecast assumptions]", 11, GREY, italic=True)

    deck.placeholder_slide(
        "Capital Structure & WACC",
        "[A ~100% equity, net-cash structure drives a ~XX% discount rate]",
        [
            ("Capital structure", ["[Debt / cash]", "[EV bridge]", "[Capital return policy]"]),
            ("WACC BUILD (CAPM)", ["[rf]", "[ERP]", "[Beta]", "[WACC]"]),
        ],
    )

    deck.placeholder_slide(
        "Valuation Summary",
        "[Multiple methods converge above the current price — target $XXX]",
        [("[Football field methods]", ["[P/E range]", "[EV/EBITDA range]", "[DCF range]", "[52-week range]"])],
    )

    deck.placeholder_slide(
        "DCF Valuation",
        "[Base-case unlevered DCF yields ~$XXX per share]",
        [
            ("Base-case assumptions", ["[Revenue growth]", "[Margin path]", "[WACC / g]"]),
            ("Valuation output", ["[EV → equity → per share]"]),
        ],
    )

    deck.placeholder_slide(
        "Comparable Companies",
        "[TICK screens cheap on absolute and relative multiples]",
        [("Read-through", ["[Peer table]", "[Relative discount thesis]"])],
    )

    deck.placeholder_slide(
        "Appendix — Bull / Bear Scenarios",
        "[Asymmetric payoff: limited downside, substantial upside]",
        [("[Bear / Base / Bull]", ["[Bear case bullets]", "[Base case bullets]", "[Bull case bullets]"])],
    )

    s = deck.slide_base(
        "Sources & Disclaimer",
        "Data provenance and standard research disclaimer",
        page=deck.pg(),
    )
    tb, tf = deck.body_box(s)
    add_para(tf, "Sources", 14, CARD, bold=True, first=True, space_after=5)
    for t in [
        "[Historical financials: Form 10-K via SEC EDGAR]",
        "[Latest earnings release and guidance]",
        "[Market data source and as-of date]",
        "Projections, DCF and comparable-company analysis: GIS Investment Research models, built from scratch for this assignment",
    ]:
        add_para(tf, t, 12.5, INK, bullet=True, space_after=5)
    add_para(tf, "Disclaimer", 14, CARD, bold=True, space_after=5)
    add_para(
        tf,
        "This presentation is prepared for educational purposes as part of the Global Investment Society "
        "selection process and does not constitute investment advice or a recommendation to buy or sell any security",
        12,
        GREY,
        italic=True,
        space_after=0,
    )

    n = deck.save(OUT)
    print(f"Saved {os.path.abspath(OUT)} with {n} slides")


if __name__ == "__main__":
    main()
