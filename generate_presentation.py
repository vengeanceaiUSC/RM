"""Generate the Lululemon (LULU) investment pitch presentation.

The deck contains a title slide (slide 1) followed by the content slides 2-13.
Per the source brief, slide 8 is "Part 2" of the slide 7 industry overview, which
shifts the remaining investment-thesis / risk / timeline slides down by one so that
all of the provided content lands in slides 2-13.

All slide body text and sources are stored verbatim in ``SLIDES`` so the content
can be verified word for word.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# Brand-ish palette
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
RED = RGBColor(0xC8, 0x10, 0x2E)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xF2, 0xF4, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# Slides 2-13. Each entry: title, subtitle (may be ""), bullets, sources.
SLIDES = [
    {
        "number": 2,
        "title": "Situation Overview and Current Investment Setup",
        "subtitle": "This slide outlines why this investment opportunity exists and the historical financial context driving the current setup",
        "bullets": [
            "The recent guidance cut triggered a massive cyclical panic pushing Lululemon down to roughly 100 dollars [1]",
            "Despite historically compounding double digit growth the market capitulated over a guidance cut of 5 to 7 percent [2]",
            "Lululemon still retains durable cash flows with clean run rate operating margins of 13.2 percent [3]",
        ],
        "sources": [
            "[1] TIKR LULU Stock Crashed 17%: https://www.tikr.com/blog/lululemon-stock-crashed-17-on-friday-the-guidance-cut-was-the-real-story",
            "[2] Lululemon Q2 FY2026 Guidance Release: https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
            "[3] Provided Valuation Model: LULUMODEL18.xlsx",
        ],
    },
    {
        "number": 3,
        "title": "Market Narrative and Analyst Sentiment Surrounding the Stock",
        "subtitle": "This slide breaks down current market sentiment and exactly what analysts are saying about the recent guidance cut",
        "bullets": [
            "Morgan Stanley issued a highly pessimistic forecast after management aggressively cut 2026 revenue guidance to 10.35 billion [1]",
            "Analysts are paralyzed by cyclical fears as the consensus price target was slashed from 176 dollars to 136 dollars [2]",
            "Firms like BTIG maintain neutral ratings due to short-term turbulence but ignore the durable competitive advantage we see [3]",
        ],
        "sources": [
            "[1] MarketBeat: Morgan Stanley Issues Pessimistic Forecast for lululemon athletica: https://www.marketbeat.com/instant-alerts/analyst-morgan-stanley-issues-pessimistic-forecast-for-lululemon-athletica-nasdaq-lulu-stock-price-2026-09-04/",
            "[2] Simply Wall St: lululemon athletica Stock Analysis: https://simplywall.st/stocks/us/consumer-durables/nasdaq-lulu/lululemon-athletica",
            "[3] GuruFocus: LULU Reiterates by BTIG - Rating Maintained at Neutral: https://www.gurufocus.com/news/9068272/lulu-reiterates-by-btig-rating-maintained-at-neutral",
        ],
    },
    {
        "number": 4,
        "title": "Investment Thesis Summary and Target Price",
        "subtitle": "This slide breaks down our actual 133 dollar base case target price and the three core pillars supporting our overweight recommendation",
        "bullets": [
            "Our discounted cash flow valuation generates a base case implied share price of 133.64 dollars representing a 33.6 percent upside from current levels [1]",
            "The first pillar is profitability because adjusting out the tariff refunds reveals LULU still maintains a highly resilient 13.2 percent clean run-rate operating margin [2]",
            "The second pillar is our mathematically sound 9.0 percent WACC which strictly bounds our 2.3 percent long-term revenue growth assumption [3]",
            "Finally international expansion remains the crucial growth engine as 4 percent growth in China Mainland easily offsets the temporary North American stagnation [4]",
        ],
        "sources": [
            "[1] Provided Valuation Model (Base Case Implied Value): LULUMODEL18_3.xlsx",
            "[2] Lululemon Q2 FY2026 Earnings Release (13.2% Margin Calc): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
            "[3] Provided Valuation Model (WACC & Revenue Drivers): LULUMODEL18_3.xlsx",
            "[4] Lululemon Q2 FY2026 Earnings Release (International Growth): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
        ],
    },
    {
        "number": 5,
        "title": "Geographic Segments and Revenue Breakdown",
        "subtitle": "This slide outlines Lululemon's geographic segments and revenue breakdown while detailing why China growth offsets temporary US declines",
        "bullets": [
            "Lululemon generated 11.1 billion dollars in total revenue with Americas contributing 7.85 billion dollars or 70.68 percent [1]",
            "Americas comparable sales fell 3 percent due to temporary cyclical macro pressure rather than structural brand degradation [2]",
            "China Mainland surged 20 percent to 1.75 billion dollars proving high growth international expansion easily offsets US temporary weakness [3]",
        ],
        "sources": [
            "[1] Lululemon FY2025 Form 10-K (Segment Revenue and Percentages): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
            "[2] Lululemon FY2025 Form 10-K (Americas Comparable Sales): https://corporate.lululemon.com/~/media/Files/L/Lululemon/investors/results-center/q2-2026-financial-supplement.pdf",
            "[3] Provided Valuation Model (China Mainland Growth & Revenue Driver): LULUMODEL18.xlsx",
        ],
    },
    {
        "number": 6,
        "title": "Business Model Unit Economics and Competitive Moats",
        "subtitle": "This slide analyzes Lululemon's unit economics and competitive moats that protect its long-term market leadership",
        "bullets": [
            "Lululemon sustains a 56.6 percent gross margin providing a massive competitive moat against retail price competition [1]",
            "E-commerce unit economics remain elite with direct-to-consumer EBIT margins hitting 23.6 percent without retail occupancy drag [2]",
            "Company-operated stores generate 1,426 dollars per square foot and an 18.6 percent EBIT margin establishing high capital efficiency [3]",
        ],
        "sources": [
            "[1] Provided Valuation Model (Gross Margin): LULUMODEL18_6.xlsx",
            "[2] Provided Valuation Model (E-commerce EBIT Margin Driver): LULUMODEL18_6.xlsx",
            "[3] Provided Valuation Model (Store EBIT Margin & SPSF): LULUMODEL18_6.xlsx",
        ],
    },
    {
        "number": 7,
        "title": "Industry Overview - Trends and Structure",
        "subtitle": "This slide explores the ongoing athleisure industry trends including market fragmentation and the barriers to entry",
        "bullets": [
            "I argue the global athleisure market remains highly fragmented meaning most players lack a durable competitive advantage [1]",
            "The premium segment maintains high barriers to entry protecting global champions from temporary cyclical noise [2]",
            "Lululemon mathematically proves its moat by generating a 30.25 percent ROE far outpacing the 6.81 percent industry median [3]",
        ],
        "sources": [
            "[1] Market.us Media (Athleisure Market Fragmentation): https://media.market.us/athleisure-industry-statistics/",
            "[2] Fortune Business Insights (Premium Athleisure Trends): https://www.fortunebusinessinsights.com/athleisure-market-110642",
            "[3] FinanceCharts (Lululemon ROE & Retail Median): https://www.financecharts.com/stocks/LULU/growth/roe",
        ],
    },
    {
        "number": 8,
        "title": "Industry Overview - Barriers to Entry and Profitability",
        "subtitle": "This slide dissects capital efficiency metrics comparing Lululemon's gross margin directly against legacy apparel competitors",
        "bullets": [
            "Lululemon commands a 56.6 percent gross margin far exceeding traditional athletic apparel peers like Nike and Under Armour [1]",
            "Nike struggles to maintain a 43 percent margin while Under Armour hovers around 46 percent reflecting their wholesale dependence [2]",
            "Lululemon's direct-to-consumer scale prevents this structural margin compression and forms an impenetrable economic moat against industry price wars [3]",
        ],
        "sources": [
            "[1] Provided Valuation Model (Gross Margin Assumptions): LULUMODEL18.xlsx",
            "[2] Investing.com (Nike & Under Armour Historical Gross Margins): https://www.investing.com/pro/NYSE:NKE/explorer/gp_margin",
            "[3] ProAnalyst LULU Market Report (Competitor Margins & DTC Moat): https://lulu.proanalyst.ai/business",
        ],
    },
    {
        "number": 9,
        "title": "Investment Thesis I",
        "subtitle": "This slide outlines the core contrarian investment thesis utilizing numerical evidence from the LULUMODEL18.xlsx file",
        "bullets": [
            "The 2026 price drop exceeding 50% has left Lululemon critically undervalued despite durable cash flows [1]",
            "Its premium Direct-to-Consumer revenue mix protects a massive 54.9% adjusted gross margin, mathematically justifying my intrinsic valuation thesis [2]",
            "Furthermore, even with Americas growth stagnating, sheer cash flow generation creates a robust intrinsic valuation buffer, where a conservative 2.25% terminal growth rate still yields over 30% upside to our $133.64 target price [3]",
        ],
        "sources": [
            "[1] https://everythingmoney.com/blog/lululemon-is-collapsing-burry-s-biggest-bet-4334",
            "[2] https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
            "[3] LULUMODEL18.xlsx, DCF Terminal Value & Valuation Summary Schedule is that ready",
        ],
    },
    {
        "number": 10,
        "title": "Investment Thesis II  Partial Margin Recovery & Brand Loyalty Floor",
        "subtitle": "This slide details how partial margin recovery supported by core brand loyalty still drives a highly compelling valuation",
        "bullets": [
            "Historical pandemic-era peak EBIT margins reached 23.7% in FY24 showing prior peak earnings power [1]",
            "Our valuation assumes a floor built on resilient baseline brand loyalty, proving Lululemon does not need to remain the hottest viral trend 100% of the time to sustain a 13.2% trough margin [2]",
            "A modest partial recovery to just 15.5% EBIT margin by FY30 still yields $133.64 per share [3]",
            "This proves returning to peak COVID profitability is completely unnecessary to unlock substantial market upside [4]",
        ],
        "sources": [
            "[1] LULUMODEL18.xlsx / SEC EDGAR Form 10-K (FY24 Peak Margins)",
            "[2] LULUMODEL18.xlsx, Clean Run-Rate EBIT Margin & Brand Royalty Assumptions",
            "[3] LULUMODEL18.xlsx, Base Case DCF Valuation Summary",
            "[4] LULUMODEL18.xlsx, Discounted Cash Flow Valuation Summary",
        ],
    },
    {
        "number": 11,
        "title": "Investment Thesis III: Geographic Growth Divergence",
        "subtitle": "This slide examines how Lululemon's top-line projections rely disproportionately on Chinese market expansion to conceal domestic North American stagnation.",
        "bullets": [
            "The revenue build reveals Americas facing near-term contraction with -4.0% comps in FY26 flatlining at a terminal 2.0% growth rate by FY30 [1].",
            "To offset this domestic anchor, the model relies entirely on disproportionate FY26\u2013FY30 Chinese footprint expansion, averaging 16 new stores annually versus just 6 domestically, and sustained double-digit (10.2% average) comp growth to overcome clear Americas expansion drawbacks [2].",
            "Consequently, if the Chinese consumer softens, this model's core top-line projections will not be optimal enough to meet our target [3].",
        ],
        "sources": [
            "[1] LULUMODEL18.xlsx, Americas FY26-FY30 Comparable Sales Growth Assumptions",
            "[2] LULUMODEL18.xlsx, FY26 Store Openings & Mainland China Revenue Build",
            "[3] LULUMODEL18.xlsx, Revenue Drivers Schedule",
        ],
    },
    {
        "number": 12,
        "title": "Risk & Mitigants",
        "subtitle": "This slide evaluates core downside risks and demonstrates how share repurchases compound EPS to drive share price recovery",
        "bullets": [
            "China deceleration risks a $56.00 bear floor, but $45.64 cumulative cash per share recovers 45% of entry price [1]",
            "Slashed CapEx saves $360M annually by relying on online e-commerce's 23.6% EBIT margin plus $101.5M inventory releases [2]",
            "Deploying $500M annually into buybacks retires 18.7 million shares, compounding EPS to $13.36 to support our target price of $133.64.",
        ],
        "sources": [
            "[1] LULUMODEL18.xlsx, Bear Case DCF Valuation Summary / https://www.barrons.com/articles/lululemon-stock-earnings-guidance-a7a7c5c0",
            "[2] LULUMODEL18.xlsx, CapEx & E-Commerce Channel EBIT Assumptions / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
            "[3] LULUMODEL18.xlsx, Share Repurchase & EPS Accretion Schedule / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
        ],
    },
    {
        "number": 13,
        "title": "Timeline of Recovery",
        "subtitle": "",
        "bullets": [
            "2026 (Q3 FY2026 Trough): Q3 FY2026 revenue laps guidance trough while China Double 11 sales confirm holiday store traffic floor stabilization across key markets [1]",
            "2027 (FY2026 Year-End): First full-year reset absorbs steep prior declines while $134.5M tariff refunds and targeted SG&A cost actions protect earnings per share [2]",
            "2027 (FY2027 Margin Inflection): Operating margins expand to 13.8% (+60 bps) as promotional headwinds anniversary, while 75% UFCF buybacks compound EPS to boost sentiment [3]",
            "2028 (FY2027\u2013FY2028 Multiple Re-Rating): Accelerating international store scaling (+12% China comps) offsets Americas softness (-4%), driving overall revenue recovery and valuation re-rating toward $133.64 [4]",
        ],
        "sources": [
            "[1] LULUMODEL18.xlsx, Scenarios & Revenue Drivers / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
            "[2] LULUMODEL18.xlsx, NOPAT Bridge & Scenarios / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
            "[3] LULUMODEL18.xlsx, Scenarios & Unlevered Free Cash Flow Schedule / https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
            "[4] LULUMODEL18.xlsx, Revenue Drivers & DCF Valuation Summary / https://stockanalysis.com/stocks/lulu/forecast/",
        ],
    },
]


def _blank_layout(prs):
    # Layout index 6 is the blank layout in the default template.
    return prs.slide_layouts[6]


def add_title_slide(prs):
    slide = prs.slides.add_slide(_blank_layout(prs))

    band = slide.shapes.add_shape(
        1, Inches(0), Inches(2.4), prs.slide_width, Inches(2.7)
    )
    band.fill.solid()
    band.fill.fore_color.rgb = NAVY
    band.line.fill.background()
    band.shadow.inherit = False

    tb = slide.shapes.add_textbox(Inches(0.9), Inches(2.55), Inches(11.5), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Lululemon Athletica (NASDAQ: LULU)"
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = WHITE

    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "A Contrarian Overweight Investment Thesis"
    r2.font.size = Pt(22)
    r2.font.color.rgb = RGBColor(0xD9, 0xE1, 0xEC)

    tb2 = slide.shapes.add_textbox(Inches(0.9), Inches(5.3), Inches(11.5), Inches(1.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p3 = tf2.paragraphs[0]
    r3 = p3.add_run()
    r3.text = "Base Case Target Price: $133.64  |  33.6% Upside"
    r3.font.size = Pt(18)
    r3.font.bold = True
    r3.font.color.rgb = RED


def add_content_slide(prs, data):
    slide = prs.slides.add_slide(_blank_layout(prs))

    # Header band
    header = slide.shapes.add_shape(
        1, Inches(0), Inches(0), prs.slide_width, Inches(1.15)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = NAVY
    header.line.fill.background()
    header.shadow.inherit = False

    # Slide number chip
    chip = slide.shapes.add_shape(1, Inches(0.35), Inches(0.28), Inches(0.6), Inches(0.6))
    chip.fill.solid()
    chip.fill.fore_color.rgb = RED
    chip.line.fill.background()
    chip.shadow.inherit = False
    ctf = chip.text_frame
    ctf.word_wrap = True
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = ctf.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run()
    cr.text = str(data["number"])
    cr.font.size = Pt(20)
    cr.font.bold = True
    cr.font.color.rgb = WHITE

    # Title
    title_box = slide.shapes.add_textbox(Inches(1.15), Inches(0.18), Inches(11.9), Inches(0.85))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = ttf.paragraphs[0]
    tr = tp.add_run()
    tr.text = data["title"]
    tr.font.size = Pt(24)
    tr.font.bold = True
    tr.font.color.rgb = WHITE

    top = 1.4
    # Subtitle
    if data["subtitle"]:
        sub_box = slide.shapes.add_textbox(Inches(0.6), Inches(top), Inches(12.1), Inches(0.6))
        stf = sub_box.text_frame
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sr = sp.add_run()
        sr.text = data["subtitle"]
        sr.font.size = Pt(14)
        sr.font.italic = True
        sr.font.color.rgb = GREY
        top += 0.75

    # Bullets
    body_box = slide.shapes.add_textbox(Inches(0.6), Inches(top), Inches(12.1), Inches(3.4))
    btf = body_box.text_frame
    btf.word_wrap = True
    for i, bullet in enumerate(data["bullets"]):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        p.space_after = Pt(10)
        run = p.add_run()
        run.text = "\u25aa  " + bullet
        run.font.size = Pt(16)
        run.font.color.rgb = DARK

    # Sources
    src_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.55), Inches(12.1), Inches(1.75))
    xtf = src_box.text_frame
    xtf.word_wrap = True
    hp = xtf.paragraphs[0]
    hr = hp.add_run()
    hr.text = "Links & Sources:"
    hr.font.size = Pt(10)
    hr.font.bold = True
    hr.font.color.rgb = NAVY
    for src in data["sources"]:
        p = xtf.add_paragraph()
        run = p.add_run()
        run.text = src
        run.font.size = Pt(8)
        run.font.color.rgb = GREY


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs)
    for data in SLIDES:
        add_content_slide(prs, data)

    out = "Lululemon_Investment_Thesis.pptx"
    prs.save(out)
    print(f"Saved {out} with {len(prs.slides._sldIdLst)} slides")


if __name__ == "__main__":
    build()
