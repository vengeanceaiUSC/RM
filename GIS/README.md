# GIS Investment Research — Pitch Deck Template

Reusable **python-pptx** template for Global Investment Society IR selection pitches. Matches the standard 22-slide structure and formatting rules from the GIS assignment.

## Files

| File | Purpose |
|------|---------|
| `gis_pitch.py` | Shared engine — colors, fonts, slide chrome, `PitchDeck` class |
| `build_blank_template.py` | Generates `GIS_Investment_Pitch_Template.pptx` with placeholders |
| `GIS_Investment_Pitch_Template.pptx` | **Full example deck** — Summit Outdoor Co. (NYSE: SUMM) with charts, tables, football field, DCF |
| `GIS_Investment_Pitch_Template.pdf` | PDF export of the example template |
| `GIS_Investment_Pitch_Template_Reference_Summit.pdf` | Same as template PDF — fictional worked example |
| `GIS_Investment_Pitch_Template_Reference_LULU.pdf` | Real filled pitch — LULU OVERWEIGHT (23 slides) |
| `template_data.py` | Sample numbers for Summit Outdoor — edit or replace for your company |
| `slide_layouts.py` | Reusable slide builders (charts, KPI boxes, football field, DCF grid) |

## Formatting rules (enforced by `gis_pitch.py`)

- **Font:** Garamond throughout
- **Header:** Running “Investment Research Division” + `Company (EXCHANGE: TICK)` left; slide title in 0.3" navy rectangle, 15pt white, right-aligned
- **Descriptor:** One sentence under the header, grey italic, **no trailing period**
- **Colors:** Navy `#1F2A44`, USC cardinal `#990000`, gold `#FFC72C`
- **Footer:** Source line (left) + page number (right)
- **Data convention:** Blue = reported, black = calculated, red = assumptions (Excel models; deck uses cardinal for headers)

## Standard slide order (22 slides)

1. Title  
2. Table of contents  
3. Investment thesis summary  
4. Situation overview  
5. Market narrative  
6. Company overview  
7. Business model & unit economics  
8. Industry overview  
9–11. Thesis I / II / III  
12. Risks & mitigants  
13. Catalyst timeline  
14–16. Financials (IS / BS / CF)  
17. Capital structure & WACC  
18. Valuation summary (football field)  
19. DCF valuation  
20. Comparable companies  
21. Appendix — bull / bear scenarios  
22. Sources & disclaimer  

Optional slides (e.g. **Macro × micro overlap**) can be inserted after Situation — see `LULU/scripts/build_pitch.py`.

## What you get

The template is **not** a list of `[placeholder]` bullets. It ships as a complete **fictional example** (Summit Outdoor Co., NYSE: SUMM) with:

- KPI boxes on the thesis slide (price target / rating / why now)
- **Revenue bar chart** and segment architecture box (company overview)
- **Operating margin line chart** + margin profile panel (unit economics)
- **Category mix pie chart** (industry)
- Thesis slides with side metric panels
- Risk ↔ mitigant dual-column layout
- Catalyst timeline with navy date blocks
- Full **IS / BS / CF tables** with historical + forecast columns
- WACC build panel
- **Football field** valuation chart with current/target markers
- DCF assumptions table + output bridge + **sensitivity grid**
- Comps table + bull/base/bear scenario columns

Replace `template_data.py` values and narrative strings — keep layouts via `slide_layouts.py`.

## Quick start — regenerate template

```bash
pip install python-pptx
cd GIS
python3 build_blank_template.py
# → GIS_Investment_Pitch_Template.pptx
```

Open the PPTX and replace `[placeholder]` text, or build programmatically (recommended for reproducibility).

## Quick start — new company pitch (code)

```python
import sys, os
sys.path.insert(0, os.path.join("..", "GIS"))
from gis_pitch import PitchDeck, add_para, rect, NAVY, CARD, INK

deck = PitchDeck(company_header="Acme Corp (NYSE: ACME)")
deck.default_source = "Source: company Form 10-K (CIK 0001234567)"

deck.title_slide(
    company_name="Acme Corp",
    ticker_line="NYSE: ACME",
    recommendation="OVERWEIGHT / LONG",
    current_price="$50.00",
    target_price="$70",
    upside_pct="+40% upside",
    descriptor="Your one-sentence thesis",
)

s = deck.slide_base("Investment Thesis Summary", "Your descriptor", page=deck.pg())
tb, tf = deck.body_box(s)
add_para(tf, "Three reasons to be long", 15, CARD, bold=True, first=True)
# ... add content ...

deck.save("ACME_Investment_Pitch_Deck.pptx")
```

## Reference implementation

`LULU/scripts/build_pitch.py` is the full worked example: pulls financials from `data.py`, builds tables, football field, DCF slide, and comps. Rebuild:

```bash
cd LULU/scripts
python3 build_pitch.py
```

## Team preferences (assignment)

1. Special Situations  
2. Global Macro  
3. Technology  
4. Generalist  

*Educational purposes — GIS selection process — not investment advice.*
