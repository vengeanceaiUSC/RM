"""Add the Revenue Drivers worksheet to the DCF workbook."""
import styles as S
from styles import (
    write, write_reported, write_assumption_docs,
    write_equation_summary, write_geo_scale_equation, write_ctrl_f, NUM, PCT, MONEY,
)
import data as D
import driver_kpis as DK

FY = D.PROJ_YEARS
FY25 = "B"
FCOLS = ["C", "D", "E", "F", "G"]
UNITS_COL = "H"
DEQ = "I"                    # dedicated equation summary column
DJ, DS, DC = "J", "K", "L"   # justification | source | Ctrl+F
K = load_kpis = DK.load_kpis
DOC_COLS = "ABCDEFGHIJKL"


def build_revenue_drivers(wb, scen_base_col, rev_row_map, deq=DEQ, dj=DJ, ds=DS, dc=DC):
    """Create Revenue Drivers tab. rev_row_map maps FY label -> Scenarios revenue row."""
    ws = wb.create_sheet("Revenue Drivers")
    ws.sheet_view.showGridLines = False
    S.set_col_widths(ws, {
        "A": 42, "B": 12,
        "C": 11, "D": 11, "E": 11, "F": 11, "G": 11,
        "H": 10,
        "I": 40, "J": 28, "K": 16, "L": 36,
    })
    kpis = load_kpis()
    rr = [1]

    def f_docs(row, eq_key, just_key=None, internal=None):
        """Equation in col I; casual why in col J; Ctrl+F only when not an equation string."""
        jkey = just_key or eq_key
        write_equation_summary(ws, row, deq, eq_key, D.SOURCE_HINT)
        ctrl = D.SOURCE_HINT.get(jkey)
        hints = D.SOURCE_HINT if ctrl and not str(ctrl).lstrip("'").startswith("=") else None
        write_assumption_docs(ws, row, dj, ds, dc, jkey, D.JUST, D.ASSUMPTION_SRC,
                              hints=hints, internal_location=internal)

    def hdr(title):
        write(ws, f"A{rr[0]}", title, S.WHITE, bold=True, size=11, fillc=S.DARK, align=S.left_indent)
        for c in DOC_COLS:
            ws[f"{c}{rr[0]}"].fill = S.fill(S.DARK)
        rr[0] += 1

    def sub(title):
        write(ws, f"A{rr[0]}", title, S.ACCENT, bold=True, size=10, align=S.left_indent)
        rr[0] += 1

    def col_hdr():
        write(ws, f"A{rr[0]}", "Driver", S.WHITE, bold=True, size=9, fillc=S.NAVY, align=S.left_indent)
        write(ws, f"{FY25}{rr[0]}", "FY2025A", S.WHITE, bold=True, size=9, fillc=S.NAVY, align=S.center)
        for y, c in zip(FY, FCOLS):
            write(ws, f"{c}{rr[0]}", y, S.WHITE, bold=True, size=9,
                  fillc=S.ACCENT if y == "FY2026E" else S.NAVY, align=S.center)
        write(ws, f"{UNITS_COL}{rr[0]}", "Units", S.WHITE, bold=True, size=9, fillc=S.NAVY, align=S.center)
        write(ws, f"{deq}{rr[0]}", "Equation", S.WHITE, bold=True, size=9, fillc=S.ACCENT, align=S.center)
        write(ws, f"{dj}{rr[0]}", "Justification  [cols I\u2013L: click + to expand]", S.WHITE, bold=True, size=9, fillc=S.NAVY)
        write(ws, f"{ds}{rr[0]}", "Source", S.WHITE, bold=True, size=9, fillc=S.NAVY)
        write(ws, f"{dc}{rr[0]}", "Ctrl+F", S.WHITE, bold=True, size=9, fillc=S.NAVY)
        rr[0] += 1

    def row(label, fy25_val, proj_vals, fmt=NUM, red=False, doc_key=None, units="", formula_row=False):
        r = rr[0]
        write(ws, f"A{r}", label, S.BLACK, size=9, align=S.left_indent)
        color = S.RED if red else S.BLACK
        if fy25_val is not None and not formula_row:
            if isinstance(fy25_val, (int, float)) and not red:
                write_reported(ws, f"{FY25}{r}", fy25_val, D.filing_url("FY2025"),
                               size=9, numfmt=fmt, align=S.right)
            else:
                write(ws, f"{FY25}{r}", fy25_val, color, size=9, numfmt=fmt, align=S.right)
        for i, c in enumerate(FCOLS):
            v = proj_vals[i] if i < len(proj_vals) else None
            if v is not None:
                if isinstance(v, str):
                    write(ws, f"{c}{r}", v, S.BLACK, size=9, numfmt=fmt, align=S.right)
                else:
                    write(ws, f"{c}{r}", v, S.RED if red else S.BLACK, size=9, numfmt=fmt, align=S.right)
        if units:
            write(ws, f"{UNITS_COL}{r}", units, S.BLACK, italic=True, size=8, align=S.center)
        if doc_key:
            write_assumption_docs(ws, r, dj, ds, dc, doc_key, D.JUST, D.ASSUMPTION_SRC, hints=D.SOURCE_HINT)
        rr[0] += 1
        return r

    hdr("BOTTOM-UP REVENUE DRIVER SCHEDULE (FY2026E–FY2030E)")
    write(ws, f"A{rr[0]}",
          "FY2025A = Firecrawl-ingested 10-K anchors (blue). FY26–30 = red operational assumptions. "
          "Consolidated revenue cross-checks to Scenarios base case.",
          S.BLACK, italic=True, size=8, align=S.left_indent)
    rr[0] += 1
    write(ws, f"{deq}{rr[0]}", "Equation (every formula row)", S.ACCENT, bold=True, size=9)
    write(ws, f"{dj}{rr[0]}", "Justification (~20 words)", S.ACCENT, bold=True, size=9)
    write(ws, f"{ds}{rr[0]}", "Source (click)", S.ACCENT, bold=True, size=9)
    write(ws, f"{dc}{rr[0]}", "Ctrl+F (prove number)", S.ACCENT, bold=True, size=9)
    ws.freeze_panes = "I5"
    rr[0] += 1

    # --- STORE FLEET BY GEO ---
    sub("I. STORE FLEET & SQUARE FOOTAGE (company-operated)")
    col_hdr()
    geo = [("americas", "Americas"), ("china", "China Mainland"), ("row", "Rest of World")]
    R = {}
    end = {}
    opn = {}
    cls = {}
    for key, label in geo:
        fy25_end = kpis[f"stores_{key}"]["FY2025"]
        fy24_beg = kpis[f"stores_{key}"]["FY2024"]
        beg_row = rr[0]
        write(ws, f"A{beg_row}", f"  {label} — beginning stores", S.BLACK, size=9, align=S.left_indent)
        write(ws, f"{FY25}{beg_row}", fy24_beg, S.BLUE, size=9, numfmt=NUM, align=S.right)
        write(ws, f"{UNITS_COL}{beg_row}", "stores", S.BLACK, italic=True, size=8, align=S.center)
        rr[0] += 1
        opn[key] = row(f"  {label} — openings", None, DK.FORECAST["openings"][key], red=True,
                       doc_key="drv_openings", units="stores")
        cls[key] = row(f"  {label} — closures", None, DK.FORECAST["closures"][key], red=True,
                       doc_key="drv_closures", units="stores")
        end[key] = rr[0]
        write(ws, f"A{end[key]}", f"  {label} — ending stores", S.BLACK, bold=True, size=9, align=S.left_indent)
        write(ws, f"{FY25}{end[key]}", fy25_end, S.BLUE, size=9, numfmt=NUM, align=S.right)
        for i, c in enumerate(FCOLS):
            write(ws, f"{c}{end[key]}",
                  f"={c}{beg_row}+{c}{opn[key]}-{c}{cls[key]}",
                  S.BLACK, bold=True, size=9, numfmt=NUM, align=S.right)
        write(ws, f"{UNITS_COL}{end[key]}", "stores", S.BLACK, italic=True, size=8, align=S.center)
        f_docs(end[key], "drv_f_end_stores")
        rr[0] += 1
        for i, c in enumerate(FCOLS):
            write(ws, f"{c}{beg_row}", f"={FY25 if i == 0 else FCOLS[i-1]}{end[key]}",
                  S.BLACK, size=9, numfmt=NUM, align=S.right)
        f_docs(beg_row, "drv_f_beg_stores")

    R["total_end"] = rr[0]
    write(ws, f"A{rr[0]}", "Total ending stores", S.BLACK, bold=True, size=9, align=S.left_indent)
    write_reported(ws, f"{FY25}{rr[0]}", kpis["stores_total"]["FY2025"], D.filing_url("FY2025"),
                   bold=True, size=9, numfmt=NUM, align=S.right)
    for c in FCOLS:
        write(ws, f"{c}{rr[0]}",
              f"={c}{end['americas']}+{c}{end['china']}+{c}{end['row']}",
              S.BLACK, bold=True, size=9, numfmt=NUM, align=S.right)
    f_docs(R["total_end"], "drv_f_total_stores")
    rr[0] += 1

    R["sqft_store"] = row("Avg square feet per store", kpis["avg_sqft_per_store"]["FY2025"],
                          DK.FORECAST["avg_sqft_per_store"][1:], red=True, doc_key="drv_sqft_store", units="sq ft")
    R["total_sqft"] = rr[0]
    write(ws, f"A{rr[0]}", "Total square footage (end stores × avg sq ft)", S.BLACK, size=9, align=S.left_indent)
    write(ws, f"{FY25}{rr[0]}",
          f"={FY25}{R['total_end']}*{FY25}{R['sqft_store']}",
          S.BLACK, size=9, numfmt=NUM, align=S.right)
    for c in FCOLS:
        write(ws, f"{c}{rr[0]}",
              f"={c}{R['total_end']}*{c}{R['sqft_store']}",
              S.BLACK, size=9, numfmt=NUM, align=S.right)
    write(ws, f"{UNITS_COL}{rr[0]}", "sq ft", S.BLACK, italic=True, size=8, align=S.center)
    f_docs(R["total_sqft"], "drv_f_total_sqft")
    rr[0] += 1

    # --- STORE PRODUCTIVITY ---
    rr[0] += 1
    sub("II. STORE PRODUCTIVITY (brick & mortar channel)")
    col_hdr()
    R["spsf"] = row("Sales per square foot ($)", kpis["sales_per_sqft"]["FY2025"],
                     DK.FORECAST["sales_per_sqft"][1:], red=True, doc_key="drv_spsf", units="$/sq ft")
    for key, label in geo:
        comp = kpis.get("comp_sales", {}).get(key, {}).get("FY2025", DK.FORECAST["comp_sales"][key][0])
        R[f"comp_{key}"] = row(f"  {label} — comparable sales growth",
                               comp, DK.FORECAST["comp_sales"][key], fmt=PCT, red=True,
                               doc_key=f"drv_comp_{key}", units="%")
    R["store_rev_base"] = row("Prior-year store channel revenue ($000)",
                              kpis["revenue_stores"]["FY2025"], [None] * 5, units="$000")
    R["store_comp_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "Comparable store revenue ($000)", S.BLACK, size=9, align=S.left_indent)
    write(ws, f"{FY25}{rr[0]}", kpis["revenue_stores"]["FY2025"], S.BLUE, size=9, numfmt=NUM, align=S.right)
    for i, c in enumerate(FCOLS):
        prev = f"{FY25 if i == 0 else FCOLS[i-1]}{R['store_rev_base']}"
        blend = (
            f"={prev}*0.71*(1+{c}{R['comp_americas']})+"
            f"{prev}*0.16*(1+{c}{R['comp_china']})+"
            f"{prev}*0.13*(1+{c}{R['comp_row']})"
        )
        write(ws, f"{c}{rr[0]}", blend, S.BLACK, size=9, numfmt=NUM, align=S.right)
    f_docs(R["store_comp_rev"], "drv_comp_store_rev")
    rr[0] += 1
    R["new_store_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "Net new store revenue contribution ($000)", S.BLACK, size=9, align=S.left_indent)
    write(ws, f"{FY25}{rr[0]}", 0, S.BLACK, size=9, numfmt=NUM, align=S.right)
    for i, c in enumerate(FCOLS):
        net_open = (
            f"({c}{opn['americas']}-{c}{cls['americas']}+"
            f"{c}{opn['china']}-{c}{cls['china']}+"
            f"{c}{opn['row']}-{c}{cls['row']})"
        )
        write(ws, f"{c}{rr[0]}",
              f"={net_open}*{c}{R['sqft_store']}*{c}{R['spsf']}/1000*0.55",
              S.BLACK, size=9, numfmt=NUM, align=S.right)
    f_docs(R["new_store_rev"], "drv_new_store_rev")
    rr[0] += 1
    R["store_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "Store channel revenue ($000)", S.BLACK, bold=True, size=9, align=S.left_indent)
    write_reported(ws, f"{FY25}{rr[0]}", kpis["revenue_stores"]["FY2025"], D.filing_url("FY2025"),
                   bold=True, size=9, numfmt=NUM, align=S.right)
    for i, c in enumerate(FCOLS):
        write(ws, f"{c}{rr[0]}",
              f"={c}{R['store_comp_rev']}+{c}{R['new_store_rev']}",
              S.BLACK, bold=True, size=9, numfmt=NUM, align=S.right)
    for i, c in enumerate(FCOLS):
        write(ws, f"{c}{R['store_rev_base']}",
              f"={FY25 if i == 0 else FCOLS[i-1]}{R['store_rev']}",
              S.BLACK, size=9, numfmt=NUM, align=S.right)
    f_docs(R["store_rev_base"], "drv_f_store_rev_base")
    f_docs(R["store_rev"], "drv_f_store_rev")
    rr[0] += 1

    row("  memo: fleet traffic (millions of visits)", 42.5,
        [40.0, 41.0, 42.0, 43.0, 44.0], fmt='#,##0.0', red=True, doc_key="drv_store_traffic", units="M visits")
    row("  memo: in-store conversion rate", 0.28,
        [0.27, 0.275, 0.28, 0.285, 0.29], fmt=PCT, red=True, doc_key="drv_store_conv")
    row("  memo: in-store average transaction ($)", 118,
        [115, 116, 118, 120, 122], fmt=MONEY, red=True, doc_key="drv_store_aov")

    # --- DTC ---
    rr[0] += 1
    sub("III. DIRECT-TO-CONSUMER (e-commerce / digital)")
    col_hdr()
    R["sessions"] = row("Unique digital sessions (millions)", kpis["ecomm_sessions_m"]["FY2025"],
                        DK.FORECAST["ecomm_sessions_m"][1:], red=True, doc_key="drv_ecomm_sessions", units="M")
    R["econv"] = row("Digital conversion rate", kpis["ecomm_conversion"]["FY2025"],
                     DK.FORECAST["ecomm_conversion"][1:], fmt=PCT, red=True, doc_key="drv_ecomm_conv")
    R["eaov"] = row("Digital average order value ($)", kpis["ecomm_aov"]["FY2025"],
                    DK.FORECAST["ecomm_aov"][1:], fmt=MONEY, red=True, doc_key="drv_ecomm_aov")
    R["ecomm_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "E-commerce revenue ($000)", S.BLACK, bold=True, size=9, align=S.left_indent)
    write_reported(ws, f"{FY25}{rr[0]}", kpis["revenue_ecomm"]["FY2025"], D.filing_url("FY2025"),
                   bold=True, size=9, numfmt=NUM, align=S.right)
    for c in FCOLS:
        write(ws, f"{c}{rr[0]}",
              f"={c}{R['sessions']}*1000000*{c}{R['econv']}*{c}{R['eaov']}/1000",
              S.BLACK, bold=True, size=9, numfmt=NUM, align=S.right)
    f_docs(R["ecomm_rev"], "drv_f_ecomm_rev")
    rr[0] += 1

    # --- OTHER + GEO + CATEGORY ---
    rr[0] += 1
    sub("IV. OTHER CHANNELS, GEOGRAPHY & PRODUCT MIX")
    col_hdr()
    R["other_rev"] = row("Other channels (wholesale / license / outlets) ($000)",
                         kpis["revenue_other"]["FY2025"],
                         [None]*5, units="$000")
    for i, c in enumerate(FCOLS):
        prev = f"{FY25 if i == 0 else FCOLS[i-1]}{R['other_rev']}"
        g = DK.FORECAST["other_rev_growth"][i]
        write(ws, f"{c}{R['other_rev']}", f"={prev}*(1+{g})", S.BLACK, size=9, numfmt=NUM, align=S.right)
    write_assumption_docs(ws, R["other_rev"], dj, ds, dc, "drv_other_rev", D.JUST, D.ASSUMPTION_SRC, hints=D.SOURCE_HINT)
    write_equation_summary(ws, R["other_rev"], deq, "drv_f_other_rev_growth", D.SOURCE_HINT)

    for key, label, doc in [
        ("americas", "Americas net revenue ($000)", "drv_geo_americas"),
        ("china", "China Mainland net revenue ($000)", "drv_geo_china"),
        ("row", "Rest of World net revenue ($000)", "drv_geo_row"),
    ]:
        fy25v = kpis[f"revenue_geo_{key}"]["FY2025"]
        r = row(f"  {label}", fy25v, [None]*5, units="$000")
        R[f"geo_{key}"] = r
        for i, c in enumerate(FCOLS):
            write(ws, f"{c}{r}",
                  f"={FY25}{r}*({c}{R['store_rev']}+{c}{R['ecomm_rev']}+{c}{R['other_rev']})/"
                  f"({FY25}{R['store_rev']}+{FY25}{R['ecomm_rev']}+{FY25}{R['other_rev']})",
                  S.BLACK, size=9, numfmt=NUM, align=S.right)
        write_assumption_docs(ws, r, dj, ds, dc, doc, D.JUST, D.ASSUMPTION_SRC, hints=D.SOURCE_HINT)
        write_geo_scale_equation(ws, r, deq, R["store_rev"], R["ecomm_rev"], R["other_rev"],
                                 fy25_col=FY25, example_col=FCOLS[0])

    R["mix_w"] = row("Women's % of revenue", kpis["mix_women"]["FY2025"],
                     DK.FORECAST["mix_women"][1:], fmt=PCT, red=True, doc_key="drv_mix_women")
    R["mix_m"] = row("Men's % of revenue", kpis["mix_men"]["FY2025"],
                     DK.FORECAST["mix_men"][1:], fmt=PCT, red=True, doc_key="drv_mix_men")
    R["mix_a"] = row("Accessories & other % of revenue", kpis["mix_accessories"]["FY2025"],
                     DK.FORECAST["mix_accessories"][1:], fmt=PCT, red=True, doc_key="drv_mix_accessories")

    # --- CONSOLIDATED ---
    rr[0] += 1
    sub("V. CONSOLIDATED REVENUE & RECONCILIATION")
    col_hdr()
    R["total_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "Bottom-up total revenue ($000)", S.BLACK, bold=True, size=10, align=S.left_indent)
    write_reported(ws, f"{FY25}{rr[0]}", D.IS["revenue"]["FY2025"], D.filing_url("FY2025"),
                   bold=True, size=10, numfmt=NUM, align=S.right)
    for c in FCOLS:
        write(ws, f"{c}{rr[0]}",
              f"={c}{R['store_rev']}+{c}{R['ecomm_rev']}+{c}{R['other_rev']}",
              S.BLACK, bold=True, size=10, numfmt=NUM, align=S.right)
    f_docs(R["total_rev"], "drv_f_total_rev")
    rr[0] += 1

    R["scen_rev"] = rr[0]
    write(ws, f"A{rr[0]}", "Scenarios base-case revenue ($000)", S.BLACK, size=9, align=S.left_indent)
    write_reported(ws, f"{FY25}{rr[0]}", D.IS["revenue"]["FY2025"], D.filing_url("FY2025"),
                   size=9, numfmt=NUM, align=S.right)
    for y, c in zip(FY, FCOLS):
        write(ws, f"{c}{rr[0]}", f"=Scenarios!{scen_base_col}{rev_row_map[y]}",
              S.BLACK, size=9, numfmt=NUM, align=S.right)
    f_docs(R["scen_rev"], "drv_f_scen_rev", internal=f"'Scenarios'!{scen_base_col}{rev_row_map['FY2026E']}")
    rr[0] += 1

    R["variance"] = rr[0]
    write(ws, f"A{rr[0]}", "Variance (bottom-up − scenarios) ($000)", S.BLACK, size=9, align=S.left_indent)
    write(ws, f"{FY25}{rr[0]}", f"={FY25}{R['total_rev']}-{FY25}{R['scen_rev']}",
          S.BLACK, size=9, numfmt=NUM, align=S.right)
    for c in FCOLS:
        write(ws, f"{c}{rr[0]}", f"={c}{R['total_rev']}-{c}{R['scen_rev']}",
              S.BLACK, size=9, numfmt=NUM, align=S.right)
    f_docs(R["variance"], "drv_f_variance")
    rr[0] += 1

    R["var_pct"] = rr[0]
    write(ws, f"A{rr[0]}", "Variance (%)", S.BLACK, size=9, align=S.left_indent)
    for col in [FY25] + FCOLS:
        write(ws, f"{col}{rr[0]}",
              f"=IF({col}{R['scen_rev']}=0,\"\",{col}{R['variance']}/{col}{R['scen_rev']})",
              S.BLACK, size=9, numfmt=PCT, align=S.right)
    f_docs(R["var_pct"], "drv_f_var_pct")
    rr[0] += 1

    write(ws, f"A{rr[0]}",
          "Note: DCF / Scenarios consolidated revenue is the valuation anchor — drivers do not feed the DCF. "
          "A small variance vs Scenarios means the bottom-up path still hangs together, so the Scenarios "
          "case stays viable. Bigger gaps = revisit drivers or Scenarios assumptions.",
          S.BLACK, italic=True, size=8, align=S.left_indent)

    S.group_columns(ws, deq, dc)

    return ws, R
