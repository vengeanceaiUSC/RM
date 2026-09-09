"""Build NOPAT Bridge tab — 5-phase reported EBIT → normalized NOPAT (base case)."""
from openpyxl.styles import Alignment
import styles as S
from styles import (
    write, write_reported, write_assumption_docs, group_columns, group_rows,
    NUM, PCT,
)
import data as D

FY = D.PROJ_YEARS
FY25_COL = "E"
FCOLS = ["F", "G", "H", "I", "J"]
FCOL = dict(zip(FY, FCOLS))
DJ, DS, DC = "B", "C", "D"
NP = D.NOPAT_FY25
YEAR_IDX = {y: i for i, y in enumerate(FY)}


def build_nopat_bridge(wb, scen_base_col, rev_rows, ebit_rows, drv, sc_tariff_cell, sc_tax_cell):
    """
    Create NOPAT Bridge worksheet (base case cols F–J).
    rev_rows / ebit_rows: Scenarios forecast row indices by year 1–5.
    drv: Revenue Drivers row map (store_rev, ecomm_rev, other_rev).
    sc_tariff_cell: e.g. 'G12' — Scenarios base-case FY26 tariff cell.
    Returns row-index dict with keys nopat, ebit_norm, t_oper, etc.
    """
    ws = wb.create_sheet("NOPAT Bridge")
    ws.sheet_view.showGridLines = False
    S.set_col_widths(ws, {
        "A": 52, "B": 22, "C": 28, "D": 28,
        "E": 12, "F": 14, "G": 11, "H": 11, "I": 11, "J": 11,
    })
    write(ws, "A1", "NOPAT NORMALIZATION PIPELINE (BASE CASE)", S.WHITE, bold=True, size=12, fillc=S.DARK)
    for c in list("BCDEFGHIJ"):
        ws[f"{c}1"].fill = S.fill(S.DARK)
    write(ws, f"{DJ}2", "Justification  [cols B\u2013D: click + to expand]", S.ACCENT, bold=True, size=9)
    write(ws, f"{DS}2", "Source (click)", S.ACCENT, bold=True, size=9)
    write(ws, f"{DC}2", "Ctrl+F (prove number)", S.ACCENT, bold=True, size=9)
    write(ws, f"{FY25_COL}2", "FY2025A", S.BLUE, bold=True, size=10, align=S.center)
    for y in FY:
        write(ws, f"{FCOL[y]}2", y, S.WHITE, bold=True, size=10, align=S.center, fillc=S.ACCENT)
    ws.freeze_panes = f"{FY25_COL}3"

    R = {}
    rr = [3]

    def hdr(title):
        write(ws, f"A{rr[0]}", title, S.ACCENT, bold=True, size=10, align=S.left_indent)
        rr[0] += 1

    def sub(title):
        write(ws, f"A{rr[0]}", title, S.DARK, bold=True, size=9, align=S.left_indent)
        rr[0] += 1

    def line(key, label, fy25_val, proj_fn, color_fy25=S.BLUE, fmt=NUM, bold=False, top=False,
             doc_key=None, internal_location=None):
        R[key] = rr[0]
        row = rr[0]
        bdr = S.top_border if top else None
        write(ws, f"A{row}", label, S.DARK if bold else S.BLACK, bold=bold, size=10, align=S.left_indent)
        if fy25_val is not None:
            if isinstance(fy25_val, (int, float)) and color_fy25 == S.BLUE:
                write_reported(ws, f"{FY25_COL}{row}", fy25_val, D.filing_url("FY2025"),
                               bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
            else:
                write(ws, f"{FY25_COL}{row}", fy25_val, color_fy25, bold=bold, size=10,
                      numfmt=fmt, align=S.right, bdr=bdr)
        for y in FY:
            c = FCOL[y]
            write(ws, f"{c}{row}", proj_fn(c, y), S.BLACK, bold=bold, size=10,
                  numfmt=fmt, align=S.right, bdr=bdr)
        if doc_key:
            write_assumption_docs(ws, row, DJ, DS, DC, doc_key, D.JUST, D.ASSUMPTION_SRC,
                                  internal_location=internal_location, hints=D.SOURCE_HINT)
        rr[0] += 1

    def assum(key, label, value, fmt=PCT, doc_key=None):
        R[key] = rr[0]
        row = rr[0]
        write(ws, f"A{row}", label, S.BLACK, size=10, align=S.left_indent)
        write(ws, f"{FY25_COL}{row}", value, S.RED, size=10, numfmt=fmt, align=S.right)
        for c in FCOLS:
            write(ws, f"{c}{row}", value, S.RED, size=10, numfmt=fmt, align=S.right)
        if doc_key:
            write_assumption_docs(ws, row, DJ, DS, DC, doc_key, D.JUST, D.ASSUMPTION_SRC,
                                  hints=D.SOURCE_HINT)
        rr[0] += 1

    def scen_ref(row_dict, year_label):
        return f"Scenarios!{scen_base_col}{row_dict[YEAR_IDX[year_label] + 1]}"

    hdr("Reported EBIT \u2192 Normalized NOPAT (5-phase pipeline)")
    write(ws, f"A{rr[0]}",
          "Phases 1\u20134 clean and forecast operating profit; Phase 5 is the agent workflow summary below.",
          S.BLACK, italic=True, size=8, align=S.left_indent)
    rr[0] += 2

    hdr("Driver assumptions")
    assum("sbc_flag",
          "Phase 1: Add back SBC? (0 = no — Convention A: expense stays in EBIT)",
          0, fmt="0", doc_key="np_sbc")
    write(ws, f"A{rr[0]}",
          "Convention A: SBC is a real economic cost (captures shareholder dilution). "
          "Forecast NOPAT uses Scenarios EBIT with SBC expensed. DCF implied value uses basic shares (111.4M).",
          S.BLACK, italic=True, size=8, align=S.left_indent)
    rr[0] += 1
    assum("m_store", "Phase 3: Store-channel EBIT margin %", NP["margin_store"], doc_key="np_m_store")
    assum("m_ecomm", "Phase 3: E-commerce EBIT margin %", NP["margin_ecomm"], doc_key="np_m_ecomm")
    assum("m_other", "Phase 3: Other-channels EBIT margin %", NP["margin_other"], doc_key="np_m_other")
    assum("t_marg", "Phase 4: Terminal marginal tax rate", 0.30, doc_key="np_t_marg")
    assum("rd_years", "Phase 2: R&D / software amortization period (years)", 4, fmt="0", doc_key="np_rd_years")
    rr[0] += 1

    p1_start = rr[0]
    sub("Phase 1 — Operating normalization")
    line("ebit_rep", "Reported operating income (EBIT)", NP["ebit_reported"],
         lambda c, y: f"={scen_ref(ebit_rows, y)}", doc_key="np_ebit_rep")
    line("add_impair", "+ Impairment / intangible amortization (add-back)", NP["impairment_amort"],
         lambda c, y: 0, doc_key="np_impair")
    line("add_restruct", "+ Restructuring costs (add-back)", NP["restructuring"],
         lambda c, y: 0, doc_key="np_restruct")
    line("add_legal", "+ Legal / M&A one-offs (add-back)", NP["legal_ma"],
         lambda c, y: 0, doc_key="np_legal")
    line("add_sbc", "+ Stock-based compensation (add-back only if flag = 1)", None,
         lambda c, y: (f"={scen_ref(rev_rows, y)}/{NP['revenue']}*"
                       f"{NP['sbc']}*${FY25_COL}${R['sbc_flag']}"),
         doc_key="np_sbc")
    write(ws, f"{FY25_COL}{R['add_sbc']}",
          f"={NP['sbc']}*${FY25_COL}${R['sbc_flag']}", S.BLACK, size=10, numfmt=NUM, align=S.right)
    line("ebit_p1", "= Adjusted EBIT (Phase 1)", None,
         lambda c, y: (f"={c}{R['ebit_rep']}+{c}{R['add_impair']}+{c}{R['add_restruct']}"
                       f"+{c}{R['add_legal']}+{c}{R['add_sbc']}"),
         bold=True, top=True, doc_key="np_ebit_p1")
    p1_end = rr[0] - 1

    p2_start = rr[0]
    sub("Phase 2 — Lease & capitalization (ASC 842 / IFRS 16)")
    line("add_lease", "+ Implied lease interest expense (reclass from rent)", NP["lease_interest"],
         lambda c, y: f"={NP['lease_interest']}*({scen_ref(rev_rows, y)}/{NP['revenue']})",
         doc_key="np_lease_int")
    line("add_rd", "+ Capitalize R&D / software (current-year expense)", NP["rd_expense_cap"],
         lambda c, y: 0, doc_key="np_rd_cap")
    line("less_amort", "\u2212 Amortization of prior capitalized intangibles", NP["intangible_amort"],
         lambda c, y: 0, doc_key="np_rd_amort")
    line("ebit_p2", "= EBIT after lease & capitalization (Phase 2)", None,
         lambda c, y: f"={c}{R['ebit_p1']}+{c}{R['add_lease']}+{c}{R['add_rd']}-{c}{R['less_amort']}",
         bold=True, top=True, doc_key="np_ebit_p2")
    p2_end = rr[0] - 1

    p3_start = rr[0]
    sub("Phase 3 — Segment / channel EBIT (driver-based)")
    line("rev_st", "  Store-channel revenue", NP["rev_store"],
         lambda c, y: f"='Revenue Drivers'!{c}{drv['store_rev']}", doc_key="np_rev_store",
         internal_location=f"'Revenue Drivers'!B{drv['store_rev']}")
    line("rev_ec", "  E-commerce revenue", NP["rev_ecomm"],
         lambda c, y: f"='Revenue Drivers'!{c}{drv['ecomm_rev']}", doc_key="np_rev_ecomm",
         internal_location=f"'Revenue Drivers'!B{drv['ecomm_rev']}")
    line("rev_ot", "  Other channels revenue", NP["rev_other"],
         lambda c, y: f"='Revenue Drivers'!{c}{drv['other_rev']}", doc_key="np_rev_other",
         internal_location=f"'Revenue Drivers'!B{drv['other_rev']}")
    line("ebit_st", "  Store-channel EBIT (= Rev × store margin)", NP["rev_store"] * NP["margin_store"],
         lambda c, y: f"={c}{R['rev_st']}*{FY25_COL}${R['m_store']}",
         color_fy25=S.BLACK, doc_key="np_ebit_store")
    line("ebit_ec", "  E-commerce EBIT (= Rev × e-comm margin)", NP["rev_ecomm"] * NP["margin_ecomm"],
         lambda c, y: f"={c}{R['rev_ec']}*{FY25_COL}${R['m_ecomm']}",
         color_fy25=S.BLACK, doc_key="np_ebit_ecomm")
    line("ebit_ot", "  Other-channels EBIT (= Rev × other margin)", NP["rev_other"] * NP["margin_other"],
         lambda c, y: f"={c}{R['rev_ot']}*{FY25_COL}${R['m_other']}",
         color_fy25=S.BLACK, doc_key="np_ebit_other")
    line("ebit_ch", "  Channel EBIT (sum of sector EBIT)", None,
         lambda c, y: (f"={c}{R['ebit_st']}+{c}{R['ebit_ec']}+{c}{R['ebit_ot']}"
                       + (f"+{sc_tariff_cell}" if y == "FY2026E" else "")),
         bold=True, doc_key="np_ebit_channel")
    fy25_ch = (NP["rev_store"] * NP["margin_store"] + NP["rev_ecomm"] * NP["margin_ecomm"]
               + NP["rev_other"] * NP["margin_other"])
    line("ebit_p3", "= EBIT after channel mix (Phase 3)", fy25_ch,
         lambda c, y: f"={c}{R['ebit_ch']}", bold=True, top=True, doc_key="np_ebit_p3")
    line("ebit_norm", "= Normalized EBIT (tax base for NOPAT)", None,
         lambda c, y: f"={scen_ref(ebit_rows, y)}",
         bold=True, top=True, doc_key="np_ebit_norm")
    write(ws, f"{FY25_COL}{R['ebit_norm']}",
          (f"={FY25_COL}{R['ebit_p1']}+{FY25_COL}{R['add_lease']}"
           f"+{FY25_COL}{R['add_rd']}-{FY25_COL}{R['less_amort']}"),
          S.BLACK, bold=True, size=10, numfmt=NUM, align=S.right, bdr=S.top_border)
    line("chk_scen", "  Check vs Scenarios EBIT (base case)", None,
         lambda c, y: f"={c}{R['ebit_norm']}-{scen_ref(ebit_rows, y)}", fmt=NUM)
    p3_end = rr[0] - 1

    p4_start = rr[0]
    sub("Phase 4 — Tax normalization to NOPAT")
    line("t_oper", "Operating effective tax rate (t_operating)", NP["t_operating"],
         lambda c, y: (f"={NP['t_operating']}+({c}${R['t_marg']}-{NP['t_operating']})*"
                       f"{YEAR_IDX[y] + 1}/5"),
         fmt=PCT, doc_key="np_t_oper")
    line("tax_exp", "Unlevered tax on normalized EBIT", None,
         lambda c, y: f"={c}{R['ebit_norm']}*{c}{R['t_oper']}", fmt=NUM, doc_key="np_tax_exp")
    line("nopat", "NORMALIZED NOPAT", None,
         lambda c, y: f"={c}{R['ebit_norm']}-{c}{R['tax_exp']}",
         bold=True, top=True, doc_key="np_nopat")
    p4_end = rr[0] - 1

    rr[0] += 1
    sub("Phase 5 — Agent workflow (execution steps)")
    write(ws, f"B{rr[0]}", "Source", S.ACCENT, bold=True, size=8, align=S.left_indent)
    write(ws, f"C{rr[0]}", "Without applied (FY26 DCF)", S.ACCENT, bold=True, size=8, align=S.left_indent)
    write(ws, f"D{rr[0]}", "With applied (FY26 DCF)", S.ACCENT, bold=True, size=8, align=S.left_indent)
    write(ws, f"F{rr[0]}", "Output", S.ACCENT, bold=True, size=8, align=S.left_indent)
    rr[0] += 1
    workflow = [
        ("1. Clean base", "SEC 10-K income statement",
         "EBIT = 13.2% × rev + tariff; SBC stays expensed (flag = 0).",
         "DCF unchanged.",
         "EBIT_adj (Phase 1)"),
        ("2. Capitalize R&D", "SG&A footnotes",
         "R&D/software expensed in full each year ($0 for LULU).",
         "DCF unchanged.",
         "EBIT_capitalized (Phase 2)"),
        ("3. Lease shift", "ASC 842 lease footnote",
         "Rent in opex; no implied lease-interest add-back.",
         "DCF unchanged.",
         "EBIT_lease-adj (Phase 2)"),
        ("4. Segment mix", "Revenue Drivers tab",
         "DCF uses consolidated 13.2% on total revenue (+ tariff).",
         "DCF unchanged.",
         "Forecast EBIT (Phase 3)"),
        ("5. NOPAT bridge", "Normalized EBIT & t_operating",
         "NOPAT = EBIT × (1 − 30%) flat sc_tax.",
         "NOPAT = EBIT × (1 − t_operating) — same EBIT, slightly higher NOPAT.",
         "Normalized NOPAT"),
        ("6. Tariff refund (Scenarios)", "Q2 FY2026 earnings release",
         "EBIT = Rev × 13.2% only (no refund).",
         "+ $134,500k in FY26 — DCF uses this.",
         "Scenarios EBIT (FY26)"),
    ]
    for step, inp, without, with_applied, out in workflow:
        write(ws, f"A{rr[0]}", step, S.BLACK, bold=True, size=9, align=S.left_indent)
        write(ws, f"B{rr[0]}", inp, S.BLACK, size=8, align=S.left_indent)
        write(ws, f"C{rr[0]}", without, S.BLACK, size=8, align=S.left_indent)
        write(ws, f"D{rr[0]}", with_applied, S.BLACK, size=8, align=S.left_indent)
        write(ws, f"F{rr[0]}", out, S.ACCENT, size=8, align=S.left_indent)
        for col in (f"B{rr[0]}", f"C{rr[0]}", f"D{rr[0]}"):
            ws[col].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        rr[0] += 1

    write(ws, f"A{rr[0]}",
          "Steps 1–4: DCF unchanged on FY26. Step 5 changes NOPAT only (tax). Step 6 changes FY26 EBIT (+ tariff).",
          S.BLACK, italic=True, size=8, align=S.left_indent)
    rr[0] += 1
    write(ws, f"A{rr[0]}",
          "Scenarios base-case NOPAT (col G) links to NORMALIZED NOPAT above. "
          "Bear/bull apply the same t_operating to their EBIT paths.",
          S.BLACK, italic=True, size=8, align=S.left_indent)

    group_columns(ws, DJ, DC)
    group_rows(ws, p1_start, p1_end, hidden=True)
    group_rows(ws, p2_start, p2_end, hidden=True)
    group_rows(ws, p3_start, p3_end, hidden=True)
    group_rows(ws, p4_start, p4_end, hidden=True)

    return ws, R
