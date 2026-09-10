"""Share Repurchase & EPS Accretion Schedule — below core DCF valuation (no IV linkage)."""
import styles as S
from styles import write, NUM, PCT, MONEY, EPSFMT
import data as D

FY25_COL = "E"
FCOLS = ["F", "G", "H", "I", "J"]
PREVF = {"F": FY25_COL, "G": "F", "H": "G", "I": "H", "J": "I"}
BUYBACK_FIXED_K = 750_000  # $750M/yr — matches 3-statement / Scenarios col G
FCF_ALLOC_SENSITIVITY = 0.75  # upside scenario only (bull CFF tab)
COE_GROWTH = 0.105


def build_repurchase_schedule(ws, r, DR, NP_R, VR):
    """
    Append repurchase / EPS accretion block to the DCF sheet.
    r: mutable single-element list [next_row] updated in place.
    DR: DCF forecast row map (needs ufcf, ebit).
    NP_R: NOPAT Bridge row map (add_lease, t_oper).
    VR: valuation row map (sh) — referenced in guardrail memo only.
    Returns RR row-index dict.
    """
    RR = {}

    def rp_row(key, label, e_val, proj_fn, fmt=NUM, color_e=S.BLACK, red=False, bold=False, top=False):
        RR[key] = r[0]
        row = r[0]
        bdr = S.top_border if top else None
        write(ws, f"A{row}", label, S.DARK if bold else S.BLACK, bold=bold, size=10, align=S.left_indent)
        if e_val is not None:
            col = S.RED if red else color_e
            write(ws, f"{FY25_COL}{row}", e_val, col, bold=bold, size=10,
                  numfmt=fmt, align=S.right, bdr=bdr)
        for c in FCOLS:
            col = S.RED if red else S.BLACK
            write(ws, f"{c}{row}", proj_fn(c), col, bold=bold, size=10,
                  numfmt=fmt, align=S.right, bdr=bdr)
        r[0] += 1

    r[0] += 2
    write(ws, f"A{r[0]}",
          "SHARE REPURCHASE & EPS ACCRETION SCHEDULE",
          S.WHITE, bold=True, size=10, fillc=S.DARK)
    for c in list("ABCDE"):
        ws[f"{c}{r[0]}"].fill = S.fill(S.DARK)
    r[0] += 1
    write(ws, f"A{r[0]}",
          "Guardrail: implied value per share above still divides equity value by Day-1 shares "
          f"(E{VR['sh']} = {D.MKT['shares_out']:,}k). This schedule does not change intrinsic DCF per share.",
          S.BLACK, italic=True, size=8, align=S.left_indent)
    r[0] += 2

    rp_row("buy_budget", "Annual repurchase budget ($000) — base case",
           BUYBACK_FIXED_K,
           lambda c: f"=${FY25_COL}${RR['buy_budget']}", fmt=NUM, red=True)
    rp_row("coe", "Cost of equity — share-price growth rate",
           COE_GROWTH,
           lambda c: f"=${FY25_COL}${RR['coe']}", fmt=PCT, red=True)
    r[0] += 1

    rp_row("ufcf", "Unlevered free cash flow (from DCF above)",
           None,
           lambda c: f"={c}{DR['ufcf']}")
    rp_row("buy_cash", "Repurchase cash deployed (= fixed budget)",
           0,
           lambda c: f"={c}{RR['buy_budget']}")
    rp_row("payout_pct", "Repurchase as % of UFCF (derived)",
           None,
           lambda c: f"={c}{RR['buy_cash']}/{c}{RR['ufcf']}", fmt=PCT)
    r[0] += 1

    rp_row("px", "Projected share price ($)",
           D.MKT["price"],
           lambda c: f"={PREVF[c]}{RR['px']}*(1+${FY25_COL}${RR['coe']})", fmt=MONEY)
    rp_row("retired", "Shares retired (000)",
           0,
           lambda c: f"={c}{RR['buy_cash']}/{c}{RR['px']}")

    RR["beg_sh"] = r[0]
    beg_row = r[0]
    write(ws, f"A{beg_row}", "Beginning shares (000)", S.BLACK, size=10, align=S.left_indent)
    write(ws, f"{FY25_COL}{beg_row}", D.MKT["shares_out"], S.BLUE, size=10, numfmt=NUM, align=S.right)
    write(ws, f"F{beg_row}", f"=${FY25_COL}${beg_row}", S.BLACK, size=10, numfmt=NUM, align=S.right)
    r[0] += 1

    RR["end_sh"] = r[0]
    end_row = r[0]
    write(ws, f"A{end_row}", "Ending shares (000)", S.DARK, bold=True, size=10, align=S.left_indent)
    write(ws, f"{FY25_COL}{end_row}", D.MKT["shares_out"], S.BLUE, bold=True, size=10,
          numfmt=NUM, align=S.right, bdr=S.top_border)
    write(ws, f"F{end_row}",
          f"=F{beg_row}-F{RR['retired']}", S.BLACK, bold=True, size=10,
          numfmt=NUM, align=S.right, bdr=S.top_border)
    for c in FCOLS[1:]:
        write(ws, f"{c}{beg_row}", f"={PREVF[c]}{end_row}", S.BLACK, size=10, numfmt=NUM, align=S.right)
        write(ws, f"{c}{end_row}",
              f"={c}{beg_row}-{c}{RR['retired']}", S.BLACK, bold=True, size=10,
              numfmt=NUM, align=S.right, bdr=S.top_border)
    r[0] += 1
    r[0] += 1

    fy25_ni = D.IS["net_income"]["FY2025"]
    fy25_eps = fy25_ni / D.MKT["shares_out"]
    rp_row("ni", "Projected net income ($000)",
           fy25_ni,
           lambda c: (f"=('NOPAT Bridge'!{c}{NP_R['ebit_norm']}-'NOPAT Bridge'!{c}{NP_R['add_lease']})"
                      f"*(1-'NOPAT Bridge'!{c}${NP_R['t_oper']})"),
           color_e=S.BLUE)
    rp_row("eps", "Accretive EPS ($/share) = Net income ÷ ending shares",
           fy25_eps,
           lambda c: f"={c}{RR['ni']}/{c}{RR['end_sh']}",
           fmt=EPSFMT, bold=True, top=True)

    return RR
