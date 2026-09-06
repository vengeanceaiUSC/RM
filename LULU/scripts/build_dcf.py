"""Builds LULU_DCF_Valuation_Model.xlsx from scratch (no third-party template).

Unlevered discounted-cash-flow valuation of lululemon athletica inc.
FY2025 (ended Feb 1, 2026) reported figures anchor the model (BLUE); the
five-year forecast is live formulas (BLACK) driven by analyst assumptions
(RED). Includes a WACC build, Gordon-growth and exit-multiple terminal value,
an EV -> equity bridge, a WACC x terminal-growth sensitivity grid, bull/base/
bear scenarios, and a comps-based football field.
Units: US$ thousands unless noted.
"""
import os
from openpyxl import Workbook
import styles as S
from styles import write, NUM, PCT, MONEY, MULT, EPSFMT
import data as D

OUT = os.path.join(os.path.dirname(__file__), "..", "LULU_DCF_Valuation_Model.xlsx")
wb = Workbook()

FY = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
FCOLS = ["D", "E", "F", "G", "H"]     # forecast columns
FCOL = dict(zip(FY, FCOLS))
BASE_REV = D.IS['revenue']['FY2025']  # 11,102,600 (blue)

# ------------------------------------------------------------------ COVER
cov = wb.active
cov.title = "Cover"
cov.sheet_view.showGridLines = False
S.set_col_widths(cov, {'A': 3, 'B': 100})
write(cov, 'B2', "GLOBAL INVESTMENT SOCIETY  |  INVESTMENT RESEARCH DIVISION", S.WHITE, bold=True, size=13, fillc=S.DARK)
cov['B2'].alignment = S.Alignment(horizontal='left', vertical='center', indent=1)
cov.row_dimensions[2].height = 26
write(cov, 'B4', "lululemon athletica inc. (NASDAQ: LULU)", S.DARK, bold=True, size=20)
write(cov, 'B5', "Discounted Cash Flow Valuation \u2014 Unlevered Free Cash Flow", S.ACCENT, bold=True, size=13)
write(cov, 'B7', "Recommendation:  LONG / OVERWEIGHT", S.GREEN, bold=True, size=14)
write(cov, 'B9', "FONT / COLOR CONVENTION", S.DARK, bold=True, size=12)
write(cov, 'B10', "Blue font  =  figures reported by the company (release, filing, or call)", S.BLUE, bold=True, size=11)
write(cov, 'B11', "Black font  =  calculations / formulas", S.BLACK, bold=True, size=11)
write(cov, 'B12', "Red font  =  analyst assumptions / inputs", S.RED, bold=True, size=11)
write(cov, 'B14', "TABS", S.DARK, bold=True, size=12)
write(cov, 'B15', "WACC  \u2022  DCF (base case + sensitivity)  \u2022  Scenarios  \u2022  Comps / Football Field", S.BLACK, size=10)
write(cov, 'B17', "SOURCES", S.DARK, bold=True, size=12)
write(cov, 'B18', "SEC EDGAR XBRL company facts, CIK 0001397187 (Form 10-K, FY2025 ended Feb 1, 2026).", S.BLACK, size=10)
write(cov, 'B19', "Market data & Q2 FY2026 results per company release dated Sep 3, 2026.", S.BLACK, size=10)
write(cov, 'B21', "Built from scratch for the GIS IR selection assignment.", S.BLACK, italic=True, size=9)

# ------------------------------------------------------------------ WACC
wacc = wb.create_sheet("WACC")
wacc.sheet_view.showGridLines = False
S.set_col_widths(wacc, {'A': 44, 'B': 2, 'C': 16, 'D': 40})
write(wacc, 'A1', "WEIGHTED AVERAGE COST OF CAPITAL", S.WHITE, bold=True, size=12, fillc=S.DARK)
for c in ['B', 'C', 'D']:
    wacc[f'{c}1'].fill = S.fill(S.DARK)
wacc.row_dimensions[1].height = 16
WR = {}
r = [3]
def w_row(key, label, value, color, fmt=PCT, note="", bold=False, top=False):
    WR[key] = r[0]
    bdr = S.top_border if top else None
    write(wacc, f'A{r[0]}', label, S.DARK if bold else S.BLACK, bold=bold, size=10, align=S.left_indent)
    if isinstance(value, str):
        write(wacc, f'C{r[0]}', value, S.BLACK, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    else:
        write(wacc, f'C{r[0]}', value, color, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    if note:
        write(wacc, f'D{r[0]}', note, S.BLACK, italic=True, size=8, align=S.left_indent)
    r[0] += 1

write(wacc, 'A2', "Cost of equity (CAPM)", S.ACCENT, bold=True, size=10)
w_row('rf', "Risk-free rate (10-yr UST)", 0.043, S.RED, note="analyst input")
w_row('erp', "Equity risk premium", 0.060, S.RED, note="analyst input")
w_row('beta', "Levered beta", 0.95, S.RED, fmt='0.00', note="~0.86 market beta; 0.95 used for near-term uncertainty")
w_row('coe', "Cost of equity = rf + \u03b2 \u00d7 ERP", f"=C{WR['rf']}+C{WR['beta']}*C{WR['erp']}", None, bold=True, top=True)
r[0] += 1
write(wacc, f'A{r[0]}', "Cost of debt", S.ACCENT, bold=True, size=10)
r[0] += 1
w_row('kd', "Pre-tax cost of debt", 0.050, S.RED, note="illustrative; LULU has no funded debt")
w_row('tax', "Tax rate", 0.270, S.RED, note="normalized marginal rate")
w_row('kdat', "After-tax cost of debt", f"=C{WR['kd']}*(1-C{WR['tax']})", None, top=True)
r[0] += 1
write(wacc, f'A{r[0]}', "Capital structure (market values)", S.ACCENT, bold=True, size=10)
r[0] += 1
w_row('we', "Equity weight", 1.00, S.RED, note="net-cash balance sheet \u2192 ~100% equity")
w_row('wd', "Debt weight", 0.00, S.RED)
w_row('wacc', "WACC", f"=C{WR['we']}*C{WR['coe']}+C{WR['wd']}*C{WR['kdat']}", None, bold=True, top=True)
wacc[f"C{WR['wacc']}"].font = S.font(color=S.GREEN, bold=True, size=12)
wacc[f"C{WR['wacc']}"].fill = S.fill(S.GREY)

def wref(key):
    return f"WACC!C{WR[key]}"

# ------------------------------------------------------------------ DCF (base)
dcf = wb.create_sheet("DCF")
dcf.sheet_view.showGridLines = False
S.set_col_widths(dcf, {'A': 42, 'B': 2, 'C': 13, 'D': 13, 'E': 13, 'F': 13, 'G': 13, 'H': 13})
write(dcf, 'A1', "DISCOUNTED CASH FLOW \u2014 BASE CASE  (US$ thousands)", S.WHITE, bold=True, size=12, fillc=S.DARK)
for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
    dcf[f'{c}1'].fill = S.fill(S.DARK)
write(dcf, 'C2', "FY2025A", S.BLUE, bold=True, size=10, align=S.center)
for y in FY:
    write(dcf, f'{FCOL[y]}2', y, S.WHITE, bold=True, size=10, align=S.center, fillc=S.ACCENT)
dcf.row_dimensions[1].height = 16

DR = {}
r = [4]
def d_row(key, label, cval, proj_fn, color_c=S.BLUE, color_p=S.BLACK, fmt=NUM, bold=False, top=False, dbl=False, red=False):
    DR[key] = r[0]
    bdr = S.top_double if dbl else (S.top_border if top else None)
    write(dcf, f'A{r[0]}', label, S.DARK if bold else S.BLACK, bold=bold, size=10, align=S.left_indent)
    if cval is not None:
        write(dcf, f'C{r[0]}', cval, color_c, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    for y in FY:
        c = FCOL[y]
        col = S.RED if red else color_p
        write(dcf, f'{c}{r[0]}', proj_fn(c), col, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    r[0] += 1

PREVF = {"D": "C", "E": "D", "F": "E", "G": "F", "H": "G"}

# Assumption rows (red)
d_row('growth', "Revenue growth %", None, lambda c: {"D": -0.061, "E": 0.010, "F": 0.025, "G": 0.030, "H": 0.030}[c], fmt=PCT, red=True)
d_row('rev', "Net revenue", BASE_REV, lambda c: f"={PREVF[c]}{DR['rev']}*(1+{c}{DR['growth']})", color_c=S.BLUE)
d_row('margin', "EBIT (operating) margin %", None, lambda c: {"D": 0.139, "E": 0.150, "F": 0.155, "G": 0.155, "H": 0.155}[c], fmt=PCT, red=True)
d_row('ebit', "EBIT", f"={D.IS['operating_income']['FY2025']}" if False else D.IS['operating_income']['FY2025'],
      lambda c: f"={c}{DR['rev']}*{c}{DR['margin']}", color_c=S.BLUE, bold=True, top=True)
d_row('taxes', "Less: cash taxes on EBIT", None, lambda c: f"=-{c}{DR['ebit']}*{wref('tax')}")
d_row('nopat', "NOPAT", None, lambda c: f"={c}{DR['ebit']}+{c}{DR['taxes']}", bold=True, top=True)
d_row('da_pct', "  D&A % of revenue", None, lambda c: 0.045, fmt=PCT, red=True)
d_row('da', "Plus: depreciation & amortization", None, lambda c: f"={c}{DR['rev']}*{c}{DR['da_pct']}")
d_row('capex_pct', "  Capex % of revenue", None, lambda c: 0.050, fmt=PCT, red=True)
d_row('capex', "Less: capital expenditures", None, lambda c: f"=-{c}{DR['rev']}*{c}{DR['capex_pct']}")
d_row('nwc_pct', "  NWC % of revenue", None, lambda c: 0.075, fmt=PCT, red=True)
d_row('dnwc', "Less: increase in net working capital", None,
      lambda c: f"=-{c}{DR['nwc_pct']}*({c}{DR['rev']}-{PREVF[c]}{DR['rev']})")
d_row('ufcf', "Unlevered free cash flow", None,
      lambda c: f"={c}{DR['nopat']}+{c}{DR['da']}+{c}{DR['capex']}+{c}{DR['dnwc']}", bold=True, top=True, dbl=True)
d_row('period', "Discount period (years)", None, lambda c: {"D": 1, "E": 2, "F": 3, "G": 4, "H": 5}[c], fmt='0')
d_row('df', "Discount factor @ WACC", None, lambda c: f"=1/(1+{wref('wacc')})^{c}{DR['period']}", fmt='0.000')
d_row('pv', "PV of unlevered FCF", None, lambda c: f"={c}{DR['ufcf']}*{c}{DR['df']}", bold=True, top=True)

r[0] += 1
# Valuation block
VR = {}
def v_row(key, label, formula, fmt=NUM, color=S.BLACK, bold=False, top=False, dbl=False, red=False):
    VR[key] = r[0]
    bdr = S.top_double if dbl else (S.top_border if top else None)
    write(dcf, f'A{r[0]}', label, S.DARK if bold else S.BLACK, bold=bold, size=10, align=S.left_indent)
    col = S.RED if red else color
    if isinstance(formula, (int, float)):
        write(dcf, f'C{r[0]}', formula, col, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    else:
        write(dcf, f'C{r[0]}', formula, col, bold=bold, size=10, numfmt=fmt, align=S.right, bdr=bdr)
    r[0] += 1

write(dcf, f'A{r[0]}', "VALUATION \u2014 GORDON GROWTH (PERPETUITY) METHOD", S.WHITE, bold=True, size=10, fillc=S.DARK)
for c in ['B', 'C']:
    dcf[f'{c}{r[0]}'].fill = S.fill(S.DARK)
r[0] += 1
v_row('sumpv', "Sum of PV of explicit FCF (FY26\u2013FY30)", f"=SUM(D{DR['pv']}:H{DR['pv']})", bold=True)
v_row('g', "Terminal growth rate (g)", 0.0225, fmt=PCT, red=True)
v_row('tv', "Terminal value = FCF\u2085\u00d7(1+g)/(WACC\u2212g)",
      f"=H{DR['ufcf']}*(1+C{VR['g']})/({wref('wacc')}-C{VR['g']})")
v_row('pvtv', "PV of terminal value", f"=C{VR['tv']}/(1+{wref('wacc')})^H{DR['period']}", bold=True)
v_row('ev', "Enterprise value", f"=C{VR['sumpv']}+C{VR['pvtv']}", bold=True, top=True)
v_row('cash', "Plus: cash & equivalents (FY2025)", D.MKT['cash'], color=S.BLUE)
v_row('debt', "Less: total debt", -D.MKT['debt'], color=S.BLUE)
v_row('eqv', "Equity value", f"=C{VR['ev']}+C{VR['cash']}+C{VR['debt']}", bold=True, top=True)
v_row('sh', "Diluted shares outstanding (000)", D.MKT['shares_out'], color=S.BLUE)
v_row('pt', "Implied value per share", f"=C{VR['eqv']}/C{VR['sh']}", fmt=MONEY, bold=True, top=True, dbl=True)
dcf[f"C{VR['pt']}"].font = S.font(color=S.GREEN, bold=True, size=13)
dcf[f"C{VR['pt']}"].fill = S.fill(S.GREY)
v_row('px', "Current share price", D.MKT['price'], fmt=MONEY, color=S.BLUE)
v_row('upside', "Implied upside / (downside)", f"=C{VR['pt']}/C{VR['px']}-1", fmt=PCT, bold=True)
dcf[f"C{VR['upside']}"].font = S.font(color=S.GREEN, bold=True, size=11)
v_row('tvpct', "  memo: % of EV from terminal value", f"=C{VR['pvtv']}/C{VR['ev']}", fmt=PCT)

r[0] += 1
write(dcf, f'A{r[0]}', "CROSS-CHECK \u2014 EXIT MULTIPLE METHOD", S.WHITE, bold=True, size=10, fillc=S.DARK)
for c in ['B', 'C']:
    dcf[f'{c}{r[0]}'].fill = S.fill(S.DARK)
r[0] += 1
v_row('ebitda', "Terminal EBITDA (FY2030E EBIT + D&A)", f"=H{DR['ebit']}+H{DR['da']}", bold=True)
v_row('exitm', "Exit EV/EBITDA multiple", 8.0, fmt=MULT, red=True)
v_row('tv2', "Terminal value (exit multiple)", f"=C{VR['ebitda']}*C{VR['exitm']}")
v_row('pvtv2', "PV of terminal value", f"=C{VR['tv2']}/(1+{wref('wacc')})^H{DR['period']}")
v_row('ev2', "Enterprise value (exit method)", f"=C{VR['sumpv']}+C{VR['pvtv2']}", bold=True, top=True)
v_row('pt2', "Implied value per share (exit method)",
      f"=(C{VR['ev2']}+C{VR['cash']}+C{VR['debt']})/C{VR['sh']}", fmt=MONEY, bold=True)
dcf[f"C{VR['pt2']}"].font = S.font(color=S.GREEN, bold=True, size=11)

# ------------------------------------------------------------------ SENSITIVITY (WACC x g)
r[0] += 2
write(dcf, f'A{r[0]}', "SENSITIVITY \u2014 IMPLIED SHARE PRICE (WACC vs terminal growth)", S.WHITE, bold=True, size=10, fillc=S.DARK)
for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
    dcf[f'{c}{r[0]}'].fill = S.fill(S.DARK)
r[0] += 1
sens_top = r[0]
waccs = [0.090, 0.095, 0.100, 0.105, 0.110]
gs = [0.015, 0.020, 0.0225, 0.025, 0.030]
write(dcf, f'A{sens_top}', "WACC \\ g", S.DARK, bold=True, size=9, align=S.center, fillc=S.LIGHT)
gcols = ['D', 'E', 'F', 'G', 'H']
for j, g in enumerate(gs):
    write(dcf, f'{gcols[j]}{sens_top}', g, S.RED, bold=True, size=9, numfmt=PCT, align=S.center, fillc=S.LIGHT)
fcf_rng = f"D{DR['ufcf']}:H{DR['ufcf']}"
lastfcf = f"H{DR['ufcf']}"
for i, wv in enumerate(waccs):
    rr = sens_top + 1 + i
    write(dcf, f'A{rr}', wv, S.RED, bold=True, size=9, numfmt=PCT, align=S.center, fillc=S.LIGHT)
    for j, g in enumerate(gs):
        # PV of explicit FCF via NPV at wv + PV of Gordon TV, +cash-debt, /shares
        f = (f"=(NPV({wv},{fcf_rng})"
             f"+({lastfcf}*(1+{g})/({wv}-{g}))/(1+{wv})^5"
             f"+C{VR['cash']}+C{VR['debt']})/C{VR['sh']}")
        col = S.GREEN if (abs(wv-0.10) < 1e-9 and abs(g-0.0225) < 1e-9) else S.BLACK
        write(dcf, f'{gcols[j]}{rr}', f, col, size=9, numfmt=MONEY, align=S.center)

# ------------------------------------------------------------------ SCENARIOS
scn = wb.create_sheet("Scenarios")
scn.sheet_view.showGridLines = False
S.set_col_widths(scn, {'A': 40, 'B': 2, 'C': 15, 'D': 15, 'E': 15})
write(scn, 'A1', "SCENARIO ANALYSIS", S.WHITE, bold=True, size=12, fillc=S.DARK)
for c in ['B', 'C', 'D', 'E']:
    scn[f'{c}1'].fill = S.fill(S.DARK)
scn.row_dimensions[1].height = 16
write(scn, 'C2', "Bear", S.WHITE, bold=True, size=11, align=S.center, fillc=S.ACCENT)
write(scn, 'D2', "Base", S.WHITE, bold=True, size=11, align=S.center, fillc=S.GREEN)
write(scn, 'E2', "Bull", S.WHITE, bold=True, size=11, align=S.center, fillc=S.DARK)

SC = {}
rr = [4]
def s_assum(key, label, bear, base, bull, fmt=PCT):
    SC[key] = rr[0]
    write(scn, f'A{rr[0]}', label, S.BLACK, size=10, align=S.left_indent)
    for col, v in zip(['C', 'D', 'E'], [bear, base, bull]):
        write(scn, f'{col}{rr[0]}', v, S.RED, size=10, numfmt=fmt, align=S.right)
    rr[0] += 1

write(scn, 'A3', "Key assumptions (5-yr forecast)", S.ACCENT, bold=True, size=10)
s_assum('g1', "FY2026E revenue growth", -0.090, -0.061, -0.040)
s_assum('gterm', "FY2028\u2013FY2030E revenue growth (avg)", -0.010, 0.028, 0.060)
s_assum('m1', "FY2026E EBIT margin", 0.125, 0.139, 0.150)
s_assum('mterm', "Terminal (FY2030E) EBIT margin", 0.120, 0.155, 0.190)
s_assum('wacc', "WACC", 0.110, 0.100, 0.090)
s_assum('g', "Terminal growth", 0.015, 0.0225, 0.030)

# Compact 5-year FCF per scenario (approximation using linear margin & growth paths)
rr[0] += 1
write(scn, f'A{rr[0]}', "Illustrative unlevered FCF & valuation", S.ACCENT, bold=True, size=10)
rr[0] += 1

def scen_block(col):
    """Write a compact 5-yr UFCF + valuation in the given scenario column using
    that column's assumptions. Revenue path: yr1 growth = g1; yrs2-5 linearly
    interpolate to gterm. Margin path: yr1=m1 linearly to mterm by yr5."""
    base = BASE_REV
    # We build helper rows in columns C/D/E starting at a fixed area to the right (cols G+).
    pass

# Simpler: compute scenario outputs with explicit per-year formulas in hidden helper columns.
# Build a mini FCF table per scenario stacked below.
rr[0] += 0
mini_top = rr[0]
labels = ["Revenue growth (path)", "Net revenue", "EBIT margin (path)", "EBIT", "NOPAT (@27% tax)",
          "+ D&A (4.5%)", "\u2212 Capex (5.0%)", "\u2212 \u0394NWC (7.5%)", "Unlevered FCF"]
# For each scenario column we lay years FY26..FY30 vertically? That is messy in 3 columns.
# Instead: present, per scenario, the resulting DCF outputs computed with closed-form using
# an average-growth / average-margin simplification, transparent as formulas.
for i, lab in enumerate([]):
    pass

# Outputs (transparent closed-form approximation):
# Revenue_t built off base with yr1=g1 then gterm for yrs2-5.
def rev_formula(col, t):
    if t == 1:
        return f"={BASE_REV}*(1+{col}{SC['g1']})"
    return f"={col}{mini_rev_row(t-1)}*(1+{col}{SC['gterm']})"

# Lay a per-scenario vertical block is complex; use a clean approach: 5 FCF rows computed
# with margin linearly from m1 (yr1) to mterm (yr5).
mrev = {}
mmar = {}
mebit = {}
mfcf = {}
def put(row, lab):
    write(scn, f'A{row}', lab, S.BLACK, size=9, align=S.left_indent)

# We'll create rows: for t in 1..5 -> revenue, ebit-margin, ufcf; then valuation.
cur = mini_top
rev_rows = {}
for t in range(1, 6):
    rev_rows[t] = cur
    put(cur, f"  Revenue \u2013 year {t}")
    for col in ['C', 'D', 'E']:
        if t == 1:
            f = f"={BASE_REV}*(1+{col}{SC['g1']})"
        else:
            f = f"={col}{rev_rows[t-1]}*(1+{col}{SC['gterm']})"
        write(scn, f'{col}{cur}', f, S.BLACK, size=9, numfmt=NUM, align=S.right)
    cur += 1
mar_rows = {}
for t in range(1, 6):
    mar_rows[t] = cur
    put(cur, f"  EBIT margin \u2013 year {t}")
    for col in ['C', 'D', 'E']:
        # linear from m1 (t=1) to mterm (t=5)
        f = f"={col}{SC['m1']}+({col}{SC['mterm']}-{col}{SC['m1']})*{(t-1)}/4"
        write(scn, f'{col}{cur}', f, S.BLACK, size=9, numfmt=PCT, align=S.right)
    cur += 1
fcf_rows = {}
for t in range(1, 6):
    fcf_rows[t] = cur
    put(cur, f"  Unlevered FCF \u2013 year {t}")
    for col in ['C', 'D', 'E']:
        prev_rev = f"{col}{rev_rows[t-1]}" if t > 1 else str(BASE_REV)
        ebit = f"{col}{rev_rows[t]}*{col}{mar_rows[t]}"
        nopat = f"({ebit})*(1-0.27)"
        da = f"{col}{rev_rows[t]}*0.045"
        capex = f"{col}{rev_rows[t]}*0.05"
        dnwc = f"0.075*({col}{rev_rows[t]}-{prev_rev})"
        f = f"={nopat}+{da}-{capex}-{dnwc}"
        write(scn, f'{col}{cur}', f, S.BLACK, size=9, numfmt=NUM, align=S.right)
    cur += 1

cur += 1
write(scn, f'A{cur}', "Enterprise value (DCF)", S.DARK, bold=True, size=10, align=S.left_indent)
ev_row = cur
for col in ['C', 'D', 'E']:
    rng = f"{col}{fcf_rows[1]}"
    fcf_cells = ",".join(f"{col}{fcf_rows[t]}" for t in range(1, 6))
    lastf = f"{col}{fcf_rows[5]}"
    f = (f"=NPV({col}{SC['wacc']},{fcf_cells})"
         f"+({lastf}*(1+{col}{SC['g']})/({col}{SC['wacc']}-{col}{SC['g']}))/(1+{col}{SC['wacc']})^5")
    write(scn, f'{col}{cur}', f, S.BLACK, bold=True, size=10, numfmt=NUM, align=S.right, bdr=S.top_border)
cur += 1
write(scn, f'A{cur}', "Implied share price", S.DARK, bold=True, size=11, align=S.left_indent)
pt_row = cur
for col in ['C', 'D', 'E']:
    f = f"=({col}{ev_row}+{D.MKT['cash']}-{D.MKT['debt']})/{D.MKT['shares_out']}"
    color = S.GREEN if col == 'D' else S.BLACK
    write(scn, f'{col}{cur}', f, color, bold=True, size=12, numfmt=MONEY, align=S.right, bdr=S.top_double, fillc=S.GREY)
cur += 1
write(scn, f'A{cur}', "Upside / (downside) vs current", S.DARK, bold=True, size=10, align=S.left_indent)
for col in ['C', 'D', 'E']:
    f = f"={col}{pt_row}/{D.MKT['price']}-1"
    write(scn, f'{col}{cur}', f, S.BLACK, bold=True, size=10, numfmt=PCT, align=S.right)
cur += 2
write(scn, f'A{cur}', "Current price $%.2f; cash $%s k; net debt $0 (net-cash balance sheet)." % (D.MKT['price'], f"{D.MKT['cash']:,}"),
      S.BLACK, italic=True, size=8, align=S.left_indent)

# ------------------------------------------------------------------ COMPS / FOOTBALL FIELD
comps = wb.create_sheet("Comps")
comps.sheet_view.showGridLines = False
S.set_col_widths(comps, {'A': 34, 'B': 2, 'C': 14, 'D': 14, 'E': 14, 'F': 14})
write(comps, 'A1', "RELATIVE VALUATION \u2014 IMPLIED PRICE RANGES (FOOTBALL FIELD)", S.WHITE, bold=True, size=12, fillc=S.DARK)
for c in ['B', 'C', 'D', 'E', 'F']:
    comps[f'{c}1'].fill = S.fill(S.DARK)
comps.row_dimensions[1].height = 16
write(comps, 'A3', "LULU operating metrics (FY2025A / FY2026E)", S.ACCENT, bold=True, size=10)
CM = {}
rr = [4]
def c_row(key, label, val, color=S.BLUE, fmt=NUM):
    CM[key] = rr[0]
    write(comps, f'A{rr[0]}', label, S.BLACK, size=10, align=S.left_indent)
    write(comps, f'C{rr[0]}', val, color, size=10, numfmt=fmt, align=S.right)
    rr[0] += 1

c_row('rev', "FY2025A revenue", D.IS['revenue']['FY2025'])
c_row('ebitda', "FY2025A EBITDA (EBIT + D&A)",
      D.IS['operating_income']['FY2025'] + D.CF['d_and_a']['FY2025'])
comps[f"C{CM['ebitda']}"].value = f"={D.IS['operating_income']['FY2025']}+{D.CF['d_and_a']['FY2025']}"
comps[f"C{CM['ebitda']}"].font = S.font(color=S.BLACK)
c_row('eps26', "FY2026E diluted EPS (guidance midpoint)", 9.61, color=S.BLUE, fmt=EPSFMT)
c_row('cash', "Cash & equivalents", D.MKT['cash'])
c_row('sh', "Diluted shares (000)", D.MKT['shares_out'])
c_row('px', "Current share price", D.MKT['price'], fmt=MONEY)
# current multiples (black formulas)
CM['ceved'] = rr[0]
write(comps, f'A{rr[0]}', "  memo: current EV / EBITDA", S.BLACK, italic=True, size=9, align=S.left_indent)
write(comps, f'C{rr[0]}', f"=(C{CM['px']}*C{CM['sh']}-C{CM['cash']})/C{CM['ebitda']}", S.BLACK, italic=True, size=9, numfmt=MULT, align=S.right)
rr[0] += 1
write(comps, f'A{rr[0]}', "  memo: current P / E (FY2026E)", S.BLACK, italic=True, size=9, align=S.left_indent)
write(comps, f'C{rr[0]}', f"=C{CM['px']}/C{CM['eps26']}", S.BLACK, italic=True, size=9, numfmt=MULT, align=S.right)
rr[0] += 2
write(comps, f'A{rr[0]}', "Methodology / multiple ranges", S.ACCENT, bold=True, size=10)
rr[0] += 1
write(comps, f'A{rr[0]}', "Method", S.WHITE, bold=True, size=10, fillc=S.DARK, align=S.left_indent)
write(comps, f'C{rr[0]}', "Low mult.", S.WHITE, bold=True, size=10, fillc=S.DARK, align=S.center)
write(comps, f'D{rr[0]}', "High mult.", S.WHITE, bold=True, size=10, fillc=S.DARK, align=S.center)
write(comps, f'E{rr[0]}', "Implied px (low)", S.WHITE, bold=True, size=10, fillc=S.DARK, align=S.center)
write(comps, f'F{rr[0]}', "Implied px (high)", S.WHITE, bold=True, size=10, fillc=S.DARK, align=S.center)
rr[0] += 1

def ff_ev_ebitda(label, lo, hi):
    write(comps, f'A{rr[0]}', label, S.BLACK, size=10, align=S.left_indent)
    write(comps, f'C{rr[0]}', lo, S.RED, size=10, numfmt=MULT, align=S.center)
    write(comps, f'D{rr[0]}', hi, S.RED, size=10, numfmt=MULT, align=S.center)
    for outcol, mcol in [('E', 'C'), ('F', 'D')]:
        f = f"=(C{CM['ebitda']}*{mcol}{rr[0]}+C{CM['cash']})/C{CM['sh']}"
        write(comps, f'{outcol}{rr[0]}', f, S.BLACK, size=10, numfmt=MONEY, align=S.center)
    rr[0] += 1

def ff_pe(label, lo, hi):
    write(comps, f'A{rr[0]}', label, S.BLACK, size=10, align=S.left_indent)
    write(comps, f'C{rr[0]}', lo, S.RED, size=10, numfmt=MULT, align=S.center)
    write(comps, f'D{rr[0]}', hi, S.RED, size=10, numfmt=MULT, align=S.center)
    for outcol, mcol in [('E', 'C'), ('F', 'D')]:
        f = f"=C{CM['eps26']}*{mcol}{rr[0]}"
        write(comps, f'{outcol}{rr[0]}', f, S.BLACK, size=10, numfmt=MONEY, align=S.center)
    rr[0] += 1

ff_ev_ebitda("EV / EBITDA (FY2025A)", 4.5, 7.5)
ff_pe("P / E (FY2026E EPS)", 10.0, 18.0)
# DCF range row references DCF sheet outputs
write(comps, f'A{rr[0]}', "DCF (bear \u2013 bull)", S.BLACK, size=10, align=S.left_indent)
write(comps, f'C{rr[0]}', "\u2014", S.BLACK, size=10, align=S.center)
write(comps, f'D{rr[0]}', "\u2014", S.BLACK, size=10, align=S.center)
write(comps, f'E{rr[0]}', f"=Scenarios!C{pt_row}", S.BLACK, size=10, numfmt=MONEY, align=S.center)
write(comps, f'F{rr[0]}', f"=Scenarios!E{pt_row}", S.BLACK, size=10, numfmt=MONEY, align=S.center)
rr[0] += 2
write(comps, f'A{rr[0]}', "Current price", S.BLACK, bold=True, size=11, align=S.left_indent)
write(comps, f'C{rr[0]}', D.MKT['price'], S.BLUE, bold=True, size=11, numfmt=MONEY, align=S.center)
rr[0] += 1
write(comps, f'A{rr[0]}', "Note: peer set includes NKE, DECK, ONON, adidas, VFC; multiples are analyst ranges (red).",
      S.BLACK, italic=True, size=8, align=S.left_indent)

wb.save(OUT)
print("Saved", os.path.abspath(OUT))
