"""
LULU financial data sourced from SEC EDGAR XBRL company facts (CIK 0001397187).
All figures in USD thousands unless noted. These are the "reported" (blue-font)
values that anchor the models. Fiscal-year convention follows lululemon's own
labeling (fiscal year is named for the calendar year it mostly covers):

    FY2022  -> 52 weeks ended Jan 29, 2023
    FY2023  -> 52 weeks ended Jan 28, 2024
    FY2024  -> 53 weeks ended Feb  2, 2025
    FY2025  -> 52 weeks ended Feb  1, 2026   (most recent 10-K)

Source filings: LULU Forms 10-K (accession 0001397187-26-000020 and prior),
retrieved from https://data.sec.gov/api/xbrl/companyfacts/CIK0001397187.json
Q2 FY2026 actuals / FY2026 guidance from the Sep 3, 2026 earnings release.
"""

# Historical fiscal years used as reported columns
HIST_YEARS = ["FY2022", "FY2023", "FY2024", "FY2025"]
HIST_END = {
    "FY2022": "Jan 29, 2023",
    "FY2023": "Jan 28, 2024",
    "FY2024": "Feb 2, 2025",
    "FY2025": "Feb 1, 2026",
}

PROJ_YEARS = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]

# ---------------------------------------------------------------------------
# INCOME STATEMENT (USD thousands) -- reported
# ---------------------------------------------------------------------------
IS = {
    "revenue":        {"FY2022": 8110518, "FY2023": 9619278, "FY2024": 10588126, "FY2025": 11102600},
    "cogs":           {"FY2022": 3618178, "FY2023": 4009873, "FY2024": 4317315,  "FY2025": 4818468},
    "gross_profit":   {"FY2022": 4492340, "FY2023": 5609405, "FY2024": 6270811,  "FY2025": 6284132},
    "sga":            {"FY2022": 2757447, "FY2023": 3397218, "FY2024": 3762379,  "FY2025": 4066556},
    # Impairment / amortization of intangibles / other operating (GP - SG&A - OpInc)
    "other_opex":     {"FY2022": 406485,  "FY2023": 79511,   "FY2024": 2735,     "FY2025": 6961},
    "operating_income": {"FY2022": 1328408, "FY2023": 2132676, "FY2024": 2505697, "FY2025": 2210615},
    # Net other income (mostly interest income), = pretax - operating income
    "other_income":   {"FY2022": 4163,    "FY2023": 43059,   "FY2024": 70380,    "FY2025": 28352},
    "pretax_income":  {"FY2022": 1332571, "FY2023": 2175735, "FY2024": 2576077,  "FY2025": 2238967},
    "tax":            {"FY2022": 477771,  "FY2023": 625545,  "FY2024": 761461,   "FY2025": 659784},
    "net_income":     {"FY2022": 854800,  "FY2023": 1550190, "FY2024": 1814616,  "FY2025": 1579183},
    "diluted_shares": {"FY2022": 128017,  "FY2023": 127060,  "FY2024": 123935,   "FY2025": 119068},
    "diluted_eps":    {"FY2022": 6.68,    "FY2023": 12.20,   "FY2024": 14.64,    "FY2025": 13.26},
}

# ---------------------------------------------------------------------------
# BALANCE SHEET (USD thousands) -- reported, at fiscal year end
# ---------------------------------------------------------------------------
BS = {
    "cash":            {"FY2022": 1154867, "FY2023": 2243971, "FY2024": 1984336, "FY2025": 1807202},
    "ar":              {"FY2022": 132906,  "FY2023": 124769,  "FY2024": 120173,  "FY2025": 190657},
    "inventories":     {"FY2022": 1447367, "FY2023": 1323602, "FY2024": 1442081, "FY2025": 1700753},
    "current_assets":  {"FY2022": 3159453, "FY2023": 4060577, "FY2024": 3980302, "FY2025": 4262701},
    "ppe_net":         {"FY2022": 1269614, "FY2023": 1545811, "FY2024": 1780617, "FY2025": 2033720},
    "rou_asset":       {"FY2022": 969419,  "FY2023": 1265610, "FY2024": 1416256, "FY2025": 1630181},
    "goodwill_intang": {"FY2022": 46105,   "FY2023": 24083,   "FY2024": 171191,  "FY2025": 191194},
    "total_assets":    {"FY2022": 5607038, "FY2023": 7091941, "FY2024": 7603292, "FY2025": 8456743},

    "accounts_payable":{"FY2022": 172732,  "FY2023": 348441,  "FY2024": 271406,  "FY2025": 331421},
    "accrued_liab":    {"FY2022": 399223,  "FY2023": 348555,  "FY2024": 559463,  "FY2025": 662982},
    "op_lease_cur":    {"FY2022": 207972,  "FY2023": 249270,  "FY2024": 275154,  "FY2025": 298724},
    "current_liab":    {"FY2022": 1492198, "FY2023": 1631261, "FY2024": 1839630, "FY2025": 1887548},
    "op_lease_noncur": {"FY2022": 862362,  "FY2023": 1154012, "FY2024": 1300637, "FY2025": 1499717},
    "deferred_tax":    {"FY2022": 55084,   "FY2023": 29522,   "FY2024": 98188,   "FY2025": 52278},
    "total_liab":      {"FY2022": 2458239, "FY2023": 2859860, "FY2024": 3279245, "FY2025": 3494903},

    "common_apic":     {"FY2022": 475256,  "FY2023": 575975,  "FY2024": 638771,  "FY2025": 669949},  # common stock + APIC
    "retained_earn":   {"FY2022": 2926127, "FY2023": 3920362, "FY2024": 4109717, "FY2025": 4522581},
    "aoci":            {"FY2022": -252584, "FY2023": -264256, "FY2024": -424441, "FY2025": -230690},
    "total_equity":    {"FY2022": 3148799, "FY2023": 4232081, "FY2024": 4324047, "FY2025": 4961840},

    "shares_out":      {"FY2022": 122205,  "FY2023": 121106,  "FY2024": 116166,  "FY2025": 111380},
}

# ---------------------------------------------------------------------------
# CASH FLOW (USD thousands) -- reported
# ---------------------------------------------------------------------------
CF = {
    "net_income":  {"FY2022": 854800,  "FY2023": 1550190, "FY2024": 1814616,  "FY2025": 1579183},
    "d_and_a":     {"FY2022": 291791,  "FY2023": 379384,  "FY2024": 446524,   "FY2025": 496228},
    "sbc":         {"FY2022": 78075,   "FY2023": 93560,   "FY2024": 90011,    "FY2025": 62203},
    "cfo":         {"FY2022": 966463,  "FY2023": 2296164, "FY2024": 2272713,  "FY2025": 1602477},
    "capex":       {"FY2022": 638657,  "FY2023": 651865,  "FY2024": 689232,   "FY2025": 680802},
    "cfi":         {"FY2022": -569937, "FY2023": -654132, "FY2024": -798174,  "FY2025": -662118},
    "cff":         {"FY2022": -467487, "FY2023": -548828, "FY2024": -1652508, "FY2025": -1208656},
    "buybacks":    {"FY2022": 444001,  "FY2023": 558652,  "FY2024": 1636879,  "FY2025": 1178349},
}

# ---------------------------------------------------------------------------
# Q2 FY2026 actuals & FY2026 guidance (Sep 3, 2026 earnings release)
# ---------------------------------------------------------------------------
GUIDANCE = {
    "fy2026_rev_low": 10350000,
    "fy2026_rev_high": 10500000,
    "fy2026_eps_low": 9.48,
    "fy2026_eps_high": 9.73,
    "fy2026_tax_rate": 0.30,
    "q3_rev_low": 2290000,
    "q3_rev_high": 2320000,
    "tariff_refund": 134500,   # $134.5M IEEPA refund, Q2 FY2026, reduced COGS
    "tariff_om_bps": 560,      # +560 bps to Q2 operating margin
}

# ---------------------------------------------------------------------------
# Market data (as of most recent close following Q2 FY2026 print, early Sep 2026)
# ---------------------------------------------------------------------------
LEASE_FY25 = BS['op_lease_cur']['FY2025'] + BS['op_lease_noncur']['FY2025']  # ASC 842 debt equiv. ($000)

MKT = {
    "price": 100.00,          # ~ price after -18% post-earnings reaction
    "shares_out": 111380,     # thousands (FY2025 10-K)
    "cash": 1807202,          # thousands
    "debt_funded": 0,         # no term loans / bonds; undrawn revolver
    "lease_cur": BS['op_lease_cur']['FY2025'],
    "lease_noncur": BS['op_lease_noncur']['FY2025'],
    "debt": LEASE_FY25,      # EV bridge: operating leases as debt equivalent ($000)
    "week52_high": 225.98,
    "week52_low": 99.64,
    "beta": 0.86,
}

# ---------------------------------------------------------------------------
# Source links for blue-font (reported) figures — click value or header link
# ---------------------------------------------------------------------------
def _sec_10k(accession, filename):
    acc = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/1397187/{acc}/{filename}"

SOURCES = {
    "edgar_company": "https://www.sec.gov/edgar/browse/?CIK=1397187",
    "edgar_xbrl": "https://www.sec.gov/edgar/browse/?CIK=1397187&owner=exclude",
    # Q2 FY2026 results & FY2026 guidance (Sep 3, 2026 press release)
    "earnings_sep2026": "https://corporate.lululemon.com/newsroom/press-releases/2026/09-03-2026-210528733",
    "earnings_hub": "https://corporate.lululemon.com/newsroom/press-releases",
    "stock_info": "https://corporate.lululemon.com/investors",
    "nasdaq_quote": "https://www.nasdaq.com/market-activity/stocks/lulu",
    "fred_dgs10": "https://fred.stlouisfed.org/series/DGS10",
    "fred_gdpc1": "https://fred.stlouisfed.org/series/GDPC1",
    "damodaran_erp": "https://www.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histimpl.html",
    "damodaran_tax": "https://www.stern.nyu.edu/~adamodar/New_Home_Page/datafile/taxrate.html",
    "damodaran_betas": "https://www.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html",
    "lulu_stats": "https://stockanalysis.com/stocks/lulu/statistics/",
    "lulu_forecast": "https://stockanalysis.com/stocks/lulu/forecast/",
    "alo_reuters_2023": "https://www.reuters.com/markets/deals/alo-yoga-parent-seeks-investment-10-bln-valuation-sources-2023-10-30/",
    "alo_reuters_2026": "https://www.reuters.com/business/finance/alo-yoga-primed-public-offering-or-sale-after-owners-sale-wholesale-t-shirt-2026-06-23/",
    "alo_forbes": "https://www.forbes.com/sites/jemimamcevoy/2025/04/01/the-founders-of-this-lululemon-rival-are-worth-nearly-5-billion-each/",
}

# FY label -> (accession, primary 10-K document)
FILING_10K = {
    "FY2022": ("0001397187-23-000012", "lulu-20230129.htm"),
    "FY2023": ("0001397187-24-000010", "lulu-20240128.htm"),
    "FY2024": ("0001397187-25-000013", "lulu-20250202.htm"),
    "FY2025": ("0001397187-26-000020", "lulu-20260201.htm"),
}

def filing_url(fy):
    """Direct Form 10-K HTML (Ctrl+F works on this page; not the EDGAR viewer TOC)."""
    acc, filename = FILING_10K[fy]
    return f"https://www.sec.gov/Archives/edgar/data/1397187/{acc.replace('-', '')}/{filename}"

# ---------------------------------------------------------------------------
# Red-assumption justifications (~20 words each) for DCF / 3-statement models
# ---------------------------------------------------------------------------
JUST = {
    # WACC tab
    "wacc_rf": "4.8% risk-free = FRED DGS10 on 2026-09-03 (4.77%) rounded; replaces the stale 4.3% input.",
    "wacc_erp": "6.0% ERP is a conservative overlay vs Damodaran Jan-2026 implied 4.23%; used to keep CoE above the risk-free 4.8%.",
    "wacc_beta_obs": "StockAnalysis Beta (5Y) 0.86 is levered equity βL — the starting point for the unlever / relever chain below.",
    "wacc_beta_unlev": "βu = βL ÷ [1 + (1−T) × funded D/E]. No term loans → funded D/E = 0, so βu = 0.86.",
    "wacc_beta_ind": "Damodaran unlevered Retail (Special Lines) βu 0.95 shown as a sector benchmark only — not the WACC input.",
    "wacc_beta": "βL = βu × [1 + (1−T) × total debt equiv. / E]. Relevers company βu for ASC 842 lease liabilities in WACC.",
    "wacc_kd": "5.0% pre-tax lease-equivalent borrowing cost; cheaper than equity. No funded revolver borrowings per FY25 10-K.",
    "wacc_tax": "30% cash tax matches FY2026 guidance (“approximately 30%”); FY25 effective was 29.5%.",
    "wacc_mkt_px": "$100 share price ≈ NASDAQ Last Sale after Q2 FY2026 guide cut; rounded from $100.61 close.",
    "wacc_mkt_eq": "Market equity = share price × shares outstanding. Capital-structure weighting numerator for relevered β and WACC weights.",
    "wacc_mkt_shares": "111,380k class A shares outstanding at FY25 year-end (10-K cover). Diluted WA 119,068 is for EPS, not market cap.",
    "wacc_lease_d": "ASC 842 operating lease liabilities = current + non-current ($1,798,441k FY25). Treated as debt equivalent in WACC and EV bridge.",
    "wacc_fund_d": "Funded term debt $0 — 10-K: no borrowings outstanding on the revolver. Leases are the only debt equivalent here.",
    "wacc_we": "Equity weight = market cap ÷ (market cap + lease debt). ~86% at $100 and FY25 lease balance.",
    "wacc_wd": "Debt weight = lease liabilities ÷ (market cap + lease debt). ~14% — ASC 842 leases, not bank debt.",
    # Scenarios tab
    "sc_g1": "−6.1% FY2026 revenue growth matches company guidance midpoint of −5% to −7% after Q2 FY2026 print.",
    "sc_gterm": "2.3% FY27–30 growth matches StockAnalysis Revenue Growth Forecast (3Y) of 2.26%; next-year consensus is +2.64%.",
    "sc_m1": "13.2% is the real/run-rate OM: Q2 18.8% minus 560bps of tariff refunds. FY26 then adds the $134.5M refund once on top.",
    "sc_tariff": "Add back $134.5M in FY26 only — already recognized (reduced COGS). +130bps on FY26 sales, not +560bps (that was Q2-only).",
    "sc_mterm": "FY30 15.5% is a partial recovery vs the FY25 10-K OM of 19.9% (unique Ctrl+F). Still well below the FY24 23.7% peak.",
    "sc_wacc": "Base WACC from WACC tab (lease-adjusted CAPM). Bear/bull bracket ±100bps around the calculated base.",
    "sc_g": "2.25% terminal g sits next to FRED GDPC1 Q2/Q2 real GDP ≈ 2.1% (24,269.613 / 23,770.976 − 1).",
    "sc_tax": "30% cash tax from FY2026 outlook; FY25 10-K effective is 29.5% (659,784 / 2,238,967).",
    "sc_da_pct": "4.5% D&A / sales = FY25 496,228 / 11,102,600 on the 10-K cash-flow statement.",
    "sc_capex_pct": "FY26 7.0% is $735M on $10.4B sales, not FY25 $11.1B. Since 7.0% steps toward FY25 6.1%, we assume a fade: DCF 5.5% = (7.0+6.0+5.5+5.0+4.0)/5.",
    "sc_capex_sales": "FY26 sales are $10.35–$10.50B (earnings). Use that $10.4B year as the 7% denominator, not FY25 $11.1B.",
    "sc_ar": "AR sits inside NWC: 190,657 / 11,102,600 = 1.7% of sales. Same NWC block as inv, OCA, AP, accrued.",
    "sc_nwc_pct": "One NWC line: 7.5% of Δsales = AR + inv + OCA − AP − accrued. AR is not a separate FCF item.",
    # DCF valuation
    "dcf_exitm": "Selected exit is Gordon TV / FY30 EBITDA. Identity: (UFCF/EBITDA)×(1+g)/(WACC−g). Not Deckers and not a peer average.",
    "dcf_debt_bridge": "Subtract ASC 842 operating lease liabilities as debt equivalent. Funded term debt $0 per FY25 10-K; leases ≈ $1.80B.",
    # Comps — peer multiples
    "comps_nke": "Nike EV/EBITDA is PitchBook daily EV / TTM EBITDA as of 04-Sep-2026. Black formula, not a typed print.",
    "comps_deck": "Deckers EV/EBITDA is PitchBook daily EV / TTM EBITDA as of 04-Sep-2026. Closest premium-footwear peer.",
    "comps_ads": "adidas EV/EBITDA is PitchBook daily EV / TTM EBITDA as of 04-Sep-2026. Global incumbent.",
    "comps_pb": "PitchBook Comps Set 04-Sep-2026. Multiple = daily EV / TTM EBITDA. UAA excluded from the mean (negative EBITDA).",
    "comps_onon": "On Holding EV/EBITDA 14.0x equals StockAnalysis 14.03x, rounded; high-growth athletic peer.",
    "comps_vfc": "VFC EV/EBITDA 10.7x equals StockAnalysis 10.71x; challenged multi-brand apparel peer.",
    "comps_alo_ask": "Reuters: about $10 billion Moelis ask for Color Image. 2026 follow-up: no deal announced. Ask, not a print.",
    "comps_alo_closed": "Reuters June 2026: no deal was announced on the 2023 process. The $10bn is still an ask.",
    "comps_alo_sales": "Forbes estimates Color Image parent sales nearly $2 billion. Mixed Alo + Bella+Canvas. Not Alo-only, not audited.",
    "comps_alo_impl": "5.0x = 10 / 2. Implied EV/Sales on an unclosed ask and parent sales. Not EV/EBITDA. Not the TV.",
    "comps_ff_ev_lo": "5.0x on FY2030E EBITDA ≈ LULU current trough print. Floor of the range, not the selected exit.",
    "comps_ff_ev_hi": "8.0x = Deckers print as the high end of the range only. Selected exit is Gordon implied, not this cap.",
    "comps_ff_pe_lo": "10x P/E low on FY2026E EPS; trough earnings multiple after guidance reset and sentiment de-rating.",
    "comps_ff_pe_hi": "18x P/E high on FY2026E EPS; modest recovery case still below historical premium LULU multiples.",
    # Sensitivity axes (summary)
    "sens_axes": "Red WACC and g grid values bracket base case ±100bps discount rate and ±75bps terminal growth for sensitivity.",
    "sens_wacc": "WACC axis 9.0–11.0% brackets 10.0% base from CAPM (rf + β×ERP) on WACC tab; ±100bps sensitivity band.",
    "sens_g": "Terminal-g axis 1.5–3.0% brackets 2.25% base; bounded by long-run real GDP and inflation benchmarks.",
    # 3-statement Assumptions tab
    "3s_rev_growth": "FY26 −6.1% is guidance midpoint. FY27 +2.6% then +2.3% tracks StockAnalysis next-year +2.64% and 3Y forecast 2.26%.",
    "3s_gm": "Clean GM 56.5%→58.0% (ex-refunds), still below peak ~58–59%. FY26 COGS is then reduced by the $134.5M IEEPA refund.",
    "3s_tariff": "Add back $134.5M in FY26 only — already recognized (reduced COGS). +130bps on FY26 sales, not +560bps (that was Q2-only).",
    "3s_sga_pct": "FY26 SG&A 42.5% matches YTD 42.3% (earnings), vs FY25 36.7% 10-K; then fades to 39.5% as volume stabilizes.",
    "3s_other_opex": "$7M annual amortization run-rate; stable intangible amortization per recent 10-K disclosure levels.",
    "3s_other_inc": "FY26 $45M other income annualizes YTD $22,829; then steps down as cash is deployed (FY25 was only $28,352).",
    "3s_tax_rate": "30% effective tax rate on projections; conservative vs recent ~29% effective, allows for jurisdictional mix.",
    "3s_da_pct": "D&A 4.5–4.6% of revenue; tracks recent depreciation intensity on PPE and lease-related amortization.",
    "3s_capex_pct": "FY26 7.0% is $735M / $10.4B, not / FY25 $11.1B. Since 7.0% steps toward FY25 6.1%, we assume a fade: 7.0 / 6.0 / 5.5 / 5.0 / 5.0.",
    "3s_sbc": "$62M SBC holds the FY25 10-K run-rate (Stock-based compensation expense 62,203).",
    "3s_inv_pct": "Inventory 33–34% of COGS; slight normalization from FY25 build as Americas demand softens and clears.",
    "3s_ap_pct": "AP 6.8% of COGS; holds near FY25 level reflecting vendor terms and production payment cadence.",
    "3s_accr_pct": "Accrued liabilities 5.8% of revenue; stable comp, marketing, and operating accrual ratio.",
    "3s_ar": "AR 1.7% of sales = FY25 Accounts receivable, net 190,657 / 11,102,600. Wholesale and marketplace receivables; hold the FY25 mix.",
    "3s_oca_pct": "OCA 5.1% is current assets minus cash, AR, and inventory (prepaids / tax receivables). AR is its own line, not inside OCA.",
    "3s_rou_pct": "ROU assets 14.7% of revenue; lease-intensive store model, stable vs FY25 operating lease footprint.",
    "3s_onca_pct": "ONCA 3.0% = FY25 total assets − CA − PPE − ROU − GW (338,947); matches earnings 'deferred taxes and other NC assets'.",
    "3s_olc_pct": "Current lease liabilities 2.7% of revenue; short-term portion of operating lease obligations.",
    "3s_olnc_pct": "Non-current lease liabilities 13.5% of revenue; long-term store lease commitments per 10-K.",
    "3s_ocl_pct": "OCL 5.5% = FY25 current liabilities minus AP, accrued, and current leases (residual ~594k); not the 45,954 'other' line alone.",
    "3s_oncl_pct": "Other non-current liabilities 0.5% of revenue; minor long-term accruals and provisions.",
    "3s_buyback": "$500M annual repurchases; continued capital return at moderated pace vs FY24–25 peak buyback levels.",
    "3s_rep_price": "FY26 repurchase price $100 matches the current quote; path $100→$130 is a recovery assumption (FY25 10-K paid $168–$199).",
    # Revenue driver schedule (bottom-up)
    "drv_openings": "Store openings skew to China (+20 FY26) vs mature Americas (+8). Anchored to FY25 +44 net stores and 11% sq-ft growth.",
    "drv_closures": "Closures assume 2–3 per mature region annually as leases roll; immaterial vs openings in expansion markets.",
    "drv_sqft_store": "4,367 avg sq ft/store = implied FY25 total sq ft / 811 stores. Modest expansion to 4,500 by FY30.",
    "drv_spsf": "FY26 SPSF $1,380 reflects Americas traffic/conversion pressure; recovers toward $1,440 by FY30.",
    "drv_comp_americas": "FY25 Americas comp −3% per 10-K (traffic, conversion, AOV down). FY26 −4%: modest step-down before flat FY27 as promo clears.",
    "drv_comp_china": "FY25 China comp +20% (+19% CCY) per 10-K. FY26 +14%: still fastest region but moderating as the store base scales.",
    "drv_comp_row": "FY25 RoW comp +9% (+7% CCY) per 10-K. FY26 +7%: hold above CCY, fade to +4% by FY30 as APAC/Europe base matures.",
    "drv_comp_store_rev": "Existing-store sales only — prior store rev grown by geo comp %. New doors are a separate line, not in here.",
    "drv_new_store_rev": "Net new doors × sq ft × $/sq ft is plain dollars. /1000 → $000 on this tab. ×0.55 'cause year-one stores don't run full.",
    "drv_f_end_stores": "Simple roll-forward: where you started, plus openings, minus closures. That's your ending store count for the year.",
    "drv_f_beg_stores": "Beginning stores = last year's ending count. Same number, just carried forward into the new fiscal year.",
    "drv_f_total_stores": "Add up Americas + China + RoW ending stores. Should tie to the 10-K total company-operated count.",
    "drv_f_total_sqft": "Ending stores times avg sq ft per box. Gives you total fleet square footage for the productivity math.",
    "drv_f_store_rev_base": "Last year's total store-channel revenue — the base you're growing comps off of, not the new boxes.",
    "drv_f_store_rev": "Brick-and-mortar channel = comp store sales plus whatever the net new stores contributed. That's the whole store line.",
    "drv_f_ecomm_rev": "Traffic × conversion × AOV is dollars. ×1M 'cause sessions are in millions; /1000 puts it in $000 like everything else.",
    "drv_f_other_rev_growth": "Take last year's other-channel revenue and grow it by your assumed %. Wholesale/license/outlets — not stores or dot-com.",
    "drv_f_geo_scale": "FY26 geo rev = FY25 geo rev × (FY26 total rev ÷ FY25 total rev). Total = store + e-comm + other.",
    "drv_f_total_rev": "Bottom-up total = store channel + e-comm + other. That's your driver-built revenue before you check Scenarios.",
    "drv_f_scen_rev": "Pulls the Scenarios base-case revenue path — that's still the number the DCF actually uses, not the driver total.",
    "drv_f_variance": "Driver-built revenue minus Scenarios revenue. Shows you how far off the bottom-up path is from what you're valuing on.",
    "drv_f_var_pct": "Variance as a percent of Scenarios revenue. Easy read on whether the driver schedule is close or way out of line.",
    "drv_ecomm_sessions": "485M FY25 sessions implied from reported e-comm revenue / conv / AOV. FY26 −3% on Americas softness.",
    "drv_ecomm_conv": "3.3% FY26 digital conversion vs 3.5% pressure cited in 10-K Americas; gradual recovery to 3.6%.",
    "drv_ecomm_aov": "AOV $305 FY26 on promo intensity; steps to $315 by FY30 as mix normalizes.",
    "drv_other_rev": "Other channels (wholesale/license/outlets) held at FY25 $1.13B base; low-single-digit growth thereafter.",
    "drv_geo_americas": "Geographic split scales with consolidated growth; FY25 Americas 70.7% of revenue per 10-K segment table.",
    "drv_geo_china": "China 15.8% of FY25 revenue; fastest comp and store growth — expansion market in the driver schedule.",
    "drv_geo_row": "Rest of World 13.5% of FY25 revenue; mid-single-digit comps and steady openings.",
    "drv_mix_women": "Women's mix fades 62.5%→60% as men's penetrates; FY25 reported 63% women's per 10-K category disclosure.",
    "drv_mix_men": "Men's mix rises 24.5%→27%; FY25 men's grew 4% vs women's 5% per 10-K category commentary.",
    "drv_mix_accessories": "Accessories steady at 13% of revenue; FY25 accessories +8% per 10-K category disclosure.",
    "drv_store_traffic": "Fleet traffic memo line — not a separate FCF item. Calibrated to FY25 store productivity commentary.",
    "drv_store_conv": "In-store conversion memo — 10-K cites lower conversion in Americas; FY26 27% improving to 29%.",
    "drv_store_aov": "In-store average transaction $115–$122; supports SPSF and comp-sales bridge on the driver tab.",
}

# Clickable source links for red assumptions (label, URL)
ASSUMPTION_SRC = {
    # WACC
    "wacc_rf": ("FRED: 10Y Treasury (DGS10)", SOURCES["fred_dgs10"]),
    "wacc_erp": ("Damodaran: Historical Implied ERP", SOURCES["damodaran_erp"]),
    "wacc_beta_obs": ("StockAnalysis: LULU Beta (5Y)", SOURCES["lulu_stats"]),
    "wacc_beta_unlev": ("WACC tab: Hamada unlever (funded D/E)", None),
    "wacc_beta_ind": ("Damodaran: US betas by sector (Jan 2026)", SOURCES["damodaran_betas"]),
    "wacc_beta": ("WACC tab: Hamada relever (total D/E incl. leases)", None),
    "wacc_kd": ("LULU FY2025 10-K: lease liabilities & no funded debt", filing_url("FY2025")),
    "wacc_tax": ("Q2 FY2026 outlook: tax rate ≈ 30%", SOURCES["earnings_sep2026"]),
    "wacc_mkt_px": ("NASDAQ LULU last sale (current price)", SOURCES["nasdaq_quote"]),
    "wacc_mkt_shares": ("LULU FY2025 10-K: shares outstanding", filing_url("FY2025")),
    "wacc_mkt_eq": ("WACC tab: price × shares", None),
    "wacc_lease_d": ("LULU FY2025 10-K: operating lease liabilities", filing_url("FY2025")),
    "wacc_fund_d": ("LULU FY2025 10-K (no funded debt)", filing_url("FY2025")),
    "wacc_we": ("WACC tab: capital structure", None),
    "wacc_wd": ("LULU FY2025 10-K: operating lease liabilities", filing_url("FY2025")),
    # Scenarios
    "sc_g1": ("LULU Q2 FY2026 earnings release", SOURCES["earnings_sep2026"]),
    "sc_gterm": ("StockAnalysis: LULU 3Y revenue forecast", SOURCES["lulu_stats"]),
    "sc_m1": ("Q2 FY2026 release: 18.8% OM minus 560bps tariffs", SOURCES["earnings_sep2026"]),
    "sc_tariff": ("Q2 FY2026 release: $134.5M IEEPA tariff refunds", SOURCES["earnings_sep2026"]),
    "sc_mterm": ("FY2025 10-K: Operating margin 19.9% (one hit)", filing_url("FY2025")),
    "sc_wacc": ("WACC tab: CAPM build", None),  # internal link set in build_dcf.py
    "sc_g": ("FRED: Real GDP (GDPC1)", SOURCES["fred_gdpc1"]),
    "sc_tax": ("Q2 FY2026 outlook: tax rate ≈ 30%", SOURCES["earnings_sep2026"]),
    "sc_da_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "sc_capex_pct": ("LULU 10-K: 2026 capex guide $725–745M", filing_url("FY2025")),
    "sc_capex_sales": ("Q2 FY2026 outlook: $10.350B–$10.500B sales", SOURCES["earnings_sep2026"]),
    "sc_ar": ("LULU FY2025 10-K: Accounts receivable, net", filing_url("FY2025")),
    "sc_nwc_pct": ("LULU BS/IS historical (10-K)", filing_url("FY2025")),
    # DCF / comps
    "dcf_exitm": ("DCF: Gordon implied exit (TV / FY30 EBITDA)", None),
    "dcf_debt_bridge": ("LULU FY2025 10-K: operating lease liabilities", filing_url("FY2025")),
    "comps_nke": ("PitchBook Comps Set 04-Sep-2026", None),
    "comps_deck": ("PitchBook Comps Set 04-Sep-2026", None),
    "comps_ads": ("PitchBook Comps Set 04-Sep-2026", None),
    "comps_pb": ("PitchBook Comps Set 04-Sep-2026", None),
    "comps_onon": ("StockAnalysis: ONON EV/EBITDA", "https://stockanalysis.com/stocks/onon/statistics/"),
    "comps_vfc": ("StockAnalysis: VFC EV/EBITDA", "https://stockanalysis.com/stocks/vfc/statistics/"),
    "comps_alo_ask": ("Reuters: Alo parent $10bn Moelis ask", SOURCES["alo_reuters_2023"]),
    "comps_alo_closed": ("Reuters: Alo 2023 process — no deal", SOURCES["alo_reuters_2026"]),
    "comps_alo_sales": ("Forbes: Color Image nearly $2bn sales", SOURCES["alo_forbes"]),
    "comps_ff_ev_lo": ("StockAnalysis: LULU EV/EBITDA", SOURCES["lulu_stats"]),
    "comps_ff_ev_hi": ("StockAnalysis: DECK EV/EBITDA", "https://stockanalysis.com/stocks/deck/statistics/"),
    "comps_ff_pe_lo": ("StockAnalysis: LULU Forward P/E", SOURCES["lulu_stats"]),
    "comps_ff_pe_hi": ("StockAnalysis: NKE Forward P/E", "https://stockanalysis.com/stocks/nke/statistics/"),
    "sens_axes": ("Scenarios: WACC & terminal g", None),  # internal links set in build_dcf.py
    "sens_wacc": ("Damodaran: Historical Implied ERP", SOURCES["damodaran_erp"]),
    "sens_g": ("FRED: Real GDP (GDPC1)", SOURCES["fred_gdpc1"]),
    # 3-statement
    "3s_rev_growth": ("StockAnalysis: LULU revenue forecast (FY27–30)", SOURCES["lulu_stats"]),
    "3s_gm": ("LULU historical gross margin (10-K)", filing_url("FY2025")),
    "3s_tariff": ("Q2 FY2026 release: $134.5M IEEPA tariff refunds", SOURCES["earnings_sep2026"]),
    "3s_sga_pct": ("LULU Q2 FY2026 earnings release (YTD SG&A %)", SOURCES["earnings_sep2026"]),
    "3s_other_opex": ("LULU FY2025 10-K", filing_url("FY2025")),
    "3s_other_inc": ("Q2 FY2026 YTD other income (earnings)", SOURCES["earnings_sep2026"]),
    "3s_tax_rate": ("LULU Q2 FY2026 outlook (≈30% tax)", SOURCES["earnings_sep2026"]),
    "3s_da_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "3s_capex_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "3s_sbc": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "3s_inv_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_ap_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_accr_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_ar": ("LULU FY2025 10-K: Accounts receivable, net", filing_url("FY2025")),
    "3s_oca_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_rou_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_onca_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_olc_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_olnc_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_ocl_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_oncl_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_buyback": ("LULU CF: share repurchases (10-K)", filing_url("FY2025")),
    "3s_rep_price": ("NASDAQ LULU last sale (current price)", SOURCES["nasdaq_quote"]),
    # Revenue drivers
    "drv_openings": ("FY2025 10-K: store count & expansion", filing_url("FY2025")),
    "drv_closures": ("FY2025 10-K: lease renewal / fleet", filing_url("FY2025")),
    "drv_sqft_store": ("FY2025 10-K: sales per square foot", filing_url("FY2025")),
    "drv_spsf": ("FY2025 10-K: comparable sales & SPSF", filing_url("FY2025")),
    "drv_comp_americas": ("FY2025 10-K: Americas comparable sales", filing_url("FY2025")),
    "drv_comp_china": ("FY2025 10-K: China Mainland comparable sales", filing_url("FY2025")),
    "drv_comp_row": ("FY2025 10-K: Rest of World comparable sales", filing_url("FY2025")),
    "drv_comp_store_rev": ("Revenue Drivers: comp-store formula", None),
    "drv_new_store_rev": ("Revenue Drivers: net-new-store formula", None),
    "drv_f_end_stores": ("Revenue Drivers: ending-stores formula", None),
    "drv_f_beg_stores": ("Revenue Drivers: beginning-stores formula", None),
    "drv_f_total_stores": ("Revenue Drivers: total-stores formula", None),
    "drv_f_total_sqft": ("Revenue Drivers: total-sqft formula", None),
    "drv_f_store_rev_base": ("Revenue Drivers: prior store-rev formula", None),
    "drv_f_store_rev": ("Revenue Drivers: store-channel formula", None),
    "drv_f_ecomm_rev": ("Revenue Drivers: e-comm formula", None),
    "drv_f_other_rev_growth": ("Revenue Drivers: other-rev formula", None),
    "drv_f_geo_scale": ("Revenue Drivers: geo-scale formula", None),
    "drv_f_total_rev": ("Revenue Drivers: total-rev formula", None),
    "drv_f_scen_rev": ("Scenarios tab: base revenue", None),
    "drv_f_variance": ("Revenue Drivers: variance formula", None),
    "drv_f_var_pct": ("Revenue Drivers: variance % formula", None),
    "drv_ecomm_sessions": ("FY2025 10-K: e-commerce revenue", filing_url("FY2025")),
    "drv_ecomm_conv": ("FY2025 10-K: digital / traffic commentary", filing_url("FY2025")),
    "drv_ecomm_aov": ("FY2025 10-K: e-commerce revenue", filing_url("FY2025")),
    "drv_other_rev": ("FY2025 10-K: channel revenue table", filing_url("FY2025")),
    "drv_geo_americas": ("FY2025 10-K: segmented net revenue", filing_url("FY2025")),
    "drv_geo_china": ("FY2025 10-K: segmented net revenue", filing_url("FY2025")),
    "drv_geo_row": ("FY2025 10-K: segmented net revenue", filing_url("FY2025")),
    "drv_mix_women": ("FY2025 10-K: product category mix", filing_url("FY2025")),
    "drv_mix_men": ("FY2025 10-K: product category mix", filing_url("FY2025")),
    "drv_mix_accessories": ("FY2025 10-K: product category mix", filing_url("FY2025")),
    "drv_store_traffic": ("FY2025 10-K: store traffic commentary", filing_url("FY2025")),
    "drv_store_conv": ("FY2025 10-K: conversion rate commentary", filing_url("FY2025")),
    "drv_store_aov": ("FY2025 10-K: average order value commentary", filing_url("FY2025")),
}

# Exact Ctrl+F strings — every quoted phrase appears verbatim in the linked source.
# Format: Ctrl+F "phrase" → what to read; model value noted where it is an assumption.
SOURCE_HINT = {
    # WACC
    "wacc_rf": 'Ctrl+F "2026-09-03" → observation 4.77. Model uses 4.8%. Also Ctrl+F "DGS10" for the series title.',
    "wacc_erp": 'Ctrl+F "4.23%" on the 2025 row (last data row). Header is "Implied ERP (FCFE)". Model uses 6.0%.',
    "wacc_beta_obs": 'Ctrl+F "Beta (5Y)" → 0.86. Levered equity βL (StockAnalysis).',
    "wacc_beta_unlev": 'βu = E[βL] ÷ (1 + (1−T) × E[funded debt] ÷ E[market equity]). Funded debt = $0 → βu = 0.86.',
    "wacc_beta_ind": 'Ctrl+F "Retail (Special Lines)" → Unlevered beta 0.95 (βu). Sector benchmark only — not used in WACC.',
    "wacc_beta": 'βL = E[βu] × (1 + (1−T) × E[total debt equiv.] ÷ E[market equity]). Total debt = leases + funded.',
    "wacc_kd": 'Ctrl+F "Current lease liabilities" + "Non-current lease liabilities" → debt equiv. 5.0% illustrative lease borrowing cost.',
    "wacc_tax": 'Ctrl+F "a tax rate of approximately 30%" on the FY2026 outlook paragraph.',
    "wacc_mkt_px": 'Ctrl+F "closed at $100.61" → Last Sale $100.61 (NASDAQ). Model uses $100 rounded.',
    "wacc_mkt_shares": 'Ctrl+F "111,380 and 116,166 issued and outstanding" → 111,380k class A shares (FY25 10-K cover).',
    "wacc_mkt_eq": 'Share price row × shares outstanding row on this tab.',
    "wacc_lease_d": 'Ctrl+F "Current lease liabilities" → 298,724 | "Non-current lease liabilities" → 1,499,717 | sum = 1,798,441 ($000)',
    "wacc_fund_d": 'Ctrl+F "no borrowings were outstanding under this facility" → funded debt $0',
    "wacc_we": 'Equity weight = market cap ÷ (market cap + lease debt). Live formula on this tab.',
    "wacc_wd": 'Debt weight = lease liabilities ÷ (market cap + lease debt). ~14% at FY25 balances.',
    # Scenarios
    "sc_g1": 'Ctrl+F "decline of 5% to 7%" → FY2026 revenue guide; model −6.1% midpoint',
    "sc_gterm": 'Ctrl+F "Revenue Growth Forecast (3Y)" → 2.26%. That is the unique line (do not search "Net revenue"). Model FY27–30 uses 2.3%.',
    "sc_m1": 'Ctrl+F "18.8%" (Q2 OM) and "560 basis points" (tariff boost). Real run-rate = 18.8% − 5.6% = 13.2%. Do not use "decreased 13%".',
    "sc_tariff": 'Ctrl+F "134.5 million" → IEEPA tariff refunds reduced COGS. Add this dollar amount to FY26 EBIT only. Full-year boost = 134.5 / FY26 sales ≈ 1.3ppt, not 5.6ppt.',
    "sc_mterm": 'Ctrl+F "19.9%" (one hit) → Operating margin decreased 380 basis points to 19.9%. That is FY25 OM. Do not search the words Income from operations (17 hits). Model FY30 15.5% is a partial-recovery assumption.',
    "sc_wacc": 'WACC tab → green cell = rf 4.8% + relevered β × ERP 6.0%',
    "sc_g": 'Ctrl+F "Q2 2026" → 24,269.613 and "Q2 2025" → 23,770.976. YoY = 2.1%. Model terminal g 2.25%.',
    "sc_tax": 'Ctrl+F "a tax rate of approximately 30%" on the earnings outlook. FY25 10-K: "Income tax expense" 659,784 ÷ "Income before income tax expense" 2,238,967 = 29.5%.',
    "sc_da_pct": 'Ctrl+F "Depreciation and amortization" → 496,228 on this 10-K HTML (not the SEC viewer). ÷ "Net revenue" 11,102,600 = 4.5%.',
    "sc_capex_pct": 'Ctrl+F "680,802" → FY25 capex $680.8M. Ctrl+F "11,102,600" → FY25 sales $11.1B. Ctrl+F "$725.0 million and $745.0 million" → FY26 capex mid $735M.',
    "sc_capex_sales": 'Ctrl+F "$10.350 billion to $10.500 billion" → FY26 net revenue guide (one hit). Mid $10.425B.',
    "sc_ar": 'Ctrl+F "Accounts receivable, net" (one hit on the BS) → 190,657. 190,657 / 11,102,600 = 1.7% of sales. Lives inside the NWC line below.',
    "sc_nwc_pct": 'One NWC. Ctrl+F "Accounts receivable, net" → 190,657 | "Inventories" → 1,700,753 | "Accounts payable" → 331,421 | "Accrued liabilities and other" → 662,982.',
    # DCF / comps
    "dcf_exitm": "No peer Ctrl+F. This cell = Gordon TV / FY30 EBITDA = (UFCF/EBITDA)×(1+g)/(WACC−g). WACC and g are sourced on those rows.",
    "comps_nke": "No public HTML. PitchBook Comps Set 04-Sep-2026. Prove daily EV and TTM EBITDA on that screen; E = I/J.",
    "comps_deck": "No public HTML. PitchBook Comps Set 04-Sep-2026. Prove daily EV and TTM EBITDA on that screen; E = I/J.",
    "comps_ads": "No public HTML. PitchBook Comps Set 04-Sep-2026. Prove daily EV and TTM EBITDA on that screen; E = I/J.",
    "comps_pb": "No public HTML. PitchBook Comps Set 04-Sep-2026. Prove daily EV and TTM EBITDA on that screen; E = I/J.",
    "comps_onon": 'Ctrl+F "EV / EBITDA" → 14.03. Model uses 14.0x.',
    "comps_vfc": 'Ctrl+F "EV / EBITDA" → 10.71. Model uses 10.7x.',
    "comps_alo_ask": 'Ctrl+F "at about $10 billion" and "hired investment bank Moelis". That is the 2023 ask, not a closed EV.',
    "comps_alo_closed": 'Ctrl+F "roughly $10 billion" and "No deal was announced". Still an ask.',
    "comps_alo_sales": 'Ctrl+F "nearly $2 billion". Forbes estimate for Color Image parent (Alo + Bella+Canvas), not Alo-only EBITDA.',
    "comps_ff_ev_lo": 'Ctrl+F "EV / EBITDA" → LULU trough ~5x; football-field floor 5.0x',
    "comps_ff_ev_hi": 'Ctrl+F "EV / EBITDA" → 7.95. Football-field cap 8.0x (Deckers). Not the selected exit.',
    "comps_ff_pe_lo": 'Ctrl+F "Forward PE" → LULU ~12.3x; low-case multiple assumption 10.0x',
    "comps_ff_pe_hi": 'Ctrl+F "Forward PE" → NKE peer benchmark; high-case assumption 18.0x',
    "sens_axes": 'Scenarios tab → Ctrl+F "WACC" and "Terminal growth" rows (base-case inputs)',
    "sens_wacc": 'Ctrl+F "4.23%" (2025 implied ERP). Sensitivity WACC axis brackets the 10.5% base.',
    "sens_g": 'Ctrl+F "GDPC1" → bounds terminal-g sensitivity grid 1.5%–3.0%',
    # 3-statement
    "3s_rev_growth": 'FY26: earnings release Ctrl+F "decline of 5% to 7%". FY27–30: this page Ctrl+F "Revenue Growth Forecast (3Y)" → 2.26%. Model 2.3%.',
    "3s_gm": 'Ctrl+F "Gross profit" → 6,284,132 ÷ "Net revenue" 11,102,600 = 56.6% FY25 anchor',
    "3s_tariff": 'Ctrl+F "134.5 million" → IEEPA tariff refunds reduced COGS. Add this dollar amount to FY26 EBIT only. Full-year boost = 134.5 / FY26 sales ≈ 1.3ppt, not 5.6ppt.',
    "3s_sga_pct": 'Ctrl+F "41.7%" (Q2 SG&A % of net revenue) and "42.3%" (first two quarters). FY25 10-K is 36.7%. Model FY26 42.5%',
    "3s_other_opex": 'Ctrl+F "Amortization of intangible assets" → 6,961 ($000); model $7,000/yr',
    "3s_other_inc": 'Ctrl+F "22,829" → YTD other income ($000) on the earnings P&L. FY26 model $45,000 annualizes that run-rate.',
    "3s_tax_rate": 'Ctrl+F "a tax rate of approximately 30%" on the earnings outlook; FY25 10-K 29.5%. 3-statement model 30%',
    "3s_da_pct": 'Ctrl+F "Depreciation and amortization" → 496,228 ÷ "Net revenue" 11,102,600 = 4.5%',
    "3s_capex_pct": 'Ctrl+F "680,802" → FY25 capex $680.8M. Ctrl+F "11,102,600" → FY25 sales $11.1B. Ctrl+F "$725.0 million and $745.0 million" → FY26 capex mid $735M.',
    "3s_sbc": 'Ctrl+F "Stock-based compensation expense" → 62,203 on this 10-K HTML. Model $62,000.',
    "3s_inv_pct": 'Ctrl+F "Inventories" → 1,700,753 ÷ "Cost of goods sold" 4,818,468 = 35.3%',
    "3s_ap_pct": 'Ctrl+F "Accounts payable" → 331,421 ÷ "Cost of goods sold" 4,818,468 = 6.9%',
    "3s_accr_pct": 'Ctrl+F "Accrued liabilities and other" → 662,982 ÷ "Net revenue" 11,102,600 = 6.0%',
    "3s_ar": 'Ctrl+F "Accounts receivable, net" (one hit on the BS) → 190,657 ($000). 190,657 / 11,102,600 = 1.7% of sales. Model holds 1.7%.',
    "3s_oca_pct": 'OCA = "Total current assets" 4,262,701 − cash 1,807,202 − "Accounts receivable, net" 190,657 − "Inventories" 1,700,753 = 564,089 (5.1% of sales). AR is not inside this residual.',
    "3s_rou_pct": 'Ctrl+F "Right-of-use lease assets" → 1,630,181 ÷ "Net revenue" 11,102,600 = 14.7%',
    "3s_onca_pct": 'Ctrl+F "Total assets" 8,456,743 − current assets − "Property and equipment, net" − "Right-of-use lease assets" − goodwill/intangibles = 338,947 (3.0%)',
    "3s_olc_pct": 'Ctrl+F "Current lease liabilities" → 298,724 ÷ "Net revenue" 11,102,600 = 2.7%',
    "3s_olnc_pct": 'Ctrl+F "Non-current lease liabilities" → 1,499,717 ÷ "Net revenue" 11,102,600 = 13.5%',
    "3s_ocl_pct": 'Ctrl+F "Total current liabilities" 1,887,548 − "Accounts payable" 331,421 − "Accrued liabilities and other" 662,982 − "Current lease liabilities" 298,724 ≈ 594,421 (5.4%)',
    "3s_oncl_pct": 'Ctrl+F "Other non-current liabilities" → 55,360 ÷ "Net revenue" 11,102,600 = 0.5%',
    "3s_buyback": 'Ctrl+F "Repurchase of common stock" → ( 1,178,349 ) ($000) FY25; model (500,000)/yr',
    "3s_rep_price": 'Ctrl+F "closed at $100.61" → Last Sale (NASDAQ). Model repurchase price starts at $100.',
    # Revenue drivers
    "drv_openings": 'Ctrl+F "Total company-operated stores" → 811 (FY25). FY25 added 44 net stores; model skews openings to China.',
    "drv_closures": 'Ctrl+F "lease" / store fleet — immaterial closures vs expansion; 2–3 per mature region.',
    "drv_sqft_store": 'Ctrl+F "sales per square foot were $1,426" → implied 3.54M sq ft / 811 stores ≈ 4,367 sq ft.',
    "drv_spsf": 'Ctrl+F "sales per square foot were $1,426" → FY25 reported $1,426; FY26 model $1,380 on traffic pressure.',
    "drv_comp_americas": 'Ctrl+F "Americas comparable sales decreased 3%" → FY25 anchor. FY26 model −4% before stabilization.',
    "drv_comp_china": 'Ctrl+F "China Mainland comparable sales increased 20%" → FY25 anchor. FY26 model +14% on a larger base.',
    "drv_comp_row": 'Ctrl+F "Rest of World comparable sales increased 9%" → FY25 +9% (+7% CCY). FY26 model +7%, fading to +4%.',
    "drv_comp_store_rev": "=prior store rev×(71%×(1+Am comp)+16%×(1+China)+13%×(1+RoW)). Existing base only — new stores are on the next row.",
    "drv_new_store_rev": (
        "=(open−close each geo)×sq ft×$/sq ft → raw annual $. /1000 flips that to $000 (this whole tab is thousands). "
        "×0.55 = year-one haircut — new boxes don't run at full SPSF yet. FY27 ex: (D7−D8+D11−D12+D15−D16)×D19×D24/1000×0.55."
    ),
    "drv_f_end_stores": "=beginning stores + openings − closures. Roll-forward — nothing fancy, just how many doors you end with.",
    "drv_f_beg_stores": "=prior-year ending stores. Same count, new column — that's where the year starts.",
    "drv_f_total_stores": "=Americas ending + China ending + RoW ending. Should foot to total company-operated stores.",
    "drv_f_total_sqft": "=total ending stores × avg sq ft per store. Fleet size in square feet for the productivity lines.",
    "drv_f_store_rev_base": "=prior column's total store-channel revenue. Base you're applying comps to — not new-store contribution.",
    "drv_f_store_rev": "=comparable store revenue + net new store revenue. Full brick-and-mortar channel for the year.",
    "drv_f_ecomm_rev": "=sessions×1,000,000×conversion×AOV/1000 → e-comm $000. ×1M 'cause sessions are millions; /1000 matches tab units.",
    "drv_f_other_rev_growth": "=prior other-channel rev × (1 + growth %). Wholesale/license/outlets — grown off last year.",
    "drv_f_geo_scale": "FY26 geo rev = FY25 geo rev × (FY26 total rev ÷ FY25 total rev). Total rev = store + e-comm + other.",
    "drv_f_total_rev": "=store channel + e-commerce + other. Bottom-up top line before you reconcile to Scenarios.",
    "drv_f_scen_rev": "=Scenarios! base-case revenue. DCF still runs off this path — not the driver total.",
    "drv_f_variance": "=bottom-up total − Scenarios revenue. Positive = drivers above Scenarios; negative = below.",
    "drv_f_var_pct": "=variance ÷ Scenarios revenue. Quick % read on how far the driver build is from the valuation path.",
    "drv_ecomm_sessions": 'Ctrl+F "E-commerce" → 4,918,697 ($000). Sessions implied from revenue / conv / AOV.',
    "drv_ecomm_conv": 'Ctrl+F "lower conversion rates" (Americas MD&A). FY26 model 3.3% digital conversion.',
    "drv_ecomm_aov": 'Ctrl+F "average order value" (Americas MD&A). FY26 model $305 AOV.',
    "drv_other_rev": 'Ctrl+F "Company-operated stores" and "E-commerce" → residual other channels ≈ $1.13B.',
    "drv_geo_americas": 'Ctrl+F "Americas" segment table → $7,847,044 net revenue (70.7%).',
    "drv_geo_china": 'Ctrl+F "China Mainland" segment → $1,754,799 net revenue (15.8%).',
    "drv_geo_row": 'Ctrl+F "Rest of World" segment → $1,500,757 net revenue (13.5%).',
    "drv_mix_women": 'Ctrl+F "women\'s, men\'s, and accessories" → 63% / 24% / 13% of net revenue.',
    "drv_mix_men": 'Ctrl+F "men\'s" category growth 4% in MD&A; mix rises to 27% by FY30.',
    "drv_mix_accessories": 'Ctrl+F "accessories" +8% category growth; hold 13% mix.',
    "drv_store_traffic": 'Ctrl+F "store traffic" → lower traffic in Americas; memo line only.',
    "drv_store_conv": 'Ctrl+F "conversion rates" → lower in Americas; in-store memo 27%→29%.',
    "drv_store_aov": 'Ctrl+F "average order value" → lower AOV in Americas; memo $115–$122.',
}

# Organized Ctrl+F block for the capex assumption cell (column D)
BETA_CTRL_F = (
    "Company β: unlever observed βL, then relever for ASC 842 lease debt.\n"
    "\n"
    "Step 1 — Observed levered βL (StockAnalysis)\n"
    '  Ctrl+F "Beta (5Y)"  \u2192  0.86\n'
    "\n"
    "Step 2 — Unlever (Hamada; funded debt only)\n"
    "  \u03b2u = \u03b2L \u00f7 [1 + (1\u2212T) \u00d7 D_funded/E]\n"
    "  Funded debt $0 (10-K) \u2192 \u03b2u = 0.86\n"
    "\n"
    "Step 3 — Relever (total debt equiv. incl. leases)\n"
    "  \u03b2L = \u03b2u \u00d7 [1 + (1\u2212T) \u00d7 D_total/E]\n"
    '  Ctrl+F "Current lease liabilities"  \u2192  298,724\n'
    '  Ctrl+F "Non-current lease liabilities"  \u2192  1,499,717\n'
    "  D_total = 1,798,441 ($000) + funded debt $0\n"
    "\n"
    "Sector benchmark (Damodaran) \u2014 not used\n"
    '  Ctrl+F "Retail (Special Lines)"  \u2192  Unlevered beta 0.95'
)

CAPEX_CTRL_F = (
    "FY25 actual (10-K)\n"
    '  Ctrl+F "680,802"  →  capex $680.8M\n'
    '  Ctrl+F "11,102,600"  →  sales $11.1B\n'
    "  680.8 / 11,102.6 = 6.1%\n"
    "\n"
    "FY26 capex (10-K)\n"
    '  Ctrl+F "$725.0 million and $745.0 million"  →  mid $735M\n'
    "\n"
    "FY26 sales (earnings)\n"
    '  Ctrl+F "$10.350 billion to $10.500 billion"  →  $10.35–$10.50B (mid $10.425B)\n'
    "  7.0% = $735M / $10.425B    ($10B year, not FY25 $11.1B)\n"
    "\n"
    "DCF blend\n"
    "  Fade  7.0 / 6.0 / 5.5 / 5.0 / 4.0\n"
    "  5.5% = (7.0+6.0+5.5+5.0+4.0)/5"
)


COVER_HINTS = {
    "edgar_xbrl": 'Ctrl+F "10-K" → FY2025 accession 0001397187-26-000020',
    "filing_fy2025": 'Ctrl+F "Net revenue" → 11,102,600 or "Depreciation and amortization" → 496,228 on this 10-K HTML file',
    "earnings_sep2026": 'Ctrl+F "decline of 5% to 7%" | "$10.350 billion to $10.500 billion" | "$9.48 to $9.73"',
    "nasdaq_quote": 'Ctrl+F "closed at $100.61" → Last Sale on NASDAQ quote page',
}

REPORTED_HINTS = {
    "10k": 'Ctrl+F "CONSOLIDATED STATEMENTS OF OPERATIONS" → locate line item in FY2025 column',
    "10k_is": 'Ctrl+F "Net revenue" → 11,102,600 ($000) in FY2025 column',
    "10k_ebit": 'Ctrl+F "19.9%" (one hit) → FY25 operating margin. Do not search the words Income from operations (17 hits). Dollar EBIT is 2,210,615 ($000).',
    "10k_bs": 'Ctrl+F "Cash and cash equivalents" → 1,807,202 ($000) FY2025',
    "10k_debt": 'Ctrl+F "Current lease liabilities" → 298,724 | "Non-current lease liabilities" → 1,499,717 | sum 1,798,441 ($000). Funded debt: "no borrowings were outstanding"',
    "10k_cf": 'Ctrl+F "Depreciation and amortization" → 496,228 | "680,802" capex ($000)',
    "10k_shares": 'Ctrl+F "Diluted weighted-average number of shares outstanding" → 119,068 (000) for EPS',
    "10k_shares_out": 'Ctrl+F "111,380 and 116,166 issued and outstanding" → 111,380k class A (FY25 10-K cover)',
    "10k_ebitda": 'Ctrl+F "19.9%" → FY25 OM (EBIT $2,210,615). Ctrl+F "496,228" → Depreciation and amortization. Do not search the words Income from operations (17 hits).',
    "earnings_rev": 'Ctrl+F "$10.350 billion to $10.500 billion" → FY2026 net revenue guidance',
    "earnings_eps": 'Ctrl+F "$9.48 to $9.73" → FY2026 diluted EPS guidance; model midpoint $9.61',
    "nasdaq": 'Ctrl+F "closed at $100.61" → Last Sale on NASDAQ quote page',
}
