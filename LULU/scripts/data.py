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
MKT = {
    "price": 100.00,          # ~ price after -18% post-earnings reaction
    "shares_out": 111380,     # thousands (FY2025 10-K)
    "cash": 1807202,          # thousands
    "debt": 0,                # no funded debt; undrawn revolver
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
    "lulu_stats": "https://stockanalysis.com/stocks/lulu/statistics/",
    "lulu_forecast": "https://stockanalysis.com/stocks/lulu/forecast/",
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
    "wacc_beta": "0.95 levered beta vs StockAnalysis Beta (5Y) 0.86; modest uplift for post-guidance volatility.",
    "wacc_kd": "5.0% illustrative pre-tax debt cost; 10-K: no borrowings outstanding on the revolver.",
    "wacc_tax": "30% cash tax matches FY2026 guidance (“approximately 30%”); FY25 effective was 29.5%.",
    "wacc_we": "100% equity weight; net-cash balance sheet with no material funded debt at market values per FY2025 10-K.",
    "wacc_wd": "0% debt weight; no outstanding term loans or bonds, so WACC effectively equals levered cost of equity here.",
    # Scenarios tab
    "sc_g1": "−6.1% FY2026 revenue growth matches company guidance midpoint of −5% to −7% after Q2 FY2026 print.",
    "sc_gterm": "2.3% FY27–30 growth matches StockAnalysis Revenue Growth Forecast (3Y) of 2.26%; next-year consensus is +2.64%.",
    "sc_m1": "13.2% is the real/run-rate OM: Q2 18.8% minus 560bps of tariff refunds. FY26 then adds the $134.5M refund once on top.",
    "sc_tariff": "Add back $134.5M in FY26 only — already recognized (reduced COGS). +130bps on FY26 sales, not +560bps (that was Q2-only).",
    "sc_mterm": "FY30 15.5% is a recovery assumption vs FY25 10-K OM 19.9% (2,210,615 / 11,102,600); still well below FY24 peak ~23.7%.",
    "sc_wacc": "10.5% base WACC = rf 4.8% + 0.95×6.0% ERP on the WACC tab (FRED 4.77% rounded).",
    "sc_g": "2.25% terminal g sits next to FRED GDPC1 Q2/Q2 real GDP ≈ 2.1% (24,269.613 / 23,770.976 − 1).",
    "sc_tax": "30% cash tax from FY2026 outlook; FY25 10-K effective is 29.5% (659,784 / 2,238,967).",
    "sc_da_pct": "4.5% D&A / sales = FY25 496,228 / 11,102,600 on the 10-K cash-flow statement.",
    "sc_capex_pct": "5.5% is a 5-year blend: 2026 10-K guide $725–745M (~7.0% of FY26 sales) fading toward 5.0%.",
    "sc_nwc_pct": "7.5% of revenue change for ΔNWC; ties working-capital swings to sales trajectory per historical sensitivity.",
    # DCF valuation
    "dcf_exitm": "8.0x FY2030E exit EV/EBITDA; mid-point of terminal football field (6.5–9.5x); ~1 turn above Gordon-implied ~7x.",
    # Comps — peer multiples
    "comps_nke": "Nike EV/EBITDA 12.0x equals StockAnalysis 11.97x, rounded; mature athletic benchmark.",
    "comps_deck": "Deckers EV/EBITDA 8.0x equals StockAnalysis 7.95x, rounded; closest premium-footwear peer.",
    "comps_onon": "On Holding EV/EBITDA 14.0x equals StockAnalysis 14.03x, rounded; high-growth athletic peer.",
    "comps_ads": "adidas (ADDYY) EV/EBITDA 9.3x equals StockAnalysis 9.31x; global incumbent.",
    "comps_vfc": "VFC EV/EBITDA 10.7x equals StockAnalysis 10.71x; challenged multi-brand apparel peer.",
    "comps_ff_ev_lo": "6.5x on FY2030E terminal EBITDA; bear exit below Gordon-implied ~7x; brackets DCF downside.",
    "comps_ff_ev_hi": "9.5x on FY2030E terminal EBITDA; bull exit above 8.0x DCF base; still below peer median ~15x.",
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
    "3s_capex_pct": "FY26 capex 7.0% matches 10-K guide midpoint ~$735M on $10.425B sales; fades 6.0%→5.0% after the build year.",
    "3s_sbc": "$62M SBC holds the FY25 10-K run-rate (Stock-based compensation expense 62,203).",
    "3s_inv_pct": "Inventory 33–34% of COGS; slight normalization from FY25 build as Americas demand softens and clears.",
    "3s_ap_pct": "AP 6.8% of COGS; holds near FY25 level reflecting vendor terms and production payment cadence.",
    "3s_accr_pct": "Accrued liabilities 5.8% of revenue; stable comp, marketing, and operating accrual ratio.",
    "3s_oca_pct": "OCA 6.8% = FY25 current assets 4,262,701 − cash 1,807,202 − inventories 1,700,753, as % of revenue (not the prepaid line alone).",
    "3s_rou_pct": "ROU assets 14.7% of revenue; lease-intensive store model, stable vs FY25 operating lease footprint.",
    "3s_onca_pct": "ONCA 3.0% = FY25 total assets − CA − PPE − ROU − GW (338,947); matches earnings 'deferred taxes and other NC assets'.",
    "3s_olc_pct": "Current lease liabilities 2.7% of revenue; short-term portion of operating lease obligations.",
    "3s_olnc_pct": "Non-current lease liabilities 13.5% of revenue; long-term store lease commitments per 10-K.",
    "3s_ocl_pct": "OCL 5.5% = FY25 current liabilities minus AP, accrued, and current leases (residual ~594k); not the 45,954 'other' line alone.",
    "3s_oncl_pct": "Other non-current liabilities 0.5% of revenue; minor long-term accruals and provisions.",
    "3s_buyback": "$500M annual repurchases; continued capital return at moderated pace vs FY24–25 peak buyback levels.",
    "3s_rep_price": "FY26 repurchase price $100 matches the current quote; path $100→$130 is a recovery assumption (FY25 10-K paid $168–$199).",
}

# Clickable source links for red assumptions (label, URL)
ASSUMPTION_SRC = {
    # WACC
    "wacc_rf": ("FRED: 10Y Treasury (DGS10)", SOURCES["fred_dgs10"]),
    "wacc_erp": ("Damodaran: Historical Implied ERP", SOURCES["damodaran_erp"]),
    "wacc_beta": ("StockAnalysis: LULU Beta (5Y)", SOURCES["lulu_stats"]),
    "wacc_kd": ("LULU FY2025 10-K (no funded debt)", filing_url("FY2025")),
    "wacc_tax": ("Q2 FY2026 outlook: tax rate ≈ 30%", SOURCES["earnings_sep2026"]),
    "wacc_we": ("LULU FY2025 10-K balance sheet", filing_url("FY2025")),
    "wacc_wd": ("LULU FY2025 10-K balance sheet", filing_url("FY2025")),
    # Scenarios
    "sc_g1": ("LULU Q2 FY2026 earnings release", SOURCES["earnings_sep2026"]),
    "sc_gterm": ("StockAnalysis: LULU 3Y revenue forecast", SOURCES["lulu_stats"]),
    "sc_m1": ("Q2 FY2026 release: 18.8% OM minus 560bps tariffs", SOURCES["earnings_sep2026"]),
    "sc_tariff": ("Q2 FY2026 release: $134.5M IEEPA tariff refunds", SOURCES["earnings_sep2026"]),
    "sc_mterm": ("FY2025 10-K: Income from operations / revenue", filing_url("FY2025")),
    "sc_wacc": ("WACC tab: CAPM build", None),  # internal link set in build_dcf.py
    "sc_g": ("FRED: Real GDP (GDPC1)", SOURCES["fred_gdpc1"]),
    "sc_tax": ("Q2 FY2026 outlook: tax rate ≈ 30%", SOURCES["earnings_sep2026"]),
    "sc_da_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "sc_capex_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "sc_nwc_pct": ("LULU BS/IS historical (10-K)", filing_url("FY2025")),
    # DCF / comps
    "dcf_exitm": ("StockAnalysis: LULU EV/EBITDA", SOURCES["lulu_stats"]),
    "comps_nke": ("StockAnalysis: NKE EV/EBITDA", "https://stockanalysis.com/stocks/nke/statistics/"),
    "comps_deck": ("StockAnalysis: DECK EV/EBITDA", "https://stockanalysis.com/stocks/deck/statistics/"),
    "comps_onon": ("StockAnalysis: ONON EV/EBITDA", "https://stockanalysis.com/stocks/onon/statistics/"),
    "comps_ads": ("StockAnalysis: adidas (ADR)", "https://stockanalysis.com/quote/otc/ADDYY/statistics/"),
    "comps_vfc": ("StockAnalysis: VFC EV/EBITDA", "https://stockanalysis.com/stocks/vfc/statistics/"),
    "comps_ff_ev_lo": ("StockAnalysis: LULU EV/EBITDA", SOURCES["lulu_stats"]),
    "comps_ff_ev_hi": ("StockAnalysis: LULU valuation", SOURCES["lulu_stats"]),
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
    "3s_oca_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_rou_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_onca_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_olc_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_olnc_pct": ("LULU lease disclosures (10-K)", filing_url("FY2025")),
    "3s_ocl_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_oncl_pct": ("LULU balance sheet (10-K)", filing_url("FY2025")),
    "3s_buyback": ("LULU CF: share repurchases (10-K)", filing_url("FY2025")),
    "3s_rep_price": ("NASDAQ LULU last sale (current price)", SOURCES["nasdaq_quote"]),
}

# Exact Ctrl+F strings — every quoted phrase appears verbatim in the linked source.
# Format: Ctrl+F "phrase" → what to read; model value noted where it is an assumption.
SOURCE_HINT = {
    # WACC
    "wacc_rf": 'Ctrl+F "2026-09-03" → observation 4.77. Model uses 4.8%. Also Ctrl+F "DGS10" for the series title.',
    "wacc_erp": 'Ctrl+F "4.23%" on the 2025 row (last data row). Header is "Implied ERP (FCFE)". Model uses 6.0%.',
    "wacc_beta": 'Ctrl+F "Beta (5Y)" → 0.86. Model uses 0.95.',
    "wacc_kd": 'Ctrl+F "no borrowings were outstanding under this facility" on this 10-K HTML page (search the document, do not use the EDGAR viewer TOC).',
    "wacc_tax": 'Ctrl+F "a tax rate of approximately 30%" on the FY2026 outlook paragraph.',
    "wacc_we": 'Ctrl+F "Cash and cash equivalents" → 1,807,202 ($000) | no funded term debt',
    "wacc_wd": 'Ctrl+F "no borrowings were outstanding under this facility" → debt weight 0%',
    # Scenarios
    "sc_g1": 'Ctrl+F "decline of 5% to 7%" → FY2026 revenue guide; model −6.1% midpoint',
    "sc_gterm": 'Ctrl+F "Revenue Growth Forecast (3Y)" → 2.26%. That is the unique line (do not search "Net revenue"). Model FY27–30 uses 2.3%.',
    "sc_m1": 'Ctrl+F "18.8%" (Q2 OM) and "560 basis points" (tariff boost). Real run-rate = 18.8% − 5.6% = 13.2%. Do not use "decreased 13%".',
    "sc_tariff": 'Ctrl+F "134.5 million" → IEEPA tariff refunds reduced COGS. Add this dollar amount to FY26 EBIT only. Full-year boost = 134.5 / FY26 sales ≈ 1.3ppt, not 5.6ppt.',
    "sc_mterm": 'Ctrl+F "Income from operations" → 2,210,615 and "Net revenue" → 11,102,600. FY25 OM = 19.9%. Model FY30 15.5% is a partial-recovery assumption, not a reported figure.',
    "sc_wacc": 'WACC tab → cell E (green) = rf 4.8% + β 0.95 × ERP 6.0% = 10.5%',
    "sc_g": 'Ctrl+F "Q2 2026" → 24,269.613 and "Q2 2025" → 23,770.976. YoY = 2.1%. Model terminal g 2.25%.',
    "sc_tax": 'Ctrl+F "a tax rate of approximately 30%" on the earnings outlook. FY25 10-K: "Income tax expense" 659,784 ÷ "Income before income tax expense" 2,238,967 = 29.5%.',
    "sc_da_pct": 'Ctrl+F "Depreciation and amortization" → 496,228 on this 10-K HTML (not the SEC viewer). ÷ "Net revenue" 11,102,600 = 4.5%.',
    "sc_capex_pct": 'Ctrl+F "725.0 million" → 2026 capex guide $725–745M (~7.0% of FY26 sales). Ctrl+F "680,802" → FY25 capex. Model 5-yr blend 5.5%.',
    "sc_nwc_pct": 'Ctrl+F "Inventories" → 1,700,753 | "Accounts payable" → 331,421 | "Accrued liabilities and other" → 662,982',
    # DCF / comps
    "dcf_exitm": 'Ctrl+F "EV / EBITDA" → LULU ~4.99x on page; model terminal exit 8.0x (assumption)',
    "comps_nke": 'Ctrl+F "EV / EBITDA" → 11.97. Model uses 12.0x.',
    "comps_deck": 'Ctrl+F "EV / EBITDA" → 7.95. Model uses 8.0x.',
    "comps_onon": 'Ctrl+F "EV / EBITDA" → 14.03. Model uses 14.0x.',
    "comps_ads": 'Ctrl+F "EV / EBITDA" → 9.31. Model uses 9.3x.',
    "comps_vfc": 'Ctrl+F "EV / EBITDA" → 10.71. Model uses 10.7x.',
    "comps_ff_ev_lo": 'Ctrl+F "EV / EBITDA" → LULU trough ~5x; bear terminal exit assumption 6.5x',
    "comps_ff_ev_hi": 'Ctrl+F "EV / EBITDA" → bull terminal exit assumption 9.5x on FY2030E EBITDA',
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
    "3s_capex_pct": 'Ctrl+F "725.0 million" and "745.0 million" on this 10-K HTML. Midpoint $735M / FY26 sales $10.425B ≈ 7.0% (FY26 model).',
    "3s_sbc": 'Ctrl+F "Stock-based compensation expense" → 62,203 on this 10-K HTML. Model $62,000.',
    "3s_inv_pct": 'Ctrl+F "Inventories" → 1,700,753 ÷ "Cost of goods sold" 4,818,468 = 35.3%',
    "3s_ap_pct": 'Ctrl+F "Accounts payable" → 331,421 ÷ "Cost of goods sold" 4,818,468 = 6.9%',
    "3s_accr_pct": 'Ctrl+F "Accrued liabilities and other" → 662,982 ÷ "Net revenue" 11,102,600 = 6.0%',
    "3s_oca_pct": 'Ctrl+F "Total current assets" → 4,262,701 minus "Cash and cash equivalents" 1,807,202 minus "Inventories" 1,700,753 = 754,746 (6.8% of revenue)',
    "3s_rou_pct": 'Ctrl+F "Right-of-use lease assets" → 1,630,181 ÷ "Net revenue" 11,102,600 = 14.7%',
    "3s_onca_pct": 'Ctrl+F "Total assets" 8,456,743 − current assets − "Property and equipment, net" − "Right-of-use lease assets" − goodwill/intangibles = 338,947 (3.0%)',
    "3s_olc_pct": 'Ctrl+F "Current lease liabilities" → 298,724 ÷ "Net revenue" 11,102,600 = 2.7%',
    "3s_olnc_pct": 'Ctrl+F "Non-current lease liabilities" → 1,499,717 ÷ "Net revenue" 11,102,600 = 13.5%',
    "3s_ocl_pct": 'Ctrl+F "Total current liabilities" 1,887,548 − "Accounts payable" 331,421 − "Accrued liabilities and other" 662,982 − "Current lease liabilities" 298,724 ≈ 594,421 (5.4%)',
    "3s_oncl_pct": 'Ctrl+F "Other non-current liabilities" → 55,360 ÷ "Net revenue" 11,102,600 = 0.5%',
    "3s_buyback": 'Ctrl+F "Repurchase of common stock" → ( 1,178,349 ) ($000) FY25; model (500,000)/yr',
    "3s_rep_price": 'Ctrl+F "LULU" → Last Sale / Previous Close (~$100). FY26 model repurchase price starts at $100.',
}

COVER_HINTS = {
    "edgar_xbrl": 'Ctrl+F "10-K" → FY2025 accession 0001397187-26-000020',
    "filing_fy2025": 'Ctrl+F "Net revenue" → 11,102,600 or "Depreciation and amortization" → 496,228 on this 10-K HTML file',
    "earnings_sep2026": 'Ctrl+F "decline of 5% to 7%" | "$10.350 billion to $10.500 billion" | "$9.48 to $9.73"',
    "nasdaq_quote": 'Ctrl+F "LULU" → Last Sale or Previous Close price',
}

REPORTED_HINTS = {
    "10k": 'Ctrl+F "CONSOLIDATED STATEMENTS OF OPERATIONS" → locate line item in FY2025 column',
    "10k_is": 'Ctrl+F "Net revenue" → 11,102,600 ($000) in FY2025 column',
    "10k_ebit": 'Ctrl+F "Income from operations" → 2,210,615 ($000) FY2025',
    "10k_bs": 'Ctrl+F "Cash and cash equivalents" → 1,807,202 ($000) FY2025',
    "10k_debt": 'Ctrl+F "no borrowings were outstanding under this facility" → no funded term debt',
    "10k_cf": 'Ctrl+F "Depreciation and amortization" → 496,228 | "680,802" capex ($000)',
    "10k_shares": 'Ctrl+F "Diluted weighted-average number of shares outstanding" → 119,068 (000)',
    "10k_ebitda": 'Ctrl+F "Income from operations" → 2,210,615 + "Depreciation and amortization" → 496,228',
    "earnings_rev": 'Ctrl+F "$10.350 billion to $10.500 billion" → FY2026 net revenue guidance',
    "earnings_eps": 'Ctrl+F "$9.48 to $9.73" → FY2026 diluted EPS guidance; model midpoint $9.61',
    "nasdaq": 'Ctrl+F "LULU" → Last Sale or Previous Close on NASDAQ quote page',
}
