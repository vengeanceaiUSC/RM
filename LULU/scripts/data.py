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
}

# FY label -> (accession, primary 10-K document)
FILING_10K = {
    "FY2022": ("0001397187-23-000012", "lulu-20230129.htm"),
    "FY2023": ("0001397187-24-000010", "lulu-20240128.htm"),
    "FY2024": ("0001397187-25-000013", "lulu-20250202.htm"),
    "FY2025": ("0001397187-26-000020", "lulu-20260201.htm"),
}

def filing_url(fy):
    """SEC EDGAR interactive viewer for the Form 10-K (reliable in browser/Excel)."""
    acc, _ = FILING_10K[fy]
    return f"https://www.sec.gov/cgi-bin/viewer?action=view&cik=1397187&accession_number={acc}&xbrl_type=v"

# ---------------------------------------------------------------------------
# Red-assumption justifications (~20 words each) for DCF / 3-statement models
# ---------------------------------------------------------------------------
JUST = {
    # WACC tab
    "wacc_rf": "4.3% risk-free anchored to FRED 10Y Treasury (DGS10); rounded early-Sep-2026 level for DCF stability.",
    "wacc_erp": "6.0% additive ERP; conservative vs Damodaran implied ~4.2% (Jan-26); within long-run historical US range.",
    "wacc_beta": "0.95 levered beta vs ~0.86 market; modest uplift for near-term earnings volatility and post-guidance de-rating.",
    "wacc_kd": "5.0% illustrative pre-tax debt cost; LULU has no funded term debt, only an undrawn revolving credit facility.",
    "wacc_tax": "27% normalized marginal cash tax rate; mid-point between recent effective rates and US federal-plus-state blend.",
    "wacc_we": "100% equity weight; net-cash balance sheet with no material funded debt at market values per FY2025 10-K.",
    "wacc_wd": "0% debt weight; no outstanding term loans or bonds, so WACC effectively equals levered cost of equity here.",
    # Scenarios tab
    "sc_g1": "−6.1% FY2026 revenue growth matches company guidance midpoint of −5% to −7% after Q2 FY2026 print.",
    "sc_gterm": "2.8% avg FY27–30 growth assumes gradual recovery; well below historical double-digit LULU revenue expansion.",
    "sc_m1": "13.9% FY2026 EBIT margin reflects Americas weakness and guided compression; above bear, below historical peak.",
    "sc_mterm": "15.5% FY2030 EBIT margin assumes partial recovery; still materially below ~20–24% peak operating margins.",
    "sc_wacc": "10.0% base WACC ties to CAPM build on WACC tab; drives scenario DCFs and base-case valuation.",
    "sc_g": "2.25% terminal growth approximates long-run GDP plus inflation; conservative perpetuity rate for mature apparel.",
    "sc_tax": "27% cash tax rate for scenario FCF; consistent with WACC normalized rate and DCF unlevered cash conversion.",
    "sc_da_pct": "4.5% D&A to revenue near FY22–25 average; reflects store fleet, distribution, and technology amortization load.",
    "sc_capex_pct": "5.0% capex to revenue aligned with recent investment intensity and continued global store expansion plans.",
    "sc_nwc_pct": "7.5% of revenue change for ΔNWC; ties working-capital swings to sales trajectory per historical sensitivity.",
    # DCF valuation
    "dcf_exitm": "8.0x FY2030E exit EV/EBITDA; mid-point of terminal football field (6.5–9.5x); ~1 turn above Gordon-implied ~7x.",
    # Comps — peer multiples
    "comps_nke": "Nike ~18x forward EV/EBITDA illustrative; mature global athletic benchmark with slower growth than LULU peak.",
    "comps_deck": "Deckers ~15x reference; premium footwear peer with HOKA/UGG momentum and strong brand heat.",
    "comps_onon": "On Holding ~25x; high-growth athletic peer setting upper bound for premium positioning and white space.",
    "comps_ads": "adidas ~12x; global incumbent in restructuring with moderate growth and complex brand portfolio.",
    "comps_vfc": "VFC ~10x; challenged multi-brand apparel operator representing lower bound for scaled apparel peers.",
    "comps_ff_ev_lo": "6.5x on FY2030E terminal EBITDA; bear exit below Gordon-implied ~7x; brackets DCF downside.",
    "comps_ff_ev_hi": "9.5x on FY2030E terminal EBITDA; bull exit above 8.0x DCF base; still below peer median ~15x.",
    "comps_ff_pe_lo": "10x P/E low on FY2026E EPS; trough earnings multiple after guidance reset and sentiment de-rating.",
    "comps_ff_pe_hi": "18x P/E high on FY2026E EPS; modest recovery case still below historical premium LULU multiples.",
    # Sensitivity axes (summary)
    "sens_axes": "Red WACC and g grid values bracket base case ±100bps discount rate and ±75bps terminal growth for sensitivity.",
    "sens_wacc": "WACC axis 9.0–11.0% brackets 10.0% base from CAPM (rf + β×ERP) on WACC tab; ±100bps sensitivity band.",
    "sens_g": "Terminal-g axis 1.5–3.0% brackets 2.25% base; bounded by long-run real GDP and inflation benchmarks.",
    # 3-statement Assumptions tab
    "3s_rev_growth": "FY26 −6.1% matches guidance; FY27–30 step up to low-single-digit then mid-single-digit recovery path.",
    "3s_gm": "Gross margin recovers gradually from promo pressure; 56.5% to 58.0% still below peak ~58–59% historical.",
    "3s_sga_pct": "SG&A leverage improves slowly as revenue stabilizes; ratio declines toward 39.5% by FY2030 from cost discipline.",
    "3s_other_opex": "$7M annual amortization run-rate; stable intangible amortization per recent 10-K disclosure levels.",
    "3s_other_inc": "Interest income declines as cash is deployed; $130M FY26 stepping down with lower cash balances.",
    "3s_tax_rate": "30% effective tax rate on projections; conservative vs recent ~29% effective, allows for jurisdictional mix.",
    "3s_da_pct": "D&A 4.5–4.6% of revenue; tracks recent depreciation intensity on PPE and lease-related amortization.",
    "3s_capex_pct": "Capex fades from 5.5% to 5.0%; reflects completion of major supply-chain projects and normalized store growth.",
    "3s_sbc": "$70M annual SBC; modest decline from FY25 as headcount growth slows in restructuring period.",
    "3s_inv_pct": "Inventory 33–34% of COGS; slight normalization from FY25 build as Americas demand softens and clears.",
    "3s_ap_pct": "AP 6.8% of COGS; holds near FY25 level reflecting vendor terms and production payment cadence.",
    "3s_accr_pct": "Accrued liabilities 5.8% of revenue; stable comp, marketing, and operating accrual ratio.",
    "3s_oca_pct": "Other current assets 6.8% of revenue; prepaid and receivable balance consistent with recent history.",
    "3s_rou_pct": "ROU assets 14.7% of revenue; lease-intensive store model, stable vs FY25 operating lease footprint.",
    "3s_onca_pct": "Other non-current assets 3.0% of revenue; deferred costs and long-term deposits at normalized level.",
    "3s_olc_pct": "Current lease liabilities 2.7% of revenue; short-term portion of operating lease obligations.",
    "3s_olnc_pct": "Non-current lease liabilities 13.5% of revenue; long-term store lease commitments per 10-K.",
    "3s_ocl_pct": "Other current liabilities 5.5% of revenue; gift cards, deferred revenue, and other short-term obligations.",
    "3s_oncl_pct": "Other non-current liabilities 0.5% of revenue; minor long-term accruals and provisions.",
    "3s_buyback": "$500M annual repurchases; continued capital return at moderated pace vs FY24–25 peak buyback levels.",
    "3s_rep_price": "Repurchase price rises with recovery thesis; $105 to $135 reflects assumed gradual share-price normalization.",
}

# Clickable source links for red assumptions (label, URL)
ASSUMPTION_SRC = {
    # WACC
    "wacc_rf": ("FRED: 10Y Treasury (DGS10)", SOURCES["fred_dgs10"]),
    "wacc_erp": ("Damodaran: Historical Implied ERP", SOURCES["damodaran_erp"]),
    "wacc_beta": ("Yahoo Finance: LULU Beta", "https://finance.yahoo.com/quote/LULU/key-statistics/"),
    "wacc_kd": ("LULU FY2025 10-K (no funded debt)", filing_url("FY2025")),
    "wacc_tax": ("Damodaran: US tax rate dataset", SOURCES["damodaran_tax"]),
    "wacc_we": ("LULU FY2025 10-K balance sheet", filing_url("FY2025")),
    "wacc_wd": ("LULU FY2025 10-K balance sheet", filing_url("FY2025")),
    # Scenarios
    "sc_g1": ("LULU Q2 FY2026 earnings release", SOURCES["earnings_sep2026"]),
    "sc_gterm": ("LULU historical revenue (10-K)", filing_url("FY2025")),
    "sc_m1": ("LULU Q2 FY2026 earnings release", SOURCES["earnings_sep2026"]),
    "sc_mterm": ("LULU historical EBIT margins (10-K)", filing_url("FY2025")),
    "sc_wacc": ("WACC tab: CAPM build", None),  # internal link set in build_dcf.py
    "sc_g": ("FRED: Real GDP (GDPC1)", SOURCES["fred_gdpc1"]),
    "sc_tax": ("Damodaran: US tax rate dataset", SOURCES["damodaran_tax"]),
    "sc_da_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "sc_capex_pct": ("LULU CF statement (10-K)", filing_url("FY2025")),
    "sc_nwc_pct": ("LULU BS/IS historical (10-K)", filing_url("FY2025")),
    # DCF / comps
    "dcf_exitm": ("StockAnalysis: LULU EV/EBITDA", SOURCES["lulu_stats"]),
    "comps_nke": ("Yahoo Finance: NKE statistics", "https://finance.yahoo.com/quote/NKE/key-statistics/"),
    "comps_deck": ("Yahoo Finance: DECK statistics", "https://finance.yahoo.com/quote/DECK/key-statistics/"),
    "comps_onon": ("Yahoo Finance: ONON statistics", "https://finance.yahoo.com/quote/ONON/key-statistics/"),
    "comps_ads": ("Yahoo Finance: adidas (ADDYY)", "https://finance.yahoo.com/quote/ADDYY/key-statistics/"),
    "comps_vfc": ("Yahoo Finance: VFC statistics", "https://finance.yahoo.com/quote/VFC/key-statistics/"),
    "comps_ff_ev_lo": ("StockAnalysis: LULU EV/EBITDA", SOURCES["lulu_stats"]),
    "comps_ff_ev_hi": ("StockAnalysis: LULU valuation", SOURCES["lulu_stats"]),
    "comps_ff_pe_lo": ("Yahoo Finance: LULU P/E", "https://finance.yahoo.com/quote/LULU/key-statistics/"),
    "comps_ff_pe_hi": ("Yahoo Finance: NKE P/E (peer)", "https://finance.yahoo.com/quote/NKE/key-statistics/"),
    "sens_axes": ("Scenarios: WACC & terminal g", None),  # internal links set in build_dcf.py
    "sens_wacc": ("Damodaran: Historical Implied ERP", SOURCES["damodaran_erp"]),
    "sens_g": ("FRED: Real GDP (GDPC1)", SOURCES["fred_gdpc1"]),
    # 3-statement
    "3s_rev_growth": ("LULU Q2 FY2026 guidance", SOURCES["earnings_sep2026"]),
    "3s_gm": ("LULU historical gross margin (10-K)", filing_url("FY2025")),
    "3s_sga_pct": ("LULU historical SG&A (10-K)", filing_url("FY2025")),
    "3s_other_opex": ("LULU FY2025 10-K", filing_url("FY2025")),
    "3s_other_inc": ("LULU FY2025 10-K", filing_url("FY2025")),
    "3s_tax_rate": ("LULU effective tax history (10-K)", filing_url("FY2025")),
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
    "3s_rep_price": ("LULU avg repurchase price (10-K/est.)", filing_url("FY2025")),
}
