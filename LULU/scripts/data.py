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
    "edgar_company": "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001397187&type=10-K",
    "edgar_xbrl": "https://data.sec.gov/api/xbrl/companyfacts/CIK0001397187.json",
    "earnings_sep2026": "https://investor.lululemon.com/news-releases",
    "stock_info": "https://investor.lululemon.com/stock-information",
    "nasdaq_quote": "https://www.nasdaq.com/market-activity/stocks/lulu",
}

# FY label -> (accession, primary 10-K document)
FILING_10K = {
    "FY2022": ("0001397187-23-000012", "lulu-20230129.htm"),
    "FY2023": ("0001397187-24-000010", "lulu-20240128.htm"),
    "FY2024": ("0001397187-25-000013", "lulu-20250202.htm"),
    "FY2025": ("0001397187-26-000020", "lulu-20260201.htm"),
}

def filing_url(fy):
    acc, doc = FILING_10K[fy]
    return _sec_10k(acc, doc)
