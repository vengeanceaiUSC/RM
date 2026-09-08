"""Extract 5-year bear/base/bull paths from the DCF Scenarios tab."""
import data as D

PROJ_YEARS = D.PROJ_YEARS
COL_BASE = 7   # G
COL_BULL = 8   # H

# 3-statement other-income path ($k) — same for base/bull forecast bridge
OTHER_INC_K = [45000, 40000, 35000, 30000, 25000]


def _year_rows(sc, needle):
    rows = {}
    for r in range(1, 200):
        lab = sc.cell(r, 1).value
        if not lab:
            continue
        s = str(lab).strip().lower()
        if needle in s and "year" in s:
            for t in range(1, 6):
                if f"year {t}" in s:
                    rows[t] = r
    return rows


def _m(v):
    return round(v / 1000) if v is not None else None


def _pct(v, d=1):
    return f"{v * 100:.{d}f}%" if v is not None else None


def _scenario_tax_rate(sc, col):
    return sc.cell(11, col).value or 0.30


def extract_paths(sc, col):
    """Build IS + CF metrics using the same definitions as the 3-statement / 10-K slides."""
    rev_r = _year_rows(sc, "revenue")
    gm_r = _year_rows(sc, "gross margin %")
    ebit_r = _year_rows(sc, "ebit")
    da_r = _year_rows(sc, "d&a")
    capex_r = _year_rows(sc, "capex")
    dnwc_r = _year_rows(sc, "δnwc")

    tax_rate = _scenario_tax_rate(sc, col)
    out = {}
    for t, fy in enumerate(PROJ_YEARS, start=1):
        rev = sc.cell(rev_r[t], col).value
        gm = sc.cell(gm_r[t], col).value
        ebit = sc.cell(ebit_r[t], col).value
        da = sc.cell(da_r[t], col).value
        capex = sc.cell(capex_r[t], col).value
        dnwc = sc.cell(dnwc_r[t], col).value if dnwc_r else 0

        gp = rev * gm if rev and gm else None
        oi = ebit
        om = oi / rev if rev and oi else None

        # Net income bridge (matches 3-statement: pretax = EBIT + other income, then tax)
        other = OTHER_INC_K[t - 1]
        pretax = (oi or 0) + other
        tax_exp = pretax * tax_rate
        ni = pretax - tax_exp

        # Cash flow — same indirect-method lines as historical / 3-statement CFS
        cfo = ni + (da or 0) + (dnwc or 0)
        capex_neg = -(capex or 0)
        fcf = cfo + capex_neg
        buyback_k = -round(0.75 * fcf) if fcf and fcf > 0 else 0

        out[fy] = {
            "revenue": _m(rev),
            "gross_profit": _m(gp),
            "operating_income": _m(oi),
            "operating_margin": _pct(om),
            "net_income": _m(ni),
            "cfo": _m(cfo),
            "capex": _m(capex_neg),
            "fcf": _m(fcf),
            "buybacks": _m(buyback_k) if buyback_k else 0,
            "dna": _m(da),
        }
    return out


def bull_balance_sheet(base_bs, base_is, bull_is, bull_cf, start_cash_m=1807):
    """Bull BS: same line items as base; cash from FY25 + cumulative (FCF + buybacks)."""
    cash = start_cash_m
    out = {k: {} for k in base_bs}
    for fy in PROJ_YEARS:
        ratio = bull_is[fy]["revenue"] / base_is[fy]["revenue"] if base_is[fy]["revenue"] else 1
        fcf = bull_cf["fcf"].get(fy) or 0
        buy = bull_cf["buybacks"].get(fy) or 0
        cash = cash + fcf + buy
        out["cash"][fy] = round(cash)
        for key in ("inventories", "total_assets", "total_liab", "total_equity"):
            out[key][fy] = round(base_bs[key][fy] * ratio)
    return out


def add_eps(is_paths, shares_by_year):
    for fy in PROJ_YEARS:
        ni = is_paths[fy]["net_income"]
        sh = shares_by_year.get(fy)
        is_paths[fy]["eps"] = round(ni * 1000 / sh, 2) if ni and sh else None


def align_income_statement(is_paths):
    """Operating margin % = operating income / revenue on every row (hist/base/bull)."""
    for fy in PROJ_YEARS:
        oi = is_paths[fy]["operating_income"]
        rev = is_paths[fy]["revenue"]
        if oi is not None and rev:
            is_paths[fy]["operating_margin"] = f"{oi / rev * 100:.1f}%"


def align_cash_flow(cf_paths):
    """FCF = CFO + capex (capex negative) — same as historical 10-K presentation."""
    for fy in PROJ_YEARS:
        cfo = cf_paths.get("cfo", {}).get(fy)
        capex = cf_paths.get("capex", {}).get(fy)
        if cfo is not None and capex is not None:
            cf_paths["fcf"][fy] = cfo + capex
