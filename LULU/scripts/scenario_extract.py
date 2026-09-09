"""Extract pitch IS / BS / CF from DCF Scenarios tab (cols G base, H bull)."""
import data as D

PROJ_YEARS = D.PROJ_YEARS
COL_BASE = 7   # G
COL_BULL = 8   # H


def _year_rows(sc, needle):
    rows = {}
    for r in range(1, 400):
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
    if v is None or (isinstance(v, str) and v.startswith("#")):
        return None
    return round(v / 1000)


def _pct(v, d=1):
    if v is None or (isinstance(v, str) and v.startswith("#")):
        return None
    return f"{v * 100:.{d}f}%"


def _eps(v):
    if v is None or (isinstance(v, str) and v.startswith("#")):
        return None
    return round(v, 2) if isinstance(v, (int, float)) else None


def align_income_statement(is_paths):
    for fy in PROJ_YEARS:
        oi = is_paths[fy]["operating_income"]
        rev = is_paths[fy]["revenue"]
        if oi is not None and rev:
            is_paths[fy]["operating_margin"] = f"{oi / rev * 100:.1f}%"


def align_cash_flow(cf_paths):
    for fy in PROJ_YEARS:
        cfo = cf_paths.get("cfo", {}).get(fy)
        capex = cf_paths.get("capex", {}).get(fy)
        if cfo is not None and capex is not None:
            cf_paths["fcf"][fy] = cfo + capex


# Matches 3-statement Assumptions repurchase price path ($/sh).
_REPURCHASE_PRICES = [100, 108, 115, 122, 130]
_BUYBACK_BUDGET_M = 500  # $500M/yr base case (Scenarios col G bb_fixed)


def align_buyback_eps(income, cash_flow):
    """Recompute forecast EPS from NI and buyback-reduced diluted share count.

    Scenarios pitch-bridge EPS had a units bug (NI×1000/shares). Forecast EPS
    should match the 3-statement: NI ($M) ÷ diluted shares (000), with shares
    reduced by $500M/yr repurchases at the assumed avg price path.
    """
    sh_k = D.IS["diluted_shares"]["FY2025"]
    for i, fy in enumerate(PROJ_YEARS):
        ni_m = income[fy].get("net_income")
        if ni_m is None:
            continue
        bb_m = abs(cash_flow["buybacks"].get(fy) or _BUYBACK_BUDGET_M)
        price = _REPURCHASE_PRICES[i]
        retired_k = bb_m * 1000 / price
        sh_k = max(sh_k - retired_k, 1)
        income[fy]["eps"] = round((ni_m * 1000) / sh_k, 2)
        income[fy]["diluted_shares_k"] = round(sh_k)


def extract_scenario(sc, col):
    """Read full pitch deck paths from Scenarios pitch-bridge block + forecast rows."""
    rev_r = _year_rows(sc, "revenue")
    gm_r = _year_rows(sc, "gross margin %")
    ebit_r = _year_rows(sc, "ebit")
    ni_r = _year_rows(sc, "net income")
    cfo_r = _year_rows(sc, "cash from operations")
    fcf_r = _year_rows(sc, "free cash flow (cfs)")
    buy_r = _year_rows(sc, "share repurchases")
    cash_r = _year_rows(sc, "cash & equivalents")
    inv_r = _year_rows(sc, "inventories")
    ta_r = _year_rows(sc, "total assets")
    tl_r = _year_rows(sc, "total liabilities")
    te_r = _year_rows(sc, "total equity")
    da_r = _year_rows(sc, "d&a")
    capex_r = _year_rows(sc, "capex")
    eps_r = _year_rows(sc, "diluted eps")

    required = {
        "revenue": rev_r, "net income": ni_r, "cash from operations": cfo_r,
        "free cash flow (cfs)": fcf_r, "share repurchases": buy_r,
        "cash & equivalents": cash_r, "total assets": ta_r,
    }
    for name, rows in required.items():
        if len(rows) < 5:
            raise KeyError(f"Scenarios pitch bridge missing rows for '{name}' (found {len(rows)})")

    per_fy = {}
    for t, fy in enumerate(PROJ_YEARS, start=1):
        rev = sc.cell(rev_r[t], col).value
        gm = sc.cell(gm_r[t], col).value
        ebit = sc.cell(ebit_r[t], col).value
        capex = sc.cell(capex_r[t], col).value or 0
        per_fy[fy] = {
            "revenue": _m(rev),
            "gross_profit": _m(rev * gm) if rev and gm else None,
            "operating_income": _m(ebit),
            "operating_margin": _pct(ebit / rev) if rev and ebit else None,
            "net_income": _m(sc.cell(ni_r[t], col).value),
            "eps": _eps(sc.cell(eps_r[t], col).value),
            "cash": _m(sc.cell(cash_r[t], col).value),
            "inventories": _m(sc.cell(inv_r[t], col).value),
            "total_assets": _m(sc.cell(ta_r[t], col).value),
            "total_liab": _m(sc.cell(tl_r[t], col).value),
            "total_equity": _m(sc.cell(te_r[t], col).value),
            "cfo": _m(sc.cell(cfo_r[t], col).value),
            "capex": _m(-capex),
            "fcf": _m(sc.cell(fcf_r[t], col).value),
            "buybacks": _m(sc.cell(buy_r[t], col).value),
            "dna": _m(sc.cell(da_r[t], col).value),
        }

    income = {
        fy: {k: per_fy[fy][k] for k in (
            "revenue", "gross_profit", "operating_income", "operating_margin", "net_income", "eps"
        )}
        for fy in PROJ_YEARS
    }
    balance = {
        key: {fy: per_fy[fy][key] for fy in PROJ_YEARS}
        for key in ("cash", "inventories", "total_assets", "total_liab", "total_equity")
    }
    cash_flow = {
        key: {fy: per_fy[fy][key] for fy in PROJ_YEARS}
        for key in ("cfo", "capex", "fcf", "buybacks", "dna")
    }
    align_income_statement(income)
    align_cash_flow(cash_flow)
    align_buyback_eps(income, cash_flow)
    return income, balance, cash_flow


# Back-compat alias
def extract_paths(sc, col):
    income, balance, cash_flow = extract_scenario(sc, col)
    return {
        fy: {**income[fy], **{k: balance[k][fy] for k in balance}, **{k: cash_flow[k][fy] for k in cash_flow}}
        for fy in PROJ_YEARS
    }


_REF_KEYS = {
    "revenue": "revenue",
    "gross_margin": "gross margin %",
    "ebit": "ebit",
    "net_income": "net income",
    "cfo": "cash from operations",
    "fcf": "free cash flow (cfs)",
    "buybacks": "share repurchases",
    "cash": "cash & equivalents",
    "inventories": "inventories",
    "total_assets": "total assets",
    "total_liab": "total liabilities",
    "total_equity": "total equity",
    "dna": "d&a",
    "capex": "capex",
    "eps": "diluted eps",
}


def discover_model_refs(sc):
    """Map pitch line items to Scenarios tab row ranges (cols F/G/H)."""
    refs = {}
    for key, needle in _REF_KEYS.items():
        rows = _year_rows(sc, needle)
        if len(rows) < 5:
            continue
        label = str(sc.cell(rows[1], 1).value or "").strip()
        refs[key] = {
            "r1": rows[1],
            "r5": rows[5],
            "label": label,
        }
    return refs


def scen_range(refs, key, col="G"):
    """Excel-style range for base-case column G."""
    r = refs[key]
    return f"Scenarios!{col}{r['r1']}:{col}{r['r5']}"
