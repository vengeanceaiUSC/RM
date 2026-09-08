"""Extract 5-year bear/base/bull paths from the DCF Scenarios tab."""
import data as D

PROJ_YEARS = D.PROJ_YEARS
COL_BASE = 7   # G
COL_BULL = 8   # H


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


def extract_paths(sc, col):
    rev_r = _year_rows(sc, "revenue")
    gm_r = _year_rows(sc, "gross margin %")
    ebit_r = _year_rows(sc, "ebit")
    nopat_r = _year_rows(sc, "nopat")
    da_r = _year_rows(sc, "d&a")
    capex_r = _year_rows(sc, "capex")
    dnwc_r = _year_rows(sc, "δnwc")
    ufcf_r = _year_rows(sc, "unlevered fcf")

    out = {}
    for t, fy in enumerate(PROJ_YEARS, start=1):
        rev = sc.cell(rev_r[t], col).value
        gm = sc.cell(gm_r[t], col).value
        ebit = sc.cell(ebit_r[t], col).value
        nopat = sc.cell(nopat_r[t], col).value
        da = sc.cell(da_r[t], col).value
        capex = sc.cell(capex_r[t], col).value
        dnwc = sc.cell(dnwc_r[t], col).value if dnwc_r else 0
        ufcf = sc.cell(ufcf_r[t], col).value if ufcf_r else None
        gp = rev * gm if rev and gm else None
        om = ebit / rev if rev and ebit else None
        cfo = (nopat or 0) + (da or 0) + (dnwc or 0)
        buyback = -round(0.75 * ufcf) if ufcf else None
        out[fy] = {
            "revenue": _m(rev),
            "gross_profit": _m(gp),
            "operating_income": _m(ebit),
            "operating_margin": _pct(om),
            "net_income": _m(nopat),
            "cfo": _m(cfo),
            "capex": _m(-capex) if capex else None,
            "fcf": _m(ufcf),
            "buybacks": buyback,
            "dna": _m(da),
            "nopat_m": _m(nopat),
        }
    return out


def bull_balance_sheet(base_bs, base_is, bull_is, start_cash_m=1807):
    """Approximate bull BS from base 3-statement scaled at bull revenue + UFCF cash build."""
    cash = start_cash_m
    out = {k: {} for k in base_bs}
    for fy in PROJ_YEARS:
        ratio = bull_is[fy]["revenue"] / base_is[fy]["revenue"] if base_is[fy]["revenue"] else 1
        ufcf = bull_is[fy].get("fcf") or 0
        buy = bull_is[fy].get("buybacks") or 0
        cash = cash + ufcf + buy  # buybacks negative
        out["cash"][fy] = round(cash)
        for key in ("inventories", "total_assets", "total_liab", "total_equity"):
            out[key][fy] = round(base_bs[key][fy] * ratio)
    return out


def add_eps(is_paths, shares_by_year):
    for fy in PROJ_YEARS:
        ni = is_paths[fy]["net_income"]
        sh = shares_by_year.get(fy)
        is_paths[fy]["eps"] = round(ni * 1000 / sh, 2) if ni and sh else None
