"""Pitch-deck figures sourced from the live DCF + 3-statement models.

Regenerate after model changes:
  cd LULU/scripts && python3 pitch_values.py
"""
import json
import os
import subprocess

import openpyxl

import data as D
from scenario_links import sync_mirror_sheet
from scenario_extract import (
    extract_paths,
    bull_balance_sheet,
    add_eps,
    align_income_statement,
    align_cash_flow,
    COL_BULL,
    PROJ_YEARS,
)

ROOT = os.path.join(os.path.dirname(__file__), "..")
DCF_PATH = os.path.join(ROOT, "LULU_DCF_Valuation_Model.xlsx")
IS_PATH = os.path.join(ROOT, "LULU_3_Statement_Model.xlsx")
OUT_JSON = os.path.join(ROOT, "data", "pitch_values.json")
PROJ_YEARS = D.PROJ_YEARS
PROJ_COL = {fy: 9 + i for i, fy in enumerate(PROJ_YEARS)}


def _recalc(path):
    import shutil
    from recalc_workbook import recalc_workbook

    tmp = "/tmp/lulu_models"
    os.makedirs(tmp, exist_ok=True)
    base = os.path.basename(path)
    dest = os.path.join(tmp, base)
    shutil.copy2(path, dest)
    if "3_Statement" in base:
        dcf_dest = os.path.join(tmp, os.path.basename(DCF_PATH))
        shutil.copy2(DCF_PATH, dcf_dest)
        dcf_wb = recalc_workbook(dcf_dest)
        dcf_wb = openpyxl.load_workbook(dcf_dest, data_only=True)
        sync_mirror_sheet(dest, dcf_wb)
        recalc_workbook(dest)
    else:
        recalc_workbook(dest)
    return openpyxl.load_workbook(dest, data_only=True)


def _m(v):
    if v is None:
        return None
    if isinstance(v, str):
        if v.startswith("#"):
            raise ValueError(f"Unrecalculated model value: {v}")
        return v
    return round(v / 1000)


def _pct(v, d=1):
    if v is None or (isinstance(v, str) and v.startswith("#")):
        return None
    if isinstance(v, str):
        return v
    return f"{v * 100:.{d}f}%"


def _num(v):
    if v is None or (isinstance(v, str) and v.startswith("#")):
        return None
    if isinstance(v, str):
        return v
    return v


def _year_slice(ws, row, years=PROJ_YEARS):
    return {fy: ws.cell(row, PROJ_COL[fy]).value for fy in years}


def extract():
    dcf_wb = _recalc(DCF_PATH)
    is_wb = _recalc(IS_PATH)
    dcf = dcf_wb["DCF"]
    sc = dcf_wb["Scenarios"]
    wacc = dcf_wb["WACC"]
    comps = dcf_wb["Comps"]
    is_ws = is_wb["Income Statement"]
    bs = is_wb["Balance Sheet"]
    cf = is_wb["Cash Flow"]

    pt_row = next(r for r in range(1, 500) if sc.cell(r, 1).value == "Implied share price")
    bear = sc.cell(pt_row, 6).value
    base = sc.cell(pt_row, 7).value
    bull = sc.cell(pt_row, 8).value

    def is_row(label):
        for r in range(1, 60):
            if is_ws.cell(r, 1).value == label:
                return r
        raise KeyError(label)

    def bs_row(label):
        for r in range(1, 60):
            if bs.cell(r, 1).value == label:
                return r
        raise KeyError(label)

    def cf_row(label):
        for r in range(1, 60):
            if cf.cell(r, 1).value == label:
                return r
        raise KeyError(label)

    r_rev = is_row("Net revenue")
    r_gp = is_row("Gross profit")
    r_oi = is_row("Operating income (EBIT)")
    r_om = is_row("Operating margin % (reported, incl. FY26 refund)")
    r_ni = is_row("Net income")
    r_eps = is_row("Diluted EPS ($)")

    r_sh = is_row("Diluted weighted-avg shares (000)")

    sens = []
    for r in range(98, 103):
        w = dcf.cell(r, 1).value
        if isinstance(w, (int, float)):
            sens.append(
                {
                    "wacc": f"{w * 100:.1f}%",
                    "prices": [dcf.cell(r, c).value for c in range(6, 11)],
                }
            )

    ff = {}
    for r in range(55, 65):
        label = comps.cell(r, 1).value
        if label and comps.cell(r, 7).value is not None:
            ff[str(label).split("(")[0].strip()] = {
                "low": comps.cell(r, 7).value,
                "high": comps.cell(r, 8).value,
            }

    income_base = {}
    for fy in PROJ_YEARS:
        col = PROJ_COL[fy]
        eps_val = is_ws.cell(r_eps, col).value
        oi_m = _m(is_ws.cell(r_oi, col).value)
        rev_m = _m(is_ws.cell(r_rev, col).value)
        income_base[fy] = {
            "revenue": rev_m,
            "gross_profit": _m(is_ws.cell(r_gp, col).value),
            "operating_income": oi_m,
            "operating_margin": f"{oi_m / rev_m * 100:.1f}%" if oi_m and rev_m else None,
            "net_income": _m(is_ws.cell(r_ni, col).value),
            "eps": round(eps_val, 2) if isinstance(eps_val, (int, float)) else None,
        }
    align_income_statement(income_base)

    balance_base = {}
    for label, key in [
        ("Cash & cash equivalents", "cash"),
        ("Inventories", "inventories"),
        ("TOTAL ASSETS", "total_assets"),
        ("TOTAL LIABILITIES", "total_liab"),
        ("TOTAL SHAREHOLDERS' EQUITY", "total_equity"),
    ]:
        row = bs_row(label)
        balance_base[key] = {fy: _m(bs.cell(row, PROJ_COL[fy]).value) for fy in PROJ_YEARS}

    cash_base = {}
    for label, key in [
        ("Net cash from operating activities", "cfo"),
        ("Capital expenditures", "capex"),
        ("Repurchase of common stock", "buybacks"),
        ("Depreciation & amortization", "dna"),
    ]:
        row = cf_row(label)
        cash_base[key] = {fy: _m(cf.cell(row, PROJ_COL[fy]).value) for fy in PROJ_YEARS}

    cash_base["fcf"] = {
        fy: (cash_base["cfo"][fy] or 0) + (cash_base["capex"][fy] or 0)
        for fy in PROJ_YEARS
    }
    align_cash_flow(cash_base)

    scen_bull = extract_paths(sc, COL_BULL)
    shares = {fy: is_ws.cell(r_sh, PROJ_COL[fy]).value for fy in PROJ_YEARS}
    income_bull = {
        fy: {k: scen_bull[fy][k] for k in (
            "revenue", "gross_profit", "operating_income", "operating_margin", "net_income"
        )}
        for fy in PROJ_YEARS
    }
    add_eps(income_bull, shares)
    align_income_statement(income_bull)

    cash_bull = {
        "cfo": {fy: scen_bull[fy]["cfo"] for fy in PROJ_YEARS},
        "capex": {fy: scen_bull[fy]["capex"] for fy in PROJ_YEARS},
        "fcf": {fy: scen_bull[fy]["fcf"] for fy in PROJ_YEARS},
        "buybacks": {fy: scen_bull[fy]["buybacks"] for fy in PROJ_YEARS},
        "dna": {fy: scen_bull[fy]["dna"] for fy in PROJ_YEARS},
    }
    align_cash_flow(cash_bull)
    balance_bull = bull_balance_sheet(balance_base, income_base, income_bull, cash_bull)

    return {
        "valuation": {
            "base_dcf": round(base, 0),
            "bear": round(bear, 0),
            "bull": round(bull, 0),
            "wacc": wacc["E41"].value,
            "coe": wacc["E32"].value,
            "beta": wacc["E25"].value,
            "rf": wacc["E3"].value,
            "erp": 0.06,
            "ev_m": dcf["E49"].value / 1000,
            "equity_m": dcf["E54"].value / 1000,
            "pv_fcf_m": dcf["E42"].value / 1000,
            "pv_tv_m": dcf["E48"].value / 1000,
            "cash_m": dcf["E50"].value / 1000,
            "shares_m": dcf["E55"].value / 1000,
            "prob_weighted": round(0.25 * bear + 0.5 * base + 0.25 * bull, 0),
        },
        "sensitivity": sens,
        "football_field": ff,
        "income_statement": {"base": income_base, "bull": income_bull},
        "balance_sheet": {"base": balance_base, "bull": balance_bull},
        "cash_flow": {"base": cash_base, "bull": cash_bull},
        "proj_years": list(PROJ_YEARS),
    }


if __name__ == "__main__":
    data = extract()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w") as f:
        json.dump(data, f, indent=2)
    v = data["valuation"]
    print(f"Wrote {OUT_JSON}")
    print(f"Base ${v['base_dcf']} | Bear ${v['bear']} | Bull ${v['bull']} | WACC {v['wacc']*100:.2f}%")
