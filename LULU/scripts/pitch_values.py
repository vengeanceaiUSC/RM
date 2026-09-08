"""Pitch-deck figures sourced from the DCF Scenarios tab (cols G base, H bull).

Regenerate after model changes:
  cd LULU/scripts && python3 build_dcf.py && python3 pitch_values.py
"""
import json
import os

import openpyxl

import data as D
from scenario_extract import extract_scenario, discover_model_refs, COL_BASE, COL_BULL, PROJ_YEARS

ROOT = os.path.join(os.path.dirname(__file__), "..")
DCF_PATH = os.path.join(ROOT, "LULU_DCF_Valuation_Model.xlsx")
OUT_JSON = os.path.join(ROOT, "data", "pitch_values.json")


def _recalc_dcf():
    import shutil
    from recalc_workbook import recalc_workbook

    tmp = "/tmp/lulu_models"
    os.makedirs(tmp, exist_ok=True)
    dest = os.path.join(tmp, os.path.basename(DCF_PATH))
    shutil.copy2(DCF_PATH, dest)
    recalc_workbook(dest)
    return openpyxl.load_workbook(dest, data_only=True)


def _wacc_cell(wacc, needle, exact=False):
    """Return (row, value in col E) for first label match on WACC tab."""
    for r in range(1, 60):
        lab = wacc.cell(r, 1).value
        if not lab:
            continue
        s = str(lab).strip().lower()
        if exact:
            if s == needle.lower():
                return r, wacc.cell(r, 5).value
        elif needle.lower() in s:
            return r, wacc.cell(r, 5).value
    raise KeyError(f"WACC row not found: {needle}")


def _extract_wacc_build(wacc):
    """Pull live cap-stack and WACC inputs from the DCF WACC tab."""
    r_mkt, mkt_eq = _wacc_cell(wacc, "market value of equity")
    r_lease, lease_d = _wacc_cell(wacc, "operating lease liabilities")
    r_fund, fund_d = _wacc_cell(wacc, "funded debt")
    r_debt, debt_tot = _wacc_cell(wacc, "total debt equivalents")
    r_bobs, beta_obs = _wacc_cell(wacc, "observed beta")
    r_bun, beta_unlev = _wacc_cell(wacc, "unlevered \u03b2")
    r_buse, beta_used = _wacc_cell(wacc, "beta used")
    r_deu, de_unlev = _wacc_cell(wacc, "d/e for unlever")
    r_der, de_relev = _wacc_cell(wacc, "d/e for relever")
    r_coe, coe = _wacc_cell(wacc, "cost of equity =")
    r_kd, kd = _wacc_cell(wacc, "pre-tax cost of debt")
    r_kdat, kdat = _wacc_cell(wacc, "after-tax cost of debt")
    r_we, we = _wacc_cell(wacc, "equity weight")
    r_wd, wd = _wacc_cell(wacc, "debt weight")
    r_wacc, wacc_val = _wacc_cell(wacc, "wacc", exact=True)
    r_rf, rf = _wacc_cell(wacc, "risk-free rate")
    r_erp, erp = _wacc_cell(wacc, "equity risk premium")
    r_tax, tax = _wacc_cell(wacc, "tax rate")
    cash_k = D.MKT["cash"]
    total_cap = mkt_eq + debt_tot
    return {
        "mkt_eq_m": round(mkt_eq / 1000),
        "lease_debt_m": round(lease_d / 1000),
        "funded_debt_m": round(fund_d / 1000),
        "total_debt_m": round(debt_tot / 1000),
        "cash_m": round(cash_k / 1000),
        "net_debt_m": round((debt_tot - cash_k) / 1000),
        "total_cap_m": round(total_cap / 1000),
        "equity_pct": round(we * 100, 1),
        "debt_pct": round(wd * 100, 1),
        "rf": rf,
        "erp": erp,
        "tax": tax,
        "beta_obs": beta_obs,
        "beta_unlev": beta_unlev,
        "beta": beta_used,
        "de_unlev": de_unlev,
        "de_relev": de_relev,
        "coe": coe,
        "kd": kd,
        "kd_at": kdat,
        "we": we,
        "wd": wd,
        "wacc": wacc_val,
        "rows": {
            "rf": r_rf, "erp": r_erp, "tax": r_tax,
            "mkt_eq": r_mkt, "lease_d": r_lease, "fund_d": r_fund, "debt_tot": r_debt,
            "beta_obs": r_bobs, "beta_unlev": r_bun, "beta": r_buse,
            "de_unlev": r_deu, "de_relev": r_der,
            "coe": r_coe, "kd": r_kd, "kd_at": r_kdat,
            "we": r_we, "wd": r_wd, "wacc": r_wacc,
        },
    }


def extract():
    dcf_wb = _recalc_dcf()
    dcf = dcf_wb["DCF"]
    sc = dcf_wb["Scenarios"]
    wacc = dcf_wb["WACC"]
    comps = dcf_wb["Comps"]

    pt_row = next(r for r in range(1, 500) if sc.cell(r, 1).value == "Implied share price")
    bear = sc.cell(pt_row, 6).value
    base = sc.cell(pt_row, 7).value
    bull = sc.cell(pt_row, 8).value

    income_base, balance_base, cash_base = extract_scenario(sc, COL_BASE)
    income_bull, balance_bull, cash_bull = extract_scenario(sc, COL_BULL)

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
        "model_refs": discover_model_refs(sc),
        "wacc_build": _extract_wacc_build(wacc),
        "source": "LULU_DCF_Valuation_Model.xlsx → Scenarios cols G (base) & H (bull)",
    }


if __name__ == "__main__":
    data = extract()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w") as f:
        json.dump(data, f, indent=2)
    v = data["valuation"]
    print(f"Wrote {OUT_JSON}")
    print(f"Base ${v['base_dcf']} | Bear ${v['bear']} | Bull ${v['bull']} | WACC {v['wacc']*100:.2f}%")
