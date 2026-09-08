"""Pitch-deck figures sourced from the live DCF + 3-statement models.

Regenerate after model changes:
  cd LULU/scripts && python3 pitch_values.py
"""
import json
import os
import subprocess

import openpyxl

ROOT = os.path.join(os.path.dirname(__file__), "..")
DCF_PATH = os.path.join(ROOT, "LULU_DCF_Valuation_Model.xlsx")
IS_PATH = os.path.join(ROOT, "LULU_3_Statement_Model.xlsx")
OUT_JSON = os.path.join(ROOT, "data", "pitch_values.json")


def _recalc(path):
    subprocess.run(
        ["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", "/tmp", path],
        check=True,
        capture_output=True,
    )
    return openpyxl.load_workbook(f"/tmp/{os.path.basename(path)}", data_only=True)


def _m(v):
    if v is None:
        return None
    if isinstance(v, str):
        return v
    return round(v / 1000)


def _pct(v, d=1):
    return f"{v * 100:.{d}f}%" if v is not None else None


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
            if is_ws.cell(r, 1).value == label:
                return r
        return None

    r_rev = is_row("Net revenue")
    r_gp = is_row("Gross profit")
    r_oi = is_row("Operating income (EBIT)")
    r_om = is_row("Operating margin % (reported, incl. FY26 refund)")
    r_ni = is_row("Net income")
    r_eps = is_row("Diluted EPS ($)")

    col = {"FY2026E": 9, "FY2028E": 11, "FY2030E": 13}

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

    data = {
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
        "income_statement": {
            "FY2026E": {
                "revenue": _m(is_ws.cell(r_rev, col["FY2026E"]).value),
                "gross_profit": _m(is_ws.cell(r_gp, col["FY2026E"]).value),
                "operating_income": _m(is_ws.cell(r_oi, col["FY2026E"]).value),
                "operating_margin": _pct(is_ws.cell(r_om, col["FY2026E"]).value),
                "net_income": _m(is_ws.cell(r_ni, col["FY2026E"]).value),
                "eps": round(is_ws.cell(r_eps, col["FY2026E"]).value, 2),
            },
            "FY2028E": {
                "revenue": _m(is_ws.cell(r_rev, col["FY2028E"]).value),
                "gross_profit": _m(is_ws.cell(r_gp, col["FY2028E"]).value),
                "operating_income": _m(is_ws.cell(r_oi, col["FY2028E"]).value),
                "operating_margin": _pct(is_ws.cell(r_om, col["FY2028E"]).value),
                "net_income": _m(is_ws.cell(r_ni, col["FY2028E"]).value),
                "eps": round(is_ws.cell(r_eps, col["FY2028E"]).value, 2),
            },
            "FY2030E": {
                "revenue": _m(is_ws.cell(r_rev, col["FY2030E"]).value),
                "gross_profit": _m(is_ws.cell(r_gp, col["FY2030E"]).value),
                "operating_income": _m(is_ws.cell(r_oi, col["FY2030E"]).value),
                "operating_margin": _pct(is_ws.cell(r_om, col["FY2030E"]).value),
                "net_income": _m(is_ws.cell(r_ni, col["FY2030E"]).value),
                "eps": round(is_ws.cell(r_eps, col["FY2030E"]).value, 2),
            },
        },
        "balance_sheet": {},
        "cash_flow": {},
    }

    for label, key in [
        ("Cash & cash equivalents", "cash"),
        ("Inventories", "inventories"),
        ("TOTAL ASSETS", "total_assets"),
        ("TOTAL LIABILITIES", "total_liab"),
        ("TOTAL SHAREHOLDERS' EQUITY", "total_equity"),
    ]:
        for r in range(1, 60):
            if bs.cell(r, 1).value == label:
                data["balance_sheet"][key] = {
                    "FY2026E": _m(bs.cell(r, 9).value),
                    "FY2028E": _m(bs.cell(r, 11).value),
                    "FY2030E": _m(bs.cell(r, 13).value),
                }
                break

    for label, key in [
        ("Net cash from operating activities", "cfo"),
        ("Capital expenditures", "capex"),
        ("Repurchase of common stock", "buybacks"),
        ("Depreciation & amortization", "dna"),
    ]:
        for r in range(1, 60):
            if cf.cell(r, 1).value == label:
                data["cash_flow"][key] = {
                    "FY2026E": _m(cf.cell(r, 9).value),
                    "FY2028E": _m(cf.cell(r, 11).value),
                    "FY2030E": _m(cf.cell(r, 13).value),
                }
                break

    # FCF = CFO + capex (capex negative)
    data["cash_flow"]["fcf"] = {
        fy: data["cash_flow"]["cfo"][fy] + data["cash_flow"]["capex"][fy]
        for fy in ("FY2026E", "FY2028E", "FY2030E")
    }
    return data


if __name__ == "__main__":
    data = extract()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w") as f:
        json.dump(data, f, indent=2)
    v = data["valuation"]
    print(f"Wrote {OUT_JSON}")
    print(f"Base ${v['base_dcf']} | Bear ${v['bear']} | Bull ${v['bull']} | WACC {v['wacc']*100:.2f}%")
