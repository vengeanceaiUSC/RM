#!/usr/bin/env python3
"""Evaluate model18_wsp_formulas.xlsx without Excel (pure Python).

Mirrors Scenarios cols F/G/H DCF bridge and pitch-bridge EPS path.
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
import data as D

ROOT = SCRIPTS.parent
DEFAULT_PATH = ROOT / "model18_wsp_formulas.xlsx"

OTHER_INCOME_K = [45_000, 40_000, 35_000, 30_000, 25_000]
REPURCHASE_PRICES = [100, 108, 115, 122, 130]
COL_MAP = {"F": "bear", "G": "base", "H": "bull"}
PROJ_YEARS = D.PROJ_YEARS


def _m(v):
    return round(v / 1000) if v is not None else None


def _pct(num, den, d=1):
    return f"{num / den * 100:.{d}f}%" if den else None


def _wacc_from_sheet(wb) -> float:
    wacc = wb["WACC"]
    e10 = D.MKT["price"] * D.MKT["shares_out"]
    e14 = D.LEASE_FY25
    e3, e4, e5 = 0.048, 0.06, 0.3
    e17 = 0.86
    e20 = 2_140_000 / 11_140_000
    e22 = e17 / (1 + (1 - e5) * e20)
    e23 = e14 / e10
    e25 = e22 * (1 + (1 - e5) * e23)
    e32 = e3 + e25 * e4
    e36 = 0.05 * (1 - e5)
    e39 = e10 / (e10 + e14)
    e40 = e14 / (e10 + e14)
    return e39 * e32 + e40 * e36


def _col_inputs(scn, dcf, col: str) -> dict:
    def c(r):
        return scn[f"{col}{r}"].value

    e5 = dcf["E5"].value
    e8 = dcf["E8"].value
    e21, e23, e25 = dcf["E21"].value, dcf["E23"].value, dcf["E25"].value
    e27, e29 = dcf["E27"].value, dcf["E29"].value
    w = c(9)
    if isinstance(w, str):
        w = None
    return {
        "fy25_rev": scn["B198"].value,
        "fy25_nwc": dcf["E32"].value,
        "cash": dcf["E50"].value,
        "debt": dcf["E51"].value,
        "shares": dcf["E55"].value,
        "sh_start": scn["B197"].value,
        "fy25_ta": scn["B194"].value,
        "fy25_tl": scn["B195"].value,
        "fy25_te": scn["B196"].value,
        "bb_k": c(21) or 0,
        "g4": c(4),
        "g5": c(5),
        "g6": c(6),
        "g7": c(7),
        "g8": c(8),
        "g10": c(10),
        "g11": c(11),
        "g12": c(12),
        "g13": c(13),
        "g14": c(14),
        "g15": (e21 / e5) * 365,
        "g16": (e23 / e8) * 365,
        "g17": c(17),
        "g18": (e25 / e8) * 365,
        "g19": e27 / e5,
        "g20": e29 / e5,
        "wacc": w,
    }


def _eval_scenario(inp: dict, wacc_default: float) -> dict:
    w = inp["wacc"] if inp["wacc"] is not None else wacc_default
    revs = [inp["fy25_rev"] * (1 + inp["g4"])]
    for _ in range(4):
        revs.append(revs[-1] * (1 + inp["g5"]))

    ebit_margins = [inp["g6"] + (inp["g8"] - inp["g6"]) * i / 4 for i in range(5)]
    ebit = [revs[i] * ebit_margins[i] + (inp["g7"] if i == 0 else 0) for i in range(5)]
    nopat = [e * (1 - inp["g11"]) for e in ebit]
    da = [revs[i] * inp["g12"] for i in range(5)]
    capex = [revs[i] * inp["g13"] for i in range(5)]
    cogs = [revs[i] * (1 - inp["g14"]) for i in range(5)]

    nwcs = []
    for i in range(5):
        dio = inp["g16"] - inp["g17"] * (i + 1)
        ar = (inp["g15"] / 365) * revs[i]
        inv = (dio / 365) * cogs[i]
        ap = (inp["g18"] / 365) * cogs[i]
        nwcs.append(ar + inv + revs[i] * inp["g19"] - ap - revs[i] * inp["g20"])

    dnwc = [inp["fy25_nwc"] - nwcs[0]] + [nwcs[i - 1] - nwcs[i] for i in range(1, 5)]
    ufcf = [nopat[i] + da[i] - capex[i] + dnwc[i] for i in range(5)]

    pv = sum(ufcf[i] / (1 + w) ** (i + 1) for i in range(5))
    tv = ufcf[4] * (1 + inp["g10"]) / (w - inp["g10"]) / (1 + w) ** 5
    ev = pv + tv
    px = (ev + inp["cash"] + inp["debt"]) / inp["shares"]

    sh = inp["sh_start"]
    eps = []
    ni_list = []
    for i in range(5):
        ni = (ebit[i] + OTHER_INCOME_K[i]) * (1 - inp["g11"])
        ni_list.append(ni)
        sh -= inp["bb_k"] / REPURCHASE_PRICES[i]
        eps.append(ni / sh)

    shares_retired_m = (inp["sh_start"] - sh) / 1000

    cfo = [ni_list[i] + da[i] + dnwc[i] for i in range(5)]
    fcf_cfs = [cfo[i] - capex[i] for i in range(5)]
    buybacks = [inp["bb_k"]] * 5
    cash_bs = []
    c = inp["cash"]
    for i in range(5):
        c = c + fcf_cfs[i] - buybacks[i]
        cash_bs.append(c)

    fy25_ta = inp.get("fy25_ta")
    fy25_tl = inp.get("fy25_tl")
    fy25_te = inp.get("fy25_te")
    income = {}
    balance = {k: {} for k in ("cash", "inventories", "total_assets", "total_liab", "total_equity")}
    cash_flow = {k: {} for k in ("cfo", "capex", "fcf", "buybacks", "dna")}

    for i, fy in enumerate(PROJ_YEARS):
        dio = inp["g16"] - inp["g17"] * (i + 1)
        inv = (dio / 365) * cogs[i]
        income[fy] = {
            "revenue": _m(revs[i]),
            "gross_profit": _m(revs[i] * inp["g14"]),
            "operating_income": _m(ebit[i]),
            "operating_margin": _pct(ebit[i], revs[i]),
            "net_income": _m(ni_list[i]),
            "eps": round(eps[i], 2),
        }
        balance["cash"][fy] = _m(cash_bs[i])
        balance["inventories"][fy] = _m(inv)
        if fy25_ta:
            scale = revs[i] / inp["fy25_rev"]
            balance["total_assets"][fy] = _m(fy25_ta * scale)
            balance["total_liab"][fy] = _m(fy25_tl * scale)
            balance["total_equity"][fy] = _m(fy25_te * scale)
        cash_flow["cfo"][fy] = _m(cfo[i])
        cash_flow["capex"][fy] = _m(-capex[i])
        cash_flow["fcf"][fy] = _m(fcf_cfs[i])
        cash_flow["buybacks"][fy] = _m(-buybacks[i])
        cash_flow["dna"][fy] = _m(da[i])

    return {
        "implied_px": px,
        "ev": ev,
        "pv_explicit": pv,
        "pv_tv": tv,
        "ufcf": ufcf,
        "revs": revs,
        "ebit": ebit,
        "nopat": nopat,
        "ni": ni_list,
        "eps": eps,
        "wacc": w,
        "shares_retired_m": shares_retired_m,
        "upside_pct": (px / D.MKT["price"] - 1) * 100,
        "income": income,
        "balance": balance,
        "cash_flow": cash_flow,
    }


def evaluate(path: Path | str = DEFAULT_PATH) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)
    scn = wb["Scenarios"]
    dcf = wb["DCF"]
    wacc = _wacc_from_sheet(wb)
    out = {"wacc": wacc, "scenarios": {}}
    for col, name in COL_MAP.items():
        inp = _col_inputs(scn, dcf, col)
        out["scenarios"][name] = _eval_scenario(inp, wacc)
    return out


if __name__ == "__main__":
    r = evaluate()
    b = r["scenarios"]["base"]
    print(f"Base DCF ${b['implied_px']:.2f} (+{b['upside_pct']:.1f}%)")
    print(f"EPS path: {[round(x, 2) for x in b['eps']]}")
