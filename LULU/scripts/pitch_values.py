"""Pitch-deck figures sourced from the DCF Scenarios tab (cols G base, H bull).

Regenerate after model changes:
  cd LULU/scripts && python3 build_dcf.py && python3 pitch_values.py
"""
import json
import os

import openpyxl

import data as D
from driver_kpis import CANONICAL
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


def _scen_cell(sc, needle, col=COL_BASE):
    for r in range(1, 400):
        lab = sc.cell(r, 1).value
        if not lab:
            continue
        if needle.lower() in str(lab).strip().lower():
            return r, sc.cell(r, col).value
    raise KeyError(f"Scenarios row not found: {needle}")


def _dcf_cell(dcf, needle):
    for r in range(1, 120):
        lab = dcf.cell(r, 1).value
        if not lab:
            continue
        if needle.lower() in str(lab).strip().lower():
            return r, dcf.cell(r, 5).value
    raise KeyError(f"DCF row not found: {needle}")


def _extract_dcf_base(dcf, sc):
    _, rev_g1 = _scen_cell(sc, "fy2026e revenue growth")
    _, rev_gterm = _scen_cell(sc, "fy2027")
    _, m1 = _scen_cell(sc, "run-rate ebit margin")
    _, mterm = _scen_cell(sc, "terminal (fy2030e) ebit")
    _, tg = _scen_cell(sc, "terminal growth")
    _, tax = _scen_cell(sc, "cash tax rate")
    _, capex_pct = _scen_cell(sc, "capex % of revenue")
    _, da_pct = _scen_cell(sc, "d&a % of revenue")
    _, tariff = _scen_cell(sc, "ieepa tariff refunds")
    _, impl_exit = _dcf_cell(dcf, "implied exit ev/ebitda")
    _, exit_sel = _dcf_cell(dcf, "selected exit ev/ebitda")
    _, tv_pct = _dcf_cell(dcf, "% of ev from terminal value")
    _, fy30_ebitda = _dcf_cell(dcf, "terminal ebitda (fy2030e")
    _, gordon_tv = _dcf_cell(dcf, "terminal value = fcf")
    _, exit_tv = _dcf_cell(dcf, "terminal value = terminal ebitda")
    return {
        "rev_growth_fy26": rev_g1,
        "rev_growth_fy27_30": rev_gterm,
        "ebit_margin_clean": m1,
        "ebit_margin_terminal": mterm,
        "tariff_refund_k": tariff,
        "terminal_g": tg,
        "tax": tax,
        "capex_pct": capex_pct,
        "da_pct": da_pct,
        "exit_multiple": exit_sel,
        "implied_exit_multiple": impl_exit,
        "tv_pct_ev": tv_pct,
        "fy30_ebitda_m": round(fy30_ebitda / 1000),
        "gordon_tv_m": round(gordon_tv / 1000),
        "exit_tv_m": round(exit_tv / 1000),
    }


def _comps_ev_ebitda(comps, needle):
    """Return PitchBook EV/EBITDA from Comps peer table (col E), or None."""
    for r in range(1, 120):
        lab = comps.cell(r, 1).value
        if not lab or needle.lower() not in str(lab).lower():
            continue
        val = comps.cell(r, 5).value
        if isinstance(val, (int, float)):
            return round(val, 1)
    return None


def _median(vals):
    s = sorted(vals)
    n = len(s)
    if not n:
        return None
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


def _extract_comps_analysis(comps, wacc_build, income_base):
    """Multivariate comps: peer multiples and implied LULU equity value ($/sh)."""
    peers = []
    for r in range(1, 150):
        name = comps.cell(r, 1).value
        if not name or name not in D.PEER_FINANCIALS:
            continue
        ev = comps.cell(r, 9).value
        ebitda_pb = comps.cell(r, 10).value
        if not isinstance(ev, (int, float)):
            continue
        fin = D.PEER_FINANCIALS[name]
        rev = fin["revenue"]
        ebit = fin["ebit"]
        ebitda = ebitda_pb if isinstance(ebitda_pb, (int, float)) and ebitda_pb > 0 else fin["ebitda"]
        pe_fwd = fin.get("pe_fwd")
        peer = {
            "name": name,
            "core": fin["core"],
            "ev_m": round(ev / 1000),
            "ev_rev": round(ev / rev, 2) if rev > 0 else None,
            "ev_ebitda": round(ev / ebitda, 1) if ebitda and ebitda > 0 else None,
            "ev_ebit": round(ev / ebit, 1) if ebit and ebit > 0 else None,
            "pe_fwd": pe_fwd,
        }
        peers.append(peer)

    rev26_k = (D.GUIDANCE["fy2026_rev_low"] + D.GUIDANCE["fy2026_rev_high"]) / 2
    eps26 = (D.GUIDANCE["fy2026_eps_low"] + D.GUIDANCE["fy2026_eps_high"]) / 2
    ebitda25_k = D.IS["operating_income"]["FY2025"] + D.CF["d_and_a"]["FY2025"]
    ebit25_k = D.IS["operating_income"]["FY2025"]
    cash_m = wacc_build["cash_m"]
    debt_m = wacc_build["total_debt_m"]
    shares_m = D.MKT["shares_out"] / 1000
    mkt_ev_m = (D.MKT["price"] * D.MKT["shares_out"]) / 1000 - cash_m + debt_m

    lulu_bases = [
        ("ev_rev", "EV / Revenue", rev26_k / 1000, "FY2026E revenue ($M)"),
        ("ev_ebitda", "EV / EBITDA", ebitda25_k / 1000, "FY2025 EBITDA ($M, TTM anchor)"),
        ("ev_ebit", "EV / EBIT", ebit25_k / 1000, "FY2025 EBIT ($M)"),
        ("pe_fwd", "P / E", eps26, "FY2026E EPS ($)"),
    ]

    core = [p for p in peers if p["core"] and p["name"] != "lululemon (LULU)"]
    implied = []
    for key, label, denom, denom_label in lulu_bases:
        mults = [p[key] for p in core if p.get(key) is not None]
        if not mults:
            continue
        med = _median(mults)
        lo_m, hi_m = min(mults), max(mults)
        if key == "pe_fwd":
            px_med = med * denom
            px_lo = lo_m * denom
            px_hi = hi_m * denom
            implied.append({
                "metric": label,
                "lulu_base": round(denom, 2),
                "lulu_base_label": denom_label,
                "peer_median": round(med, 1),
                "peer_low": round(lo_m, 1),
                "peer_high": round(hi_m, 1),
                "implied_ev_m": None,
                "implied_px_median": round(px_med),
                "implied_px_low": round(px_lo),
                "implied_px_high": round(px_hi),
            })
            continue
        ev_med = med * denom
        ev_lo = lo_m * denom
        ev_hi = hi_m * denom
        eq_med = ev_med + cash_m - debt_m
        eq_lo = ev_lo + cash_m - debt_m
        eq_hi = ev_hi + cash_m - debt_m
        implied.append({
            "metric": label,
            "lulu_base": round(denom, 2) if key == "ev_rev" else round(denom),
            "lulu_base_label": denom_label,
            "peer_median": round(med, 1) if key != "ev_rev" else round(med, 2),
            "peer_low": round(lo_m, 1) if key != "ev_rev" else round(lo_m, 2),
            "peer_high": round(hi_m, 1) if key != "ev_rev" else round(hi_m, 2),
            "implied_ev_m": round(ev_med),
            "implied_px_median": round(eq_med / shares_m),
            "implied_px_low": round(eq_lo / shares_m),
            "implied_px_high": round(eq_hi / shares_m),
        })

    core_stats = {}
    for key in ("ev_rev", "ev_ebitda", "ev_ebit", "pe_fwd"):
        vals = [p[key] for p in core if p.get(key) is not None]
        if vals:
            core_stats[key] = {
                "median": round(_median(vals), 2 if key == "ev_rev" else 1),
                "low": round(min(vals), 2 if key == "ev_rev" else 1),
                "high": round(max(vals), 2 if key == "ev_rev" else 1),
            }

    lulu_row = next((p for p in peers if p["name"] == "lululemon (LULU)"), None)
    return {
        "peers": peers,
        "core_stats": core_stats,
        "implied": implied,
        "lulu_trading": {
            "price": D.MKT["price"],
            "ev_rev": round(mkt_ev_m / (rev26_k / 1000), 2),
            "ev_ebitda": lulu_row["ev_ebitda"] if lulu_row else None,
            "ev_ebit": round(mkt_ev_m / (ebit25_k / 1000), 1),
            "pe_fwd": round(D.MKT["price"] / eps26, 1),
            "ev_m": round(mkt_ev_m),
        },
        "source": "PitchBook Comps Set 04-Sep-2026 (EV, TTM EBITDA); revenue/EBIT from company filings; forward P/E from consensus",
    }


def _extract_precedent(wacc_build):
    """Private-market precedent references — context only, not control-premium math."""
    rev26_m = (D.GUIDANCE["fy2026_rev_low"] + D.GUIDANCE["fy2026_rev_high"]) / 2 / 1000
    cash_m = wacc_build["cash_m"]
    debt_m = wacc_build["total_debt_m"]
    mkt_ev_m = (D.MKT["price"] * D.MKT["shares_out"]) / 1000 - cash_m + debt_m
    lulu_ev_sales = round(mkt_ev_m / rev26_m, 2)
    return {
        "deals": [dict(d) for d in D.PRECEDENT_TRANSACTIONS],
        "lulu_ev_sales": lulu_ev_sales,
        "lulu_price": D.MKT["price"],
        "lulu_ev_m": round(mkt_ev_m),
        "note": (
            "Private precedents inform what strategics/PE may pay for premium athleisure assets. "
            "We do not derive a control premium from an unclosed ask vs LULU's public trading multiple. "
            "Valuation anchors remain DCF and public comps (prior slide)."
        ),
    }


def _extract_sotp(income_base, wacc_build, base_dcf, dcf_base, comps):
    """Geographic SOTP on FY30E base-case revenue mix (10-K segment geography).

    Segment EV/EBITDA spreads anchor to the Gordon-implied exit multiple from the
    base-case DCF (TV / FY30 EBITDA identity) — the same ~7.4x selected in the model.
    The 5.0–8.0x football-field band on the Comps tab is a separate triangulation check.
    """
    fy25_m = D.IS["revenue"]["FY2025"] / 1000
    fy30_m = income_base["FY2030E"]["revenue"]
    rev_scale = fy30_m / fy25_m
    geo = {
        "Americas": CANONICAL["revenue_geo_americas"]["FY2025"] / 1000,
        "China Mainland": CANONICAL["revenue_geo_china"]["FY2025"] / 1000,
        "Rest of World": CANONICAL["revenue_geo_row"]["FY2025"] / 1000,
    }
    gordon_exit = dcf_base["implied_exit_multiple"]
    gordon_label = f"{gordon_exit:.1f}x"
    fy30_ebitda_m = dcf_base["fy30_ebitda_m"]
    lulu_ttm = _comps_ev_ebitda(comps, "lululemon")
    nke_ttm = _comps_ev_ebitda(comps, "nike")
    ads_ttm = _comps_ev_ebitda(comps, "adidas")
    peer_bits = []
    if lulu_ttm:
        peer_bits.append(f"LULU ~{lulu_ttm:.1f}x TTM")
    if nke_ttm and ads_ttm:
        peer_bits.append(f"NKE ~{nke_ttm:.1f}x / adidas ~{ads_ttm:.1f}x TTM (PitchBook)")
    peer_ref = "; ".join(peer_bits) if peer_bits else "PitchBook pubcomps (Comps tab)"

    # Spreads vs Gordon-implied exit (selected TV identity on consolidated FY30E EBITDA).
    seg_cfg = [
        (
            "Americas", -1.0, -0.25, 0.165,
            f"Mature discount vs Gordon {gordon_label}; near {peer_ref.split(';')[0] if peer_bits else 'LULU trough'}",
        ),
        (
            "China Mainland", 0.5, 2.0, 0.195,
            f"Growth premium vs Americas; still below NKE/adidas TTM ({peer_ref.split(';')[-1].strip() if len(peer_bits) > 1 else 'PitchBook comps'})",
        ),
        (
            "Rest of World", -0.25, 0.75, 0.180,
            f"Near Gordon anchor ({gordon_label}); expansion markets between Americas and China",
        ),
    ]
    raw_ebitda = {}
    for name, _, _, ebitda_m, _ in seg_cfg:
        raw_ebitda[name] = geo[name] * rev_scale * ebitda_m
    ebitda_scale = fy30_ebitda_m / sum(raw_ebitda.values())

    segments = []
    ev_lo = ev_hi = 0.0
    for name, lo_spread, hi_spread, ebitda_m, mult_note in seg_cfg:
        rev30 = geo[name] * rev_scale
        ebitda = raw_ebitda[name] * ebitda_scale
        mlo = round(gordon_exit + lo_spread, 1)
        mhi = round(gordon_exit + hi_spread, 1)
        seg_ev_lo = ebitda * mlo
        seg_ev_hi = ebitda * mhi
        ev_lo += seg_ev_lo
        ev_hi += seg_ev_hi
        segments.append({
            "segment": name,
            "fy30_rev_m": round(rev30),
            "ebitda_margin_pct": round(ebitda_m * 100, 1),
            "fy30_ebitda_m": round(ebitda),
            "ev_ebitda_lo": mlo,
            "ev_ebitda_hi": mhi,
            "ev_lo_m": round(seg_ev_lo),
            "ev_hi_m": round(seg_ev_hi),
            "multiple_note": mult_note,
        })
    wavg_lo = ev_lo / fy30_ebitda_m
    wavg_hi = ev_hi / fy30_ebitda_m
    cash = wacc_build["cash_m"]
    debt = wacc_build["total_debt_m"]
    shares_m = D.MKT["shares_out"] / 1000
    eq_lo = ev_lo + cash - debt
    eq_hi = ev_hi + cash - debt
    return {
        "segments": segments,
        "total_ev_lo_m": round(ev_lo),
        "total_ev_hi_m": round(ev_hi),
        "cash_m": cash,
        "debt_m": debt,
        "equity_lo_m": round(eq_lo),
        "equity_hi_m": round(eq_hi),
        "implied_px_lo": round(eq_lo / shares_m),
        "implied_px_hi": round(eq_hi / shares_m),
        "consolidated_dcf_px": round(base_dcf),
        "fy30_rev_m": round(fy30_m),
        "fy30_ebitda_m": fy30_ebitda_m,
        "gordon_exit_multiple": round(gordon_exit, 2),
        "consolidated_multiple_lo": round(wavg_lo, 1),
        "consolidated_multiple_hi": round(wavg_hi, 1),
        "ff_ev_ebitda_band": "5.0–8.0x",
        "multiple_source": (
            f"Segment spreads vs Gordon-implied exit {gordon_label} "
            f"(FCF₅×(1+g)/(WACC−g) ÷ FY30 EBITDA — selected TV in DCF). "
            f"5.0–8.0x Comps football-field band is a separate check, not the SOTP anchor. "
            f"Peer tape: {peer_ref}."
        ),
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

    wacc_build = _extract_wacc_build(wacc)
    dcf_base = _extract_dcf_base(dcf, sc)
    comps_analysis = _extract_comps_analysis(comps, wacc_build, income_base)
    precedent = _extract_precedent(wacc_build)
    sotp = _extract_sotp(income_base, wacc_build, round(base, 0), dcf_base, comps)

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
        "wacc_build": wacc_build,
        "dcf_base": dcf_base,
        "comps_analysis": comps_analysis,
        "precedent": precedent,
        "sotp": sotp,
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
