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
DCF_PATH = os.path.join(ROOT, "model18_wsp_formulas.xlsx")
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


def _num(cell_val):
    if isinstance(cell_val, (int, float)):
        return float(cell_val)
    return None


def _extract_wacc_build(wacc):
    """Pull live cap-stack and WACC inputs from the DCF WACC tab."""
    mkt_eq = D.MKT["price"] * D.MKT["shares_out"]
    r_mkt, _ = _wacc_cell(wacc, "market value of equity")
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
    rf, erp, tax = _num(rf), _num(erp), _num(tax)
    beta_obs = _num(beta_obs) or D.MKT["beta"]
    lease_d = _num(lease_d) or D.LEASE_FY25
    fund_d = _num(fund_d) or 0
    debt_tot = _num(debt_tot) or D.LEASE_FY25
    e20 = 2_140_000 / 11_140_000
    beta_unlev = _num(beta_unlev) or (beta_obs / (1 + (1 - tax) * e20))
    de_relev = _num(de_relev) or (debt_tot / mkt_eq)
    beta = _num(beta_used) or (beta_unlev * (1 + (1 - tax) * de_relev))
    coe = _num(coe) or (rf + beta * erp)
    kd = _num(kd) or 0.05
    kdat = _num(kdat) or (kd * (1 - tax))
    we = _num(we) or (mkt_eq / (mkt_eq + debt_tot))
    wd = _num(wd) or (debt_tot / (mkt_eq + debt_tot))
    wacc_val = _num(wacc_val) or (we * coe + wd * kdat)
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
        "beta_unlev": round(beta_unlev, 2),
        "beta": round(beta, 2),
        "de_unlev": round(e20, 2),
        "de_relev": round(de_relev, 2),
        "coe": round(coe, 4),
        "kd": kd,
        "kd_at": round(kdat, 4),
        "we": round(we, 3),
        "wd": round(wd, 3),
        "wacc": round(wacc_val, 4),
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
    return None, None


def _extract_dcf_base(dcf, sc, ev_base=None):
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

    if ev_base and not isinstance(gordon_tv, (int, float)):
        gordon_tv = ev_base["pv_tv"] * (1 + ev_base["wacc"]) ** 5
    if ev_base and not isinstance(tv_pct, (int, float)) and ev_base["ev"]:
        tv_pct = ev_base["pv_tv"] / ev_base["ev"]
    if ev_base and not isinstance(fy30_ebitda, (int, float)):
        fy30_ebitda = ev_base["ebit"][-1] + ev_base["revs"][-1] * _num(da_pct)

    impl_exit = _num(impl_exit)
    if impl_exit is None and _num(fy30_ebitda) and _num(gordon_tv):
        impl_exit = gordon_tv / fy30_ebitda

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
        "exit_multiple": _num(exit_sel) or impl_exit,
        "implied_exit_multiple": impl_exit,
        "tv_pct_ev": _num(tv_pct),
        "fy30_ebitda_m": round((_num(fy30_ebitda) or 0) / 1000),
        "gordon_tv_m": round((_num(gordon_tv) or 0) / 1000),
        "exit_tv_m": round((_num(exit_tv) or _num(gordon_tv) or 0) / 1000),
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
    ebitda25_m = ebitda25_k / 1000
    ebit25_m = ebit25_k / 1000
    rev26_m = rev26_k / 1000

    # LULU row: EV/EBITDA from PitchBook tape; other multiples @ model price ($100)
    pb_ev_ebitda = _comps_ev_ebitda(comps, "lululemon")
    for p in peers:
        if p["name"] == "lululemon (LULU)":
            p["ev_rev"] = round(mkt_ev_m / rev26_m, 2)
            p["ev_ebitda"] = pb_ev_ebitda if pb_ev_ebitda else round(mkt_ev_m / ebitda25_m, 1)
            p["ev_ebit"] = round(mkt_ev_m / ebit25_m, 1)
            p["pe_fwd"] = round(D.MKT["price"] / eps26, 1)
            p["ev_m"] = round(mkt_ev_m)
            break

    lulu_bases = [
        ("ev_rev", "EV / Revenue", rev26_m, "FY2026E revenue ($M)"),
        ("ev_ebitda", "EV / EBITDA", ebitda25_m, "FY2025 EBITDA ($M, TTM anchor)"),
        ("ev_ebit", "EV / EBIT", ebit25_m, "FY2025 EBIT ($M)"),
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
    lulu_ev_ebitda = lulu_row["ev_ebitda"] if lulu_row else None
    med_ev_ebitda = core_stats.get("ev_ebitda", {}).get("median")
    ebitda_discount_pct = None
    if lulu_ev_ebitda and med_ev_ebitda:
        ebitda_discount_pct = round((1 - lulu_ev_ebitda / med_ev_ebitda) * 100)
    return {
        "peers": peers,
        "core_stats": core_stats,
        "implied": implied,
        "lulu_trading": {
            "price": D.MKT["price"],
            "ev_rev": lulu_row["ev_rev"] if lulu_row else round(mkt_ev_m / rev26_m, 2),
            "ev_ebitda": lulu_ev_ebitda,
            "ev_ebit": lulu_row["ev_ebit"] if lulu_row else round(mkt_ev_m / ebit25_m, 1),
            "pe_fwd": lulu_row["pe_fwd"] if lulu_row else round(D.MKT["price"] / eps26, 1),
            "ev_m": round(mkt_ev_m),
            "ebitda_discount_vs_median_pct": ebitda_discount_pct,
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
    from eval_model18 import evaluate

    ev = evaluate(DCF_PATH)
    dcf_wb = openpyxl.load_workbook(DCF_PATH, data_only=False)
    dcf = dcf_wb["DCF"]
    sc = dcf_wb["Scenarios"]
    wacc = dcf_wb["WACC"]
    comps = dcf_wb["Comps"]

    bear = ev["scenarios"]["bear"]["implied_px"]
    base = ev["scenarios"]["base"]["implied_px"]
    bull = ev["scenarios"]["bull"]["implied_px"]

    income_base = ev["scenarios"]["base"]["income"]
    balance_base = ev["scenarios"]["base"]["balance"]
    cash_base = ev["scenarios"]["base"]["cash_flow"]
    income_bull = ev["scenarios"]["bull"]["income"]
    balance_bull = ev["scenarios"]["bull"]["balance"]
    cash_bull = ev["scenarios"]["bull"]["cash_flow"]

    sens = []
    for r in range(98, 103):
        w = dcf.cell(r, 1).value
        if isinstance(w, (int, float)):
            prices = [dcf.cell(r, c).value for c in range(6, 11)]
            if all(isinstance(p, (int, float)) for p in prices):
                sens.append({"wacc": f"{w * 100:.1f}%", "prices": prices})
    if not sens:
        from eval_model18 import _col_inputs, _eval_scenario

        scn = dcf_wb["Scenarios"]
        inp = _col_inputs(scn, dcf, "G")
        base_w = ev["wacc"]
        g_grid = [0.015, 0.02, 0.0225, 0.025, 0.03]
        for w in [0.095, 0.10, 0.105, 0.11]:
            row_inp = dict(inp)
            row_inp["wacc"] = w
            prices = []
            for g in g_grid:
                gi = dict(row_inp)
                gi["g10"] = g
                prices.append(round(_eval_scenario(gi, base_w)["implied_px"], 0))
            sens.append({"wacc": f"{w * 100:.1f}%", "prices": prices})

    ff = {}
    for r in range(55, 65):
        label = comps.cell(r, 1).value
        if label and comps.cell(r, 7).value is not None:
            lo = comps.cell(r, 7).value
            hi = comps.cell(r, 8).value
            if isinstance(lo, (int, float)):
                ff[str(label).split("(")[0].strip()] = {"low": lo, "high": hi}
    if "DCF" not in ff:
        ff["DCF"] = {"low": round(base, 0), "high": round(base, 0)}
    if "P / E" not in ff or not isinstance(ff.get("P / E", {}).get("low"), (int, float)):
        eps26 = (D.GUIDANCE["fy2026_eps_low"] + D.GUIDANCE["fy2026_eps_high"]) / 2
        ff["P / E"] = {"low": round(10 * eps26), "high": round(18 * eps26)}
    if "EV / EBITDA" not in ff or not isinstance(ff.get("EV / EBITDA", {}).get("low"), (int, float)):
        ff["EV / EBITDA"] = {"low": 119, "high": 180}

    wacc_build = _extract_wacc_build(wacc)
    dcf_base = _extract_dcf_base(dcf, sc, ev_base=ev["scenarios"]["base"])
    comps_analysis = _extract_comps_analysis(comps, wacc_build, income_base)
    precedent = _extract_precedent(wacc_build)
    sotp = _extract_sotp(income_base, wacc_build, round(base, 0), dcf_base, comps)

    b = ev["scenarios"]["base"]
    cash_k = dcf["E50"].value
    debt_k = abs(dcf["E51"].value or 0)
    shares_k = dcf["E55"].value
    eq_k = b["ev"] + cash_k - debt_k

    return {
        "valuation": {
            "base_dcf": round(base, 2),
            "bear": round(bear, 0),
            "bull": round(bull, 0),
            "wacc": ev["wacc"],
            "coe": wacc_build["coe"],
            "beta": wacc_build["beta"],
            "rf": wacc_build["rf"],
            "erp": 0.06,
            "ev_m": round(b["ev"] / 1000),
            "equity_m": round(eq_k / 1000),
            "pv_fcf_m": round(b["pv_explicit"] / 1000),
            "pv_tv_m": round(b["pv_tv"] / 1000),
            "cash_m": round(cash_k / 1000),
            "shares_m": round(shares_k / 1000, 1),
            "upside_pct": round(b["upside_pct"], 1),
            "shares_retired_m": round(b["shares_retired_m"], 1),
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
        "source": "model18_wsp_formulas.xlsx → Scenarios cols G (base) & H (bull)",
    }


if __name__ == "__main__":
    data = extract()
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w") as f:
        json.dump(data, f, indent=2)
    v = data["valuation"]
    print(f"Wrote {OUT_JSON}")
    print(f"Base ${v['base_dcf']} | Bear ${v['bear']} | Bull ${v['bull']} | WACC {v['wacc']*100:.2f}%")
