#!/usr/bin/env python3
"""Recalculate model18_wsp_formulas.xlsx via LibreOffice and read cached values."""
from __future__ import annotations

import subprocess
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DEFAULT = ROOT / "model18_wsp_formulas.xlsx"


def recalc_workbook(path: Path = DEFAULT) -> openpyxl.Workbook:
    """Return data_only workbook after LibreOffice recalc."""
    out_dir = Path("/tmp/model18_recalc")
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / path.name
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", str(out_dir), str(path)],
        check=False,
        timeout=120,
    )
    if not out.exists():
        raise RuntimeError(f"LibreOffice recalc failed: {out} not found")
    return openpyxl.load_workbook(out, data_only=True)


def read_dcf_outputs(wb: openpyxl.Workbook) -> dict:
    scn = wb["Scenarios"]
    dcf = wb["DCF"]
    out = {"scenarios": {}, "base": {}}
    for col, name in [("B", "bear"), ("C", "base"), ("D", "bull")]:
        ev = scn[f"{col}131"].value
        px = scn[f"{col}132"].value
        up = scn[f"{col}133"].value
        out["scenarios"][name] = {
            "implied_px": float(px),
            "upside_pct": float(up) * 100,
            "ev": float(ev),
        }
    cash = float(dcf["E50"].value or 0)
    debt = float(dcf["E51"].value or 0)
    shares = float(dcf["E55"].value or 0)
    base = out["scenarios"]["base"]
    out["bridge"] = {
        "pv_explicit": base["ev"] - _pv_tv_from_wb(wb),  # fallback below
        "cash_m": round(cash / 1000),
        "equity_m": round((base["ev"] + cash + debt) / 1000),
        "shares_m": round(shares / 1000, 1),
    }
    # read PV from DCF text block isn't available; compute from scenarios
    b = out["scenarios"]["base"]
    out["bridge"]["ev_m"] = round(b["ev"] / 1000)
    out["bridge"]["equity_m"] = round((b["ev"] + cash + debt) / 1000)
    out["bridge"]["pv_fcf_m"] = round(_sum_pv_explicit(scn, "C"))
    out["bridge"]["pv_tv_m"] = round(b["ev"] / 1000) - out["bridge"]["pv_fcf_m"]

    # base-case pitch bridge (col C)
    def row5(start: int) -> list[float]:
        return [float(scn.cell(start + i, 3).value or 0) for i in range(5)]

    out["base"]["revenue_m"] = [round(v / 1000) for v in row5(25)]
    out["base"]["ni_m"] = [round(v / 1000) for v in row5(143)]
    # Prefer DCF repurchase schedule (deck cites row 78); fall back to Scenarios bridge.
    dcf = wb["DCF"]
    eps_cols = ["F", "G", "H", "I", "J"]
    dcf_eps = [
        round(float(dcf[f"{col}78"].value or 0), 2) for col in eps_cols
    ]
    scn_eps = [
        round(float(scn.cell(188 + i, 3).value or 0), 2) for i in range(5)
    ]
    out["base"]["eps"] = dcf_eps if any(dcf_eps) else scn_eps
    out["base"]["cfo_m"] = [round(v / 1000) for v in row5(148)]
    out["base"]["fcf_m"] = [round(v / 1000) for v in row5(153)]
    out["base"]["ebit_m"] = [round(v / 1000) for v in row5(45)]
    out["base"]["gp_m"] = [round(float(scn.cell(25 + i, 3).value or 0) * float(scn.cell(30 + i, 3).value or 0) / 1000) for i in range(5)]
    out["base"]["cash_m"] = [round(v / 1000) for v in row5(163)]
    out["base"]["inv_m"] = [round(scn.cell(80 + i, 3).value / 1000) for i in range(5)]
    out["base"]["ta_m"] = [round(v / 1000) for v in row5(168)]
    out["base"]["tl_m"] = [round(v / 1000) for v in row5(173)]
    out["base"]["te_m"] = [round(v / 1000) for v in row5(178)]
    out["base"]["dna_m"] = [round(v / 1000) for v in row5(55)]
    out["base"]["capex_m"] = [round(abs(scn.cell(60 + i, 3).value or 0) / 1000) for i in range(5)]
    out["base"]["buy_m"] = [round(abs(scn.cell(158 + i, 3).value or 0) / 1000) for i in range(5)]
    out["sensitivity"] = read_sensitivity_grid(dcf)
    return out


def read_sensitivity_grid(dcf) -> list[dict]:
    """DCF rows 99–103: WACC × terminal g implied share price."""
    sens: list[dict] = []
    for r in range(99, 104):
        w = dcf.cell(r, 1).value
        if not isinstance(w, (int, float)):
            continue
        prices = [dcf.cell(r, c).value for c in range(6, 11)]
        if not all(isinstance(p, (int, float)) for p in prices):
            continue
        sens.append({"wacc": f"{w * 100:.1f}%", "prices": [round(p) for p in prices]})
    return sens


def _pv_tv_from_wb(wb):
    return 0


def _sum_pv_explicit(scn, col: str) -> float:
    """Approximate PV explicit from UFCF rows 125-129 — use EV - TV from sheet."""
    import re
    # NPV not cached separately; use ufcf * discount rough - skip
    w = float(scn[f"{col}9"].value or 0.09)
    ufcf = [float(scn.cell(125 + i, ord(col) - ord("A") + 1).value or 0) for i in range(5)]
    return sum(v / (1 + w) ** (i + 1) for i, v in enumerate(ufcf)) / 1000


if __name__ == "__main__":
    data = read_dcf_outputs(recalc_workbook())
    import json
    print(json.dumps(data, indent=2))
