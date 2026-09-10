#!/usr/bin/env python3
"""Apply human-analyst formula and structure fixes (5 checklist items).

- Round Scenarios WC hardcodes to 1-2 decimals
- Draggable EBIT margin ramp on DCF
- Expose FY25 D&A on DCF; Comps links to cell not literal
- Revenue growth driver row above Net Revenue with drag-right rev formulas
- Rename WACC Yahoo line items; sources stay in comments
- Delete Scratch tab

Run:  cd LULU && python3 scripts/fix_human_analyst_structure.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

SCN = "Scenarios"
FCOLS = ["C", "D", "E", "F", "G"]
FY25 = "B"
FY25_DA = 496228
STEP = f"({SCN}!$D$8-{SCN}!$D$6)/4"

# Scenarios row → rounded values (base/bear/bull cols C-E)
SCN_ROUNDS: dict[int, float] = {
    15: 6.3,      # DSO
    16: 128.8,    # DIO
    18: 25.1,     # DPO
    19: 0.051,    # Prepaid 5.1%
    20: 0.060,    # Accrued 6.0%
}

WACC_RENAMES: dict[int, tuple[str, str | None]] = {
    17: ("Total Debt", "Yahoo Finance: LULU Total Debt (mrq)"),
    18: ("Market Capitalization", "Yahoo Finance: LULU Market Cap"),
    19: ("Target Debt-to-Equity", "Yahoo Finance: debt (mrq) ÷ market cap"),
}

NUMFMT = {
    "day": "0.0",
    "pct": "0.0%",
}


def _row(ws, *prefixes: str, exclude: tuple[str, ...] = ()) -> int | None:
    for r in range(1, ws.max_row + 1):
        lab = str(ws.cell(r, 1).value or "").strip().lower()
        if any(lab.startswith(x.lower()) or lab == x.lower() for x in exclude):
            continue
        for p in prefixes:
            p = p.lower()
            if lab.startswith(p) or lab == p:
                return r
    return None


def _set_comment(cell, text: str) -> None:
    if not text:
        return
    cell.comment = Comment(text[:32000], "Analyst")


def round_scenarios(wb) -> int:
    ws = wb[SCN]
    n = 0
    for row, val in SCN_ROUNDS.items():
        for col in (3, 4, 5):
            cell = ws.cell(row, col)
            cell.value = val
            if row in (15, 16, 18):
                cell.number_format = NUMFMT["day"]
            else:
                cell.number_format = NUMFMT["pct"]
            n += 1
    return n


def _derived_growth_row(ws, r_rev: int) -> int | None:
    """Old AI row: Revenue growth % with =C{r}/{B}{r}-1 style formulas."""
    for r in range(r_rev + 1, min(r_rev + 4, ws.max_row + 1)):
        lab = str(ws.cell(r, 1).value or "").lower()
        if not lab.startswith("revenue growth"):
            continue
        cval = ws.cell(r, 3).value
        if isinstance(cval, str) and "/" in cval:
            return r
    return None


def _row_exact(ws, label: str) -> int | None:
    want = label.lower().strip()
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().lower() == want:
            return r
    return None


def fix_dcf_structure(wb) -> dict[str, int]:
    dcf = wb["DCF"]
    stats = {"inserts": 0, "formulas": 0}

    r_rev = _row(dcf, "net revenue")
    if r_rev is None:
        raise ValueError("Net revenue row not found")

    lab5 = str(dcf.cell(5, 1).value or "").lower()
    driver_ready = lab5.startswith("revenue growth") and r_rev == 6

    # --- 4. Insert growth driver row above Net Revenue ---
    if not driver_ready and r_rev == 5:
        dcf.insert_rows(5)
        stats["inserts"] += 1
        r_rev = 6

    derived = _derived_growth_row(dcf, r_rev)
    if derived:
        dcf.delete_rows(derived, 1)
        stats["inserts"] -= 1

    r_rev = _row(dcf, "net revenue") or r_rev

    dcf["A5"] = "Revenue Growth %"
    dcf["C5"] = f"={SCN}!$D$4"
    dcf["D5"] = f"={SCN}!$D$5"
    for col in FCOLS[2:]:
        dcf[f"{col}5"] = f"={SCN}!$D$5"
    stats["formulas"] += 1

    # Net revenue: prior × (1 + growth), drag right
    dcf[f"C{r_rev}"] = f"={FY25}{r_rev}*(1+C5)"
    prev = FY25
    for col in FCOLS:
        dcf[f"{col}{r_rev}"] = f"={prev}{r_rev}*(1+{col}5)"
        prev = col
    stats["formulas"] += 1

    # --- 3. FY25 D&A line on DCF (insert before forecast formula pass) ---
    r_da_pct = _row(dcf, "d&a % of revenue")
    r_fy25_da = _row(dcf, "fy25 d&a")
    if r_da_pct and r_fy25_da is None:
        dcf.insert_rows(r_da_pct)
        stats["inserts"] += 1
        r_fy25_da = r_da_pct

    if r_fy25_da:
        dcf.cell(r_fy25_da, 1).value = "FY25 D&A ($000)"
        dcf.cell(r_fy25_da, 2).value = FY25_DA
        for col in FCOLS:
            dcf[f"{col}{r_fy25_da}"] = None
        _set_comment(
            dcf.cell(r_fy25_da, 2),
            "FY25 D&A from 10-K cash flow statement (Depreciation and amortization).",
        )

    r_gm = _row(dcf, "gross margin")
    r_cogs = _row(dcf, "cogs", "cost of goods sold")
    r_da = _row(dcf, "plus: d&a", exclude=("fy25", "d&a %"))
    r_cx = _row(dcf, "less: capex", "capex")

    # --- 2. Draggable EBIT margin ramp ---
    r_em = _row(dcf, "ebit margin")
    if r_em:
        dcf[f"C{r_em}"] = f"={SCN}!$D$6"
        dcf[f"D{r_em}"] = f"=C{r_em}+{STEP}"
        for i, col in enumerate(FCOLS[2:], start=2):
            prev_col = FCOLS[i - 1]
            dcf[f"{col}{r_em}"] = f"={prev_col}{r_em}+{STEP}"
        stats["formulas"] += 1

    # Fix tariff ref if stale col B
    r_tar = _row(dcf, "tariff")
    if r_tar:
        for col in FCOLS:
            v = dcf[f"{col}{r_tar}"].value
            if isinstance(v, str) and "$B$7" in v:
                dcf[f"{col}{r_tar}"] = v.replace("$B$7", "$D$7")

    # Re-anchor P&L lines to net revenue row
    if r_gm and r_cogs:
        for col in FCOLS:
            dcf[f"{col}{r_cogs}"] = f"={col}{r_rev}*(1-{col}{r_gm})"
    r_ebit = _row_exact(dcf, "EBIT")
    if r_ebit and r_em:
        if r_tar:
            dcf[f"C{r_ebit}"] = f"=C{r_rev}*C{r_em}+C{r_tar}"
            for col in FCOLS[1:]:
                dcf[f"{col}{r_ebit}"] = f"={col}{r_rev}*{col}{r_em}"
        else:
            for col in FCOLS:
                dcf[f"{col}{r_ebit}"] = f"={col}{r_rev}*{col}{r_em}"
    r_dap = _row(dcf, "d&a % of revenue")
    r_cxp = _row(dcf, "capex % of revenue")
    if r_da and r_dap:
        for col in FCOLS:
            dcf[f"{col}{r_da}"] = f"={col}{r_rev}*{SCN}!$D$12"
    if r_cx and r_cxp:
        for col in FCOLS:
            dcf[f"{col}{r_cx}"] = f"=-{col}{r_rev}*{SCN}!$D$13"

    # Reported EBIT margin should divide by net revenue, not growth row
    r_ebit_m = _row(dcf, "reported ebit margin")
    if r_ebit_m and r_ebit:
        for col in FCOLS:
            dcf[f"{col}{r_ebit_m}"] = f"={col}{r_ebit}/{col}{r_rev}"

    # D&A % row is an assumption link, not a revenue calc
    if r_dap:
        for col in FCOLS:
            dcf[f"{col}{r_dap}"] = f"={SCN}!$D$12"

    r_ebit_row = r_ebit or _row_exact(dcf, "EBIT") or 11
    r_fy25_da = r_fy25_da or _row(dcf, "fy25 d&a") or 15
    comps = wb["Comps"]
    comps["B5"] = f"=DCF!B{r_ebit_row}+DCF!B{r_fy25_da}"
    stats["formulas"] += 1

    return stats


def rename_wacc(wb) -> int:
    ws = wb["WACC"]
    n = 0
    for row, (label, src) in WACC_RENAMES.items():
        old = ws.cell(row, 1).value
        if old != label:
            ws.cell(row, 1).value = label
            n += 1
        if src:
            cell = ws.cell(row, 2)
            if cell.value is not None:
                existing = cell.comment.text if cell.comment else ""
                if src not in existing:
                    _set_comment(cell, f"{src}\n{existing}".strip())
    return n


def delete_scratch(wb) -> bool:
    if "Scratch" in wb.sheetnames:
        del wb["Scratch"]
        return True
    return False


def fix(path: Path = TARGET) -> dict:
    tmp = path.with_suffix(".human.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)

    result = {
        "scenarios_rounded": round_scenarios(wb),
        "dcf": fix_dcf_structure(wb),
        "wacc_renamed": rename_wacc(wb),
        "scratch_deleted": delete_scratch(wb),
    }

    wb.save(tmp)
    tmp.replace(path)
    return result


if __name__ == "__main__":
    print(f"Human analyst structure fix: {fix()}")
