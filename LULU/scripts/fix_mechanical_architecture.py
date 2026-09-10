#!/usr/bin/env python3
"""Fix Excel formula mechanics without changing hardcoded assumptions.

Run:  cd LULU && python3 scripts/fix_mechanical_architecture.py
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from polish_model18_altered import (  # noqa: E402
    _rewrite_all_formulas,
    _rewrite_sheet_internal_formulas,
)

ROOT = SCRIPTS.parent
TARGET = ROOT / "unbeiesgbar_final.xlsx"

SCN = "Scenarios"
FCOLS = ["C", "D", "E", "F", "G"]
FY25 = "B"
ACTIVE = "DCF!$B$3"
# After Source column B is removed, Bear/Base/Bull live in B/C/D.
BEAR, BASE, BULL = "B", "C", "D"

SOURCE_COL_BY_SHEET: dict[str, int] = {
    "WACC": 3,
    "Scenarios": 2,
    "NOPAT Bridge": 8,
    "DCF": 8,
    "Comps": 8,
    "Revenue Drivers": 10,
}


def _choose(row: int) -> str:
    return (
        f"=CHOOSE({ACTIVE},Scenarios!${BEAR}${row},"
        f"Scenarios!${BASE}${row},Scenarios!${BULL}${row})"
    )


def _row(ws, *prefixes: str, exact: str | None = None) -> int | None:
    for r in range(1, ws.max_row + 1):
        lab = str(ws.cell(r, 1).value or "").strip()
        low = lab.lower()
        if exact and low == exact.lower():
            return r
        for p in prefixes:
            pl = p.lower()
            if pl == "ebit" and "margin" in low:
                continue
            if low.startswith(pl) or low == pl:
                return r
    return None


def _margin_step() -> str:
    return (
        f"(CHOOSE({ACTIVE},Scenarios!${BEAR}$8,Scenarios!${BASE}$8,Scenarios!${BULL}$8)"
        f"-CHOOSE({ACTIVE},Scenarios!${BEAR}$6,Scenarios!${BASE}$6,Scenarios!${BULL}$6))/4"
    )


def add_scenario_toggle(wb) -> None:
    dcf = wb["DCF"]
    dcf["A3"] = "Active case (1=Bear, 2=Base, 3=Bull)"
    if dcf["B3"].value not in (1, 2, 3):
        dcf["B3"] = 2


def fix_dcf_mechanics(wb) -> dict[str, int]:
    dcf = wb["DCF"]
    scn = wb[SCN]
    n = 0

    r_rev = _row(dcf, "net revenue") or 6
    r_g = _row(dcf, "revenue growth") or 5
    r_gm = _row(dcf, "gross margin") or 7
    r_cogs = _row(dcf, "cogs", "cost of goods sold") or 8
    r_em = _row(dcf, "ebit margin") or 9
    r_tar = _row(dcf, "tariff") or 10
    r_ebit = _row(dcf, exact="EBIT") or 11
    r_ebit_m = _row(dcf, "reported ebit margin") or 12
    r_tax = _row(dcf, "less: taxes") or 13
    r_nopat = _row(dcf, "nopat") or 14
    r_dap = _row(dcf, "d&a %") or 16
    r_da = _row(dcf, "plus: d&a", "depreciation") or 17
    r_cxp = _row(dcf, "capex %") or 18
    r_cx = _row(dcf, "less: capex", "capital expenditures") or 19
    r_dso = _row(dcf, "dso") or 21
    r_ar = _row(dcf, "accounts receivable") or 22
    r_dio = _row(dcf, "dio") or 23
    r_inv = _row(dcf, "inventories") or 24
    r_dpo = _row(dcf, "dpo") or 25
    r_ap = _row(dcf, "accounts payable") or 26
    r_ppct = _row(dcf, "prepaid expenses (%)") or 27
    r_pp = _row(dcf, "prepaid expenses (other") or 28
    r_accpct = _row(dcf, "accrued liabilities (%)") or 29
    r_acc = _row(dcf, "accrued liabilities and") or 30
    r_coa = _row(dcf, "current operating assets") or 31
    r_col = _row(dcf, "current operating liabilities") or 32
    r_nwc = _row(dcf, "net working capital") or 33
    r_dnwc = _row(dcf, "delta nwc") or 34
    r_ufcf = _row(dcf, "unlevered free cash") or 36

    gm_assump = _choose(14)
    step = _margin_step()
    run_rate = _choose(6)
    da_pct = _choose(12)
    cx_pct = _choose(13)
    dso = _choose(15)
    dio0 = _choose(16)
    dio_decl = _choose(17)
    dpo = _choose(18)
    pp_pct = _choose(19)
    acc_pct = _choose(20)

    dcf[f"C{r_g}"] = _choose(4)
    for col in FCOLS[1:]:
        dcf[f"{col}{r_g}"] = _choose(5)
        n += 1

    dcf[f"C{r_rev}"] = f"={FY25}{r_rev}*(1+C{r_g})"
    prev = FY25
    for col in FCOLS:
        dcf[f"{col}{r_rev}"] = f"={prev}{r_rev}*(1+{col}{r_g})"
        prev = col
        n += 1

    for col in FCOLS:
        dcf[f"{col}{r_cogs}"] = f"={col}{r_rev}*(1-{gm_assump.lstrip('=')})"
        dcf[f"{col}{r_gm}"] = f"=({col}{r_rev}-{col}{r_cogs})/{col}{r_rev}"
        n += 2

    dcf[f"C{r_em}"] = run_rate
    for i, col in enumerate(FCOLS[1:], start=1):
        dcf[f"{col}{r_em}"] = f"={FCOLS[i - 1]}{r_em}+{step}"
        n += 1

    dcf[f"C{r_tar}"] = _choose(7)
    for col in FCOLS[1:]:
        dcf[f"{col}{r_tar}"] = 0

    dcf[f"C{r_ebit}"] = f"=C{r_rev}*C{r_em}+C{r_tar}"
    for col in FCOLS[1:]:
        dcf[f"{col}{r_ebit}"] = f"={col}{r_rev}*{col}{r_em}"
        n += 1

    for col in FCOLS:
        dcf[f"{col}{r_ebit_m}"] = f"={col}{r_ebit}/{col}{r_rev}"
        n += 1

    for col in FCOLS:
        dcf[f"{col}{r_tax}"] = f"=-'NOPAT Bridge'!{col}40"
        dcf[f"{col}{r_nopat}"] = f"='NOPAT Bridge'!{col}41"
        n += 2

    for col in FCOLS:
        dcf[f"{col}{r_dap}"] = da_pct
        dcf[f"{col}{r_da}"] = f"={col}{r_rev}*{da_pct.lstrip('=')}"
        dcf[f"{col}{r_cxp}"] = cx_pct
        dcf[f"{col}{r_cx}"] = f"=-{col}{r_rev}*{cx_pct.lstrip('=')}"
        n += 4

    for i, col in enumerate(FCOLS):
        dcf[f"{col}{r_dso}"] = dso
        dcf[f"{col}{r_dio}"] = f"={dio0.lstrip('=')}-{i}*{dio_decl.lstrip('=')}"
        dcf[f"{col}{r_dpo}"] = dpo
        dcf[f"{col}{r_ar}"] = f"={col}{r_rev}*{col}{r_dso}/365"
        dcf[f"{col}{r_inv}"] = f"={col}{r_cogs}*{col}{r_dio}/365"
        dcf[f"{col}{r_ap}"] = f"={col}{r_cogs}*{col}{r_dpo}/365"
        dcf[f"{col}{r_ppct}"] = pp_pct
        dcf[f"{col}{r_pp}"] = f"={col}{r_rev}*{col}{r_ppct}"
        dcf[f"{col}{r_accpct}"] = acc_pct
        dcf[f"{col}{r_acc}"] = f"={col}{r_rev}*{col}{r_accpct}"
        dcf[f"{col}{r_coa}"] = f"={col}{r_ar}+{col}{r_inv}+{col}{r_pp}"
        dcf[f"{col}{r_col}"] = f"={col}{r_ap}+{col}{r_acc}"
        dcf[f"{col}{r_nwc}"] = f"={col}{r_coa}-{col}{r_col}"
        n += 12

    dcf[f"C{r_dnwc}"] = f"={FY25}{r_nwc}-C{r_nwc}"
    for i, col in enumerate(FCOLS[1:], start=1):
        dcf[f"{col}{r_dnwc}"] = f"={FCOLS[i - 1]}{r_nwc}-{col}{r_nwc}"
        n += 1
    for col in FCOLS:
        dcf[f"{col}{r_ufcf}"] = f"={col}{r_nopat}+{col}{r_da}+{col}{r_cx}+{col}{r_dnwc}"
        n += 1

    scn["A24"] = "Forecast summary — active case ($000)"
    for i, h in enumerate(["FY26E", "FY27E", "FY28E", "FY29E", "FY30E"]):
        scn.cell(24, 4 + i).value = h
    summary = [
        ("Net revenue ($000)", r_rev),
        ("COGS ($000)", r_cogs),
        ("Gross margin %", r_gm),
        ("EBIT margin %", r_em),
        ("EBIT ($000)", r_ebit),
        ("NOPAT ($000)", r_nopat),
        ("Unlevered FCF ($000)", r_ufcf),
    ]
    for i, (label, rr) in enumerate(summary):
        r = 25 + i
        scn.cell(r, 1).value = label
        for j, fc in enumerate(FCOLS):
            scn.cell(r, 4 + j).value = f"=DCF!{fc}{rr}"
            n += 1

    return {"dcf_formulas": n}


def fix_nopat_mechanics(wb) -> int:
    nb = wb["NOPAT Bridge"]
    n = 0
    sbc_pct = "$B$20/DCF!$B$6"
    for col in FCOLS:
        nb[f"{col}20"] = f"=IF($B$7=1,DCF!{col}6*{sbc_pct},0)"
        nb[f"{col}23"] = f"=$B$23*(DCF!{col}6/DCF!$B$6)"
        nb[f"{col}16"] = f"=DCF!{col}11"
        nb[f"{col}36"] = f"=DCF!{col}11"
        n += 4
    return n


def _build_comment(source, url: str | None) -> str | None:
    parts: list[str] = []
    if source and str(source).strip() not in ("Notes", "Source", ""):
        parts.append(f"Source: {str(source).strip()}")
    if url and str(url).startswith("http"):
        parts.append(str(url))
    return "\n".join(parts)[:32000] if parts else None


def _has_source_column(ws, src_col: int) -> bool:
    for r in range(1, min(ws.max_row, 25) + 1):
        v = ws.cell(r, src_col).value
        if isinstance(v, str) and v.strip().lower() == "source":
            return True
    return False


def _value_cols(sheet: str) -> list[int]:
    if sheet == SCN:
        return [3, 4, 5]
    if sheet == "WACC":
        return [2]
    if sheet == "Revenue Drivers":
        return list(range(2, 9))
    return list(range(2, 8))


def embed_sources_delete_columns(wb) -> dict[str, int]:
    stats = {"comments": 0, "deleted": 0, "skipped": 0}
    for sheet, src_col in sorted(SOURCE_COL_BY_SHEET.items(), key=lambda x: -x[1]):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        if not _has_source_column(ws, src_col):
            stats["skipped"] += 1
            continue
        pending: list[tuple[int, str, int]] = []
        for r in range(1, ws.max_row + 1):
            src = ws.cell(r, src_col)
            if not src.value and not src.hyperlink:
                continue
            label = ws.cell(r, 1).value
            if isinstance(label, str) and label.strip() in ("Source", "Notes", "Driver"):
                continue
            url = src.hyperlink.target if src.hyperlink else None
            body = _build_comment(src.value, url)
            if not body:
                continue
            pending.append((r, body, _value_cols(sheet)[0]))

        max_col = ws.max_column + 1
        shift = {get_column_letter(c): get_column_letter(c - 1) for c in range(src_col + 1, max_col + 1)}
        ws.delete_cols(src_col, 1)
        _rewrite_all_formulas(wb, sheet, shift)
        _rewrite_sheet_internal_formulas(ws, shift)

        for r, body, target_col in pending:
            new_col = target_col - 1 if target_col > src_col else target_col
            cell = ws.cell(r, new_col)
            if cell.value is None and not isinstance(cell.value, (int, float)):
                continue
            if cell.comment and str(cell.comment.text).strip() == body.strip():
                continue
            cell.comment = Comment(body, "Analyst")
            stats["comments"] += 1
        stats["deleted"] += 1
    return stats


def verify_no_circular(wb) -> list[str]:
    issues = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                coord = cell.coordinate.replace("$", "")
                if coord in v.replace("'", ""):
                    issues.append(f"{ws.title}!{cell.coordinate}: {v}")
    return issues


def fix(path: Path = TARGET) -> dict:
    tmp = path.with_suffix(".mech.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    add_scenario_toggle(wb)
    stats: dict = {}
    stats["sources"] = embed_sources_delete_columns(wb)
    stats.update(fix_dcf_mechanics(wb))
    stats["nopat"] = fix_nopat_mechanics(wb)
    circular = verify_no_circular(wb)
    stats["circular_remaining"] = len(circular)
    if circular:
        stats["circular_samples"] = circular[:5]
    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    result = fix()
    print(f"Mechanical architecture fix: {result}")
    if result.get("circular_remaining"):
        raise SystemExit(1)
