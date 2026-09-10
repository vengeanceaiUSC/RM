#!/usr/bin/env python3
"""Strip instructional manual from unbeiesgbar2model.xlsx → analyst-facing workbook.

Removes Justification/Source/Proof columns, memo rows, beta walkthroughs,
template branding, and programmatic phase labels. Preserves formulas and
hardcoded values (with column refs rebound).

Run:  cd LULU && python3 scripts/strip_model_manual.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from polish_model18_altered import (  # noqa: E402
    _rewrite_all_formulas,
    _rewrite_sheet_internal_formulas,
)
from restore_outline_groups import restore_outline_groups  # noqa: E402

ROOT = SCRIPTS.parent
INPUT = ROOT / "unbeiesgbar2model.xlsx"
OUTPUT = ROOT / "unbeiesgbar2model.xlsx"

# (start_col_letter, end_col_letter) doc blocks to delete per sheet.
DOC_COL_RANGES: dict[str, list[tuple[str, str]]] = {
    "WACC": [("B", "D")],
    "Scenarios": [("B", "D")],
    "NOPAT Bridge": [("B", "D")],
    "Comps": [("B", "D")],
    "DCF": [("B", "D"), ("L", "M")],
    "Revenue Drivers": [("I", "L")],
}

LABEL_RENAMES = [
    (re.compile(r"Phase 1 — Operating normalization", re.I), "EBIT adjustments"),
    (re.compile(r"Phase 1:\s*Add back SBC\?.*", re.I), "SBC add-back flag (0 = expensed)"),
    (re.compile(r"Phase 2 — Lease & capitalization.*", re.I), "Lease & capitalization"),
    (re.compile(r"Phase 2:\s*", re.I), ""),
    (re.compile(r"Phase 3 — Segment / channel EBIT.*", re.I), "Segment EBIT bridge"),
    (re.compile(r"Phase 3:\s*", re.I), ""),
    (re.compile(r"Phase 4 — Tax normalization to NOPAT", re.I), "Tax normalization"),
    (re.compile(r"Phase 4:\s*", re.I), ""),
    (re.compile(r"Phase 5 — Reconciliation", re.I), "Reconciliation"),
    (re.compile(r"EBIT_adj \(Phase 1\)", re.I), "Adjusted EBIT"),
    (re.compile(r"\(Phase [1-4]\)", re.I), ""),
    (re.compile(r"Phase 5 summarizes the reconciliation workflow\.?", re.I), ""),
    (re.compile(r"Phases 1–4 clean and forecast operating profit;?\s*", re.I), ""),
    (re.compile(r"Reported EBIT → Normalized NOPAT \(5-phase pipeline\)", re.I),
     "Reported EBIT → Normalized NOPAT"),
    (re.compile(r"Full Assumptions Guide \(PDF\)", re.I), None),
    (re.compile(r"^\s*memo:\s*", re.I), ""),
    (re.compile(r"\bConvention A\b[^.]*\.?", re.I), ""),
    (re.compile(r"\bplugged:\s*", re.I), ""),
    (re.compile(r"β walkthrough \(Hamada; Yahoo Key Statistics source\):", re.I), ""),
    (re.compile(r"^\(\d+\)\s*", re.I), ""),
]

ROW_DELETE_PATTERNS = [
    re.compile(r"^\s*memo:", re.I),
    re.compile(r"^current \$298,724k \+ non-current", re.I),
    re.compile(r"β walkthrough", re.I),
    re.compile(r"^\(\d+\)\s+Yahoo βL", re.I),
    re.compile(r"^\(\d+\)\s+Unlever at Yahoo", re.I),
    re.compile(r"^\(\d+\)\s+Relever at WACC", re.I),
    re.compile(r"^Yahoo βL = 0\.86 \(Beta", re.I),
    re.compile(r"^Unlever at Yahoo market D/E", re.I),
    re.compile(r"^Relever at WACC D/E", re.I),
    re.compile(r"^Book D/E 44\.69% on Yahoo page", re.I),
]

BANNED_TEXT = [
    "justification",
    "click + to expand",
    "assumptions guide",
    "font / color convention",
    "convention a",
    "phase 1",
    "phase 2",
    "phase 3",
    "phase 5",
    "built from scratch for the gis",
    "global investment society",
]


def _col_shift_for_delete(start: str, end: str, max_col: int = 30) -> dict[str, str]:
    s, e = column_index_from_string(start), column_index_from_string(end)
    delta = e - s + 1
    shift: dict[str, str] = {}
    for c in range(1, max_col + 1):
        if s <= c <= e:
            continue
        old = get_column_letter(c)
        if c > e:
            shift[old] = get_column_letter(c - delta)
    return shift


def _delete_col_range(wb, sheet_name: str, start: str, end: str) -> None:
    ws = wb[sheet_name]
    s_idx = column_index_from_string(start)
    count = column_index_from_string(end) - s_idx + 1
    col_shift = _col_shift_for_delete(start, end, max_col=ws.max_column + count)
    ws.delete_cols(s_idx, count)
    _rewrite_all_formulas(wb, sheet_name, col_shift)
    _rewrite_sheet_internal_formulas(ws, col_shift)


def _rename_labels(ws) -> int:
    changed = 0
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str) or cell.value.startswith("="):
                continue
            original = cell.value
            new = original
            cleared = False
            for pat, repl in LABEL_RENAMES:
                if repl is None:
                    if pat.search(new):
                        cleared = True
                        break
                    continue
                new = pat.sub(repl, new)
            if cleared:
                if cell.value is not None:
                    cell.value = None
                    changed += 1
                continue
            new = re.sub(r"\s{2,}", " ", new).strip()
            if new != original:
                cell.value = new if new else None
                changed += 1
    return changed


def _delete_matching_rows(ws) -> int:
    to_del: list[int] = []
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        combined = label + str(ws.cell(r, 2).value or "")
        if any(p.search(label) or p.search(combined) for p in ROW_DELETE_PATTERNS):
            to_del.append(r)
    for r in sorted(to_del, reverse=True):
        ws.delete_rows(r, 1)
    return len(to_del)


def _simplify_cover(cov) -> None:
    cov["B2"].value = "lululemon athletica inc. (NASDAQ: LULU)"
    cov["B4"].value = "Unlevered DCF Valuation Model"
    cov["B5"].value = "FY2025A – FY2030E | US$ thousands"
    for r in range(9, 18):
        cov.cell(r, 2).value = None
    cov["B9"].value = "Tabs: WACC · Scenarios · Revenue Drivers · NOPAT Bridge · DCF · Comps"
    cov["B14"].value = None
    cov["B15"].value = None
    cov["B16"].value = None
    cov["B17"].value = None
    cov["B25"].value = None
    if cov["B7"].value:
        cov["B7"].value = str(cov["B7"].value).replace("  ", " ")


def _exact_label_row(ws, text: str, col: int = 1) -> int | None:
    target = text.strip().lower()
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").strip().lower() == target:
            return r
    return None


def _label_row(ws, *needles: str, col: int = 1) -> int | None:
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, col).value or "").lower()
        if all(n.lower() in label for n in needles):
            return r
    return None


def _rebind_wacc(ws) -> None:
    v = "B"

    def ref(row: int | None) -> str:
        if row is None:
            raise ValueError("Missing WACC label row")
        return f"{v}{row}"

    r_mkt = _label_row(ws, "market value of equity")
    r_px = _label_row(ws, "share price")
    r_sh = _label_row(ws, "shares outstanding")
    r_lease = _label_row(ws, "operating lease")
    r_funded = _label_row(ws, "funded debt")
    r_total = _label_row(ws, "total debt equivalents")
    r_obs = _label_row(ws, "observed beta")
    r_ydebt = _label_row(ws, "yahoo total debt")
    r_ymcap = _label_row(ws, "yahoo market cap")
    r_de_u = _label_row(ws, "d/e for unlever")
    r_bu = _label_row(ws, "unlevered")
    r_de_r = _label_row(ws, "d/e for relever")
    r_beta = _label_row(ws, "beta used")
    r_tax = _label_row(ws, "tax rate")
    r_pretax = _label_row(ws, "pre-tax cost of debt")
    r_after = _label_row(ws, "after-tax cost of debt")
    r_coe = _label_row(ws, "cost of equity = rf")
    r_eqw = _label_row(ws, "equity weight")
    r_dw = _label_row(ws, "debt weight")
    r_wacc = _exact_label_row(ws, "WACC")

    ws[f"{v}{r_mkt}"] = f"={v}{r_px}*{v}{r_sh}"
    ws[f"{v}{r_total}"] = f"={v}{r_lease}+{v}{r_funded}"
    ws[f"{v}{r_de_u}"] = f"={v}{r_ydebt}/{v}{r_ymcap}"
    ws[f"{v}{r_bu}"] = f"={v}{r_obs}/(1+(1-{v}{r_tax})*{v}{r_de_u})"
    ws[f"{v}{r_de_r}"] = f"={v}{r_total}/{v}{r_mkt}"
    ws[f"{v}{r_beta}"] = f"={v}{r_bu}*(1+(1-{v}{r_tax})*{v}{r_de_r})"
    ws[f"{v}{r_coe}"] = f"={v}3+{v}{r_beta}*{v}4"
    ws[f"{v}{r_after}"] = f"={v}{r_pretax}*(1-{v}{r_tax})"
    ws[f"{v}{r_eqw}"] = f"={v}{r_mkt}/({v}{r_mkt}+{v}{r_total})"
    ws[f"{v}{r_dw}"] = f"={v}{r_total}/({v}{r_mkt}+{v}{r_total})"
    ws[f"{v}{r_wacc}"] = f"={v}{r_eqw}*{v}{r_coe}+{v}{r_dw}*{v}{r_after}"


def _rebind_scenarios_wacc(wb) -> None:
    scn = wb["Scenarios"]
    wacc = wb["WACC"]
    r_wacc = _exact_label_row(wacc, "WACC")
    if r_wacc:
        # Base-case WACC input links to WACC tab (was col G row 9, now col D after strip).
        scn["D9"] = f"=WACC!B{r_wacc}"


def _rebind_formulas(wb) -> None:
    _rebind_wacc(wb["WACC"])
    _rebind_scenarios_wacc(wb)


def _fix_freeze_panes(wb) -> None:
    """After deleting doc cols B–D, value col is now B; freeze after column A."""
    for name in ("WACC", "Scenarios", "NOPAT Bridge", "DCF", "Comps"):
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        view = ws.sheet_view
        view.pane = None
        if hasattr(view, "selection") and view.selection:
            view.selection[0].activeCell = "B3"
            view.selection[0].sqref = "B3"


def strip_workbook(src: Path = INPUT, dst: Path = OUTPUT) -> Path:
    if not src.exists():
        raise FileNotFoundError(src)
    tmp = dst.with_suffix(".stripping.xlsx")
    shutil.copy2(src, tmp)
    wb = openpyxl.load_workbook(tmp)

    _simplify_cover(wb["Cover"])
    for sheet_name in wb.sheetnames:
        if sheet_name == "Cover":
            continue
        ws = wb[sheet_name]
        _rename_labels(ws)
        _delete_matching_rows(ws)

    # Delete doc columns last-to-first within each range set per sheet.
    for sheet_name, ranges in DOC_COL_RANGES.items():
        if sheet_name not in wb.sheetnames:
            continue
        for start, end in sorted(
            ranges,
            key=lambda r: column_index_from_string(r[1]),
            reverse=True,
        ):
            _delete_col_range(wb, sheet_name, start, end)

    _rebind_formulas(wb)
    restore_outline_groups(wb, col_hidden=False)
    _fix_freeze_panes(wb)

    wb.save(tmp)
    tmp.replace(dst)
    print(f"Saved {dst}")
    return dst


def verify(dst: Path = OUTPUT) -> None:
    wb = openpyxl.load_workbook(dst, data_only=False)
    hits: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                low = v.lower()
                for banned in BANNED_TEXT:
                    if banned in low:
                        hits.append(f"{ws.title}!{cell.coordinate}: {banned}")
                        break

    if hits:
        raise AssertionError("Banned template text remains:\n" + "\n".join(hits[:15]))

    wacc = wb["WACC"]
    r_beta = _label_row(wacc, "beta used")
    if not r_beta or not str(wacc[f"B{r_beta}"].value or "").startswith("="):
        raise AssertionError("Beta-used formula not found on WACC")

    assert wacc.max_column <= 3, f"WACC still has doc cols (max_col={wacc.max_column})"
    print(f"Verify OK: manual stripped; beta formula at WACC B{r_beta}")


if __name__ == "__main__":
    strip_workbook()
    verify()
