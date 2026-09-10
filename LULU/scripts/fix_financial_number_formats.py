#!/usr/bin/env python3
"""Apply analyst-style number formats so large $000 figures never show as scientific notation.

Excel 'General' on values >= ~1e7 displays as 1.11026E+07 — an immediate AI tell.
This pass sets row-aware formats on hardcodes AND formulas across the model grid.

Run:  cd LULU && python3 scripts/fix_financial_number_formats.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

SHEETS = ("WACC", "Scenarios", "NOPAT Bridge", "DCF", "Comps", "Revenue Drivers")

# Value columns per sheet (1-based column indices)
VALUE_COLS: dict[str, tuple[int, ...]] = {
    "WACC": (2, 4, 5),
    "Scenarios": (2, 3, 4, 5),
    "NOPAT Bridge": (2, 3, 4, 5, 6, 7),
    "DCF": (2, 3, 4, 5, 6, 7),
    "Comps": (2, 3, 4, 5, 6, 7),
    "Revenue Drivers": (2, 3, 4, 5, 6, 7, 8, 9, 10),
}

NUM = "#,##0;(#,##0)"
NUM0 = "#,##0"
PCT = "0.0%"
MULT = "0.0x"
MONEY = "$#,##0.00"
DAY = "0.0"
FACTOR = "0.000"

SKIP_LABEL = re.compile(
    r"sanity check|pass|review|selected exit is|spread vs|financial statement bridge|"
    r"cross-check|valuation —|scenario analysis|key assumptions|forecast paths|"
    r"pitch deck|memo:|enterprise value bridge|scenarios tie-out",
    re.I,
)


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _format_for_label(label: str) -> str | None:
    lab = label.lower().strip()
    if not lab or SKIP_LABEL.search(lab):
        return None

    if "variance (%)" in lab or (lab.endswith("(%)") and "variance" in lab):
        return PCT
    if any(k in lab for k in ("dso", "dio", "dpo")) or ("days" in lab and "decline" not in lab):
        return DAY
    if "decline (days" in lab:
        return DAY
    if any(k in lab for k in ("ev/ebitda", "exit multiple", "implied exit", "current ev /")):
        return MULT
    if "discount factor" in lab:
        return FACTOR
    # Dollar tax / EBIT rows — must come before generic "tax" → % rule
    if any(
        k in lab
        for k in (
            "less: taxes",
            "unlevered tax on",
            "tax on normalized",
            "normalized ebit",
        )
    ):
        return NUM
    if any(k in lab for k in ("tax rate", "effective tax rate", "cash tax rate", "marginal tax")):
        return PCT
    if any(
        k in lab
        for k in (
            "share price",
            "per share",
            "implied value per share",
            "implied share price",
            "accretive eps",
            "diluted eps",
        )
    ):
        return MONEY
    if "shares" in lab and "000" in lab:
        return NUM0
    if any(
        k in lab
        for k in (
            "growth",
            "margin",
            " rate",
            "wacc",
            "weight",
            "mix %",
            "mix",
            "erp",
            "premium",
            "cost of equity",
            "cost of debt",
            "beta",
            "gross margin",
            "prepaid",
            "accrued liabilities (%)",
            "capex %",
            "d&a %",
            "fcf to buybacks",
            "buybacks (bull)",
        )
    ):
        if "%" in lab or any(
            k in lab for k in ("growth", "margin", "rate", "weight", "mix", "wacc", "erp")
        ):
            return PCT
    if any(
        k in lab
        for k in (
            "$000",
            "($000",
            "revenue",
            "ebit",
            "cogs",
            "cash",
            "debt",
            "capex",
            "d&a",
            "nopat",
            "fcf",
            "free cash",
            "assets",
            "liabilit",
            "equity",
            "enterprise",
            "terminal value",
            "buyback",
            "repurchase",
            "net income",
            "inventory",
            "payable",
            "accrued",
            "pv of",
            "operating income",
            "projected net",
            "tariff",
            "refund",
            "sbc",
            "lease",
            "impairment",
            "retired",
            "ufcf",
            "ebitda",
            "market cap",
            "market value",
            "lease liab",
            "fund",
            "goodwill",
            "capital",
        )
    ):
        return NUM
    return None


def fix(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".numfmt.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"by_label": 0, "propagate": 0, "general_sweep": 0}

    for sn in SHEETS:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        cols = VALUE_COLS.get(sn, (2, 3, 4, 5, 6, 7))

        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "")
            fmt = _format_for_label(label)
            if not fmt:
                continue
            for c in cols:
                cell = ws.cell(r, c)
                if cell.value is None:
                    continue
                if isinstance(cell.value, str) and not _is_formula(cell.value):
                    continue
                if cell.number_format != fmt:
                    cell.number_format = fmt
                    stats["by_label"] += 1

        # Propagate: if col B on a row has comma format, match C–G formulas on same row
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2)
            bfmt = str(b.number_format or "General")
            if bfmt in ("General", "0") or "%" in bfmt or "x" in bfmt:
                continue
            for c in cols:
                if c == 2:
                    continue
                cell = ws.cell(r, c)
                if cell.value is None:
                    continue
                if not (_is_formula(cell.value) or isinstance(cell.value, (int, float))):
                    continue
                if str(cell.number_format or "General") == "General":
                    cell.number_format = bfmt
                    stats["propagate"] += 1

    # Final sweep: any remaining General on numeric/formula cells in value cols
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").lower()
        fmt_from_label = _format_for_label(str(ws.cell(r, 1).value or ""))
        for c in cols:
            cell = ws.cell(r, c)
            if str(cell.number_format or "General") != "General":
                continue
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and not _is_formula(v):
                continue
            if isinstance(v, (int, float)) and abs(v) < 1000 and fmt_from_label != PCT:
                if not any(k in label for k in ("price", "share", "shares", "000")):
                    continue
            fmt = fmt_from_label or NUM
            cell.number_format = fmt
            stats["general_sweep"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"Financial number formats: {fix()}")
