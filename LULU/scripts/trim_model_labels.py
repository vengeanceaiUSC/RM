#!/usr/bin/env python3
"""Trim verbose AI-style labels to standard finance shorthand.

Run:  cd LULU && python3 scripts/trim_model_labels.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# Exact or substring replacements on column A labels
LABEL_TRIMS: list[tuple[str, str]] = [
    ("Less: unlevered tax (t_operating from NOPAT Bridge)", "Less: Taxes"),
    ("Less: unlevered tax", "Less: Taxes"),
    ("Operating lease liabilities (ASC 842 debt equiv.)", "Capitalized Operating Leases"),
    ("Operating lease liabilities", "Capitalized Operating Leases"),
    ("Funded debt (term loans / bonds)", "Funded Debt"),
    ("Clean / run-rate EBIT margin %", "EBIT Margin %"),
    ("Cost of goods sold (COGS)", "COGS"),
    ("Plus: FY26 IEEPA tariff refunds (one-time)", "Tariff Refund (FY26)"),
    ("EBIT (incl. FY26 refund)", "EBIT"),
    ("Plus: depreciation & amortization", "Plus: D&A"),
    ("Less: capital expenditures", "Less: Capex"),
    ("DSO (days) - flat vs FY25", "DSO (days)"),
    ("DIO (days) - FY25 anchor, -1 day/yr", "DIO (days)"),
    ("DPO (days) - flat vs FY25", "DPO (days)"),
    ("FY2027-FY2030E revenue growth (avg)", "FY27-30 Rev Growth"),
    ("FY2026E revenue growth", "FY26 Rev Growth"),
    ("Run-rate EBIT margin (clean, ex-refunds)", "Run-Rate EBIT Margin"),
    ("FY26 IEEPA tariff refunds ($k, already in guide)", "FY26 Tariff Refund"),
    ("Terminal (FY2030E) EBIT margin", "Terminal EBIT Margin"),
    ("Observed Beta (5Y Monthly)", "Observed Beta (5Y)"),
    ("D/E for relever (WACC: FY25 lease debt / market equity)", "D/E (Relever)"),
    ("Sector beta-u benchmark (Damodaran Special Lines) - not used", "Sector Beta (unused)"),
    ("Beta used (relevered for WACC capital structure)", "Relevered Beta"),
    ("Pre-tax cost of debt (lease-equivalent)", "Pre-Tax Cost of Debt"),
    ("After-tax cost of debt", "After-Tax Cost of Debt"),
    ("Cost of equity = rf + beta x erp", "Cost of Equity"),
    ("Revenue - year", "Rev Y"),
    ("Gross margin % - year", "GM% Y"),
    ("Clean EBIT margin - year", "EBIT Margin Y"),
    ("Cost of goods sold (COGS) - year", "COGS Y"),
    ("EBIT - year", "EBIT Y"),
    ("NOPAT - year", "NOPAT Y"),
    ("D&A - year", "D&A Y"),
]

STRIP_PATTERNS = [
    re.compile(r"\s*\(Phase \d+\)\s*", re.I),
    re.compile(r"Phase \d+:\s*", re.I),
    re.compile(r"\s*-\s*flat vs FY25\s*", re.I),
]


def trim(path: Path = TARGET) -> dict[str, int]:
    tmp = path.with_suffix(".trim.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    stats = {"trimmed": 0}

    for ws in wb.worksheets:
        for row in ws.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            if not isinstance(cell.value, str) or cell.value.startswith("="):
                continue
            original = cell.value
            new = original
            for old, repl in LABEL_TRIMS:
                if old.lower() in new.lower():
                    new = re.sub(re.escape(old), repl, new, flags=re.I)
            for pat in STRIP_PATTERNS:
                new = pat.sub(" ", new)
            new = re.sub(r"\s{2,}", " ", new).strip()
            if new != original:
                cell.value = new
                stats["trimmed"] += 1

    wb.save(tmp)
    tmp.replace(path)
    return stats


if __name__ == "__main__":
    print(f"Trimmed labels: {trim()}")
