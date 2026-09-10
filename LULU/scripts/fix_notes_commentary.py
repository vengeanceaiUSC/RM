#!/usr/bin/env python3
"""Fix Notes column commentary — match actual model values, 8–15 words, no AI phrasing.

Run:  cd LULU && python3 scripts/fix_notes_commentary.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# (sheet, normalized label substring) → new note (8–15 words, states model value)
NOTE_FIXES: list[tuple[str, str, str]] = [
    # Capex — model uses 5.5%, not FY26 7% guide
    ("Scenarios", "capex % of revenue", "5.5% model rate; FY26 7% capex guide fades down."),
    ("DCF", "capex % of revenue", "5.5% model rate; FY26 7% capex guide fades down."),
    # WACC
    ("WACC", "equity risk premium", "Damodaran ERP plus 177 bps US risk premium."),
    # Scenarios working capital
    ("Scenarios", "dso (days)", "6.3 days DSO held flat vs FY25 actual."),
    ("Scenarios", "dpo (days)", "25.1 days DPO held flat vs FY25 actual."),
    # DCF bridge / WC schedule
    ("DCF", "gross margin %", "56.6% FY25 gross margin held flat in forecast."),
    ("DCF", "clean / run-rate ebit margin", "13.2% clean OM from Q2 run-rate ex tariff."),
    ("DCF", "d&a % of revenue", "4.5% of sales from FY25 D&A over revenue."),
    ("DCF", "working capital", "NWC schedule nets to 13.2% of FY25 sales."),
    ("DCF", "dso (days)", "6.3 days DSO held flat vs FY25 actual."),
    ("DCF", "accounts receivable, net", "AR 1.7% of sales from FY25 balance sheet."),
    ("DCF", "dio (days)", "128.8 day DIO anchor; minus 1 day per year."),
    ("DCF", "inventories", "Inventory dollar balance driven by DIO times COGS."),
    ("DCF", "dpo (days)", "25.1 days DPO held flat vs FY25 actual."),
    ("DCF", "accounts payable", "AP 3.0% of sales from FY25 payables balance."),
    ("DCF", "prepaid expenses (% of revenue)", "5.1% of revenue from FY25 other current assets."),
    ("DCF", "prepaid expenses (other current", "5.1% of revenue from FY25 other current assets."),
    ("DCF", "accrued liabilities (% of revenue)", "6.0% of revenue from FY25 accrued liabilities balance."),
    ("DCF", "accrued liabilities and other", "6.0% of revenue from FY25 accrued liabilities balance."),
    ("DCF", "current operating assets", "Operating assets sum to 22.1% of FY25 sales."),
    ("DCF", "current operating liabilities", "Operating liabilities sum to 9.0% of FY25 sales."),
    ("DCF", "net working capital", "Net working capital equals 13.2% of FY25 sales."),
    ("DCF", "delta nwc", "Year-over-year dollar change in net working capital schedule."),
    ("DCF", "selected exit ev/ebitda", "Gordon terminal value divided by FY30 EBITDA multiple."),
    ("DCF", "wacc \\ g", "Sensitivity grid; terminal g spans 1.5% to 3.0%."),
    ("DCF", "revenue growth %", "Minus 6.1% FY26 revenue per company guidance midpoint."),
    ("DCF", "plus: fy26 ieepa", "Add back 134.5M tariff refund in FY26 only."),
    ("DCF", "terminal growth rate", "2.25% terminal g anchored to FRED GDPC1 macro series."),
]

NOTE_COLS = {
    "WACC": 2,
    "Scenarios": 2,
    "NOPAT Bridge": 2,
    "Comps": 2,
    "DCF": 2,
    "Revenue Drivers": 9,
}

BAD_PATTERNS = re.compile(
    r"not a direct|not one %|not % of revenue|overlay|read-through|anchor\.|\.\.$|,\s*not\.$| for\.$| minus\.$| /\.$|÷\.$",
    re.I,
)


def _norm(label) -> str:
    if label is None:
        return ""
    return re.sub(r"\s+", " ", str(label).lower().strip())


def _word_count(text: str) -> int:
    return len(text.split())


def fix_notes(path: Path = TARGET) -> list[str]:
    tmp = path.with_suffix(".notes_fix.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    changes: list[str] = []

    for sheet, col in NOTE_COLS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            label = _norm(ws.cell(r, 1).value)
            if not label:
                continue
            cell = ws.cell(r, col)
            old = cell.value
            if not isinstance(old, str) or old.strip() in ("Notes", "Source"):
                continue

            new = None
            matches = [(needle, note) for sh, needle, note in NOTE_FIXES if sh == sheet and needle in label]
            if matches:
                # Prefer longest label match (e.g. "net working capital" over "working capital")
                needle, new = max(matches, key=lambda x: len(x[0]))

            # Catch truncated / AI residue not in explicit map
            if new is None and BAD_PATTERNS.search(old):
                continue  # flagged but no auto-fix without map entry

            if new and new != old:
                if not (8 <= _word_count(new) <= 15):
                    raise ValueError(f"{sheet} R{r}: note must be 8–15 words: {new!r}")
                cell.value = new
                changes.append(f"{sheet}!{cell.coordinate}: {old!r} → {new!r}")

    wb.save(tmp)
    tmp.replace(path)
    return changes


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    issues: list[str] = []

    for coord in ("DCF!B17", "Scenarios!B13"):
        sheet, cell = coord.split("!")
        val = str(wb[sheet][cell].value or "")
        if "5.5" not in val:
            issues.append(f"{coord} must state 5.5% model rate: {val!r}")

    for sheet, col in NOTE_COLS.items():
        if sheet not in wb.sheetnames or sheet not in ("WACC", "Scenarios", "DCF"):
            continue
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            note = ws.cell(r, col).value
            if not isinstance(note, str) or note in ("Notes", "Source"):
                continue
            if BAD_PATTERNS.search(note):
                issues.append(f"{sheet} B{r}: AI/truncated residue: {note!r}")

    if issues:
        raise AssertionError("Verify failed:\n" + "\n".join(issues[:20]))
    print("Verify OK: Capex notes state 5.5% model rate; no truncated AI phrasing")


if __name__ == "__main__":
    changes = fix_notes()
    print(f"Fixed {len(changes)} notes in {TARGET.name}")
    for c in changes:
        print(f"  {c}")
    verify()
