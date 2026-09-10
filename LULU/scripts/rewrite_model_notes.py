#!/usr/bin/env python3
"""Rewrite Notes column (col B) in model18altered.xlsx → model18_final.xlsx.

Run:  cd LULU/scripts && python3 rewrite_model_notes.py
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "model18altered.xlsx"
OUTPUT = ROOT / "model18_final.xlsx"

SHEETS = ("WACC", "Scenarios", "NOPAT Bridge", "Comps")

# Header cells: replace "Justification …" with "Notes"
HEADER_CELLS: dict[tuple[str, int], str] = {
    ("WACC", 2): "Notes",
    ("Scenarios", 3): "Notes",
    ("NOPAT Bridge", 2): "Notes",
    ("Comps", 2): "Notes",
    ("Comps", 16): "Notes",
    ("Comps", 41): "Notes",
}

# Label → Notes mapping (col A label on each sheet)
NOTES_MAP: dict[str, dict[str, str]] = {
    "WACC": {
        "Risk-free rate (10-yr UST)": "FRED DGS10 anchor (4.8%).",
        "Equity risk premium": "Damodaran implied ERP + 177bps overlay.",
        "Tax rate": "FY26 mgmt guidance (~30%).",
        "Share price ($)": "NASDAQ last sale (~$100).",
        "Shares outstanding (000)": "FY25 10-K share count.",
        "Market value of equity": "Price × shares outstanding.",
        "Operating lease liabilities (ASC 842 debt equiv.)": "ASC 842 lease debt equiv.",
        "Funded debt (term loans / bonds)": "No term debt (FY25 10-K).",
        "Observed Beta (5Y Monthly) — levered βL": "Yahoo βL (5Y monthly).",
        "Yahoo Total Debt (mrq)": "Yahoo MRQ total debt.",
        "Yahoo Market Cap (same page as βL)": "Price × shares (Yahoo page).",
        "D/E for unlever (Yahoo debt mrq ÷ Yahoo market cap)": "Yahoo MRQ D/E for Hamada.",
        "Yahoo book D/E (mrq) — reference only": "Yahoo book D/E reference.",
        "Unlevered βu": "Hamada unlever (Yahoo D/E).",
        "D/E for relever (WACC: FY25 lease debt / market equity)": "FY25 lease debt / mkt cap.",
        "Sector βu benchmark (Damodaran Special Lines) — not used": "Damodaran benchmark (unused).",
        "Beta used (relevered for WACC capital structure)": "Relevered β for CAPM.",
        "Pre-tax cost of debt (lease-equivalent)": "Lease-equivalent borrowing cost.",
        "Equity weight": "Mkt cap / total capital.",
        "Debt weight (leases + funded)": "Lease debt / total capital.",
    },
    "Scenarios": {
        "FY2026E revenue growth": "Q2 FY26 mgmt guidance midpoint.",
        "FY2027–FY2030E revenue growth (avg)": "StockAnalysis 3Y revenue forecast.",
        "Run-rate EBIT margin (clean, ex-refunds)": "Q2 FY26 run-rate OM (ex-tariff).",
        "FY26 IEEPA tariff refunds ($k, already in guide)": "FY26 one-time tariff refund.",
        "Terminal (FY2030E) EBIT margin": "Partial recovery vs FY25 OM.",
        "WACC": "Lease-adjusted CAPM (WACC tab).",
        "Terminal growth": "FRED GDPC1 anchor (~2.1%).",
        "Cash tax rate": "FY26 mgmt guidance (~30%).",
        "D&A % of revenue": "FY25 D&A run-rate vs sales.",
        "Capex % of revenue": "FY26 guide fade to 5.5%.",
        "Gross margin %": "Held flat vs FY25 actuals.",
        "DSO (days) — flat vs FY25": "Flat vs FY25 actuals.",
        "FY25 DIO anchor (days)": "FY25 anchor minus 1 day/yr.",
        "DIO decline (days / forecast yr)": "Minus 1 day per forecast yr.",
        "DPO (days) — flat vs FY25": "Flat vs FY25 actuals.",
        "Prepaid expenses (% of revenue)": "FY25 OCA % of revenue.",
        "Accrued liabilities (% of revenue)": "FY25 accrued % of revenue.",
    },
    "NOPAT Bridge": {
        "Phase 1: Add back SBC? (0 = no — Convention A: expense stays in EBIT)": (
            "SBC flag (Convention A = 0)."
        ),
        "Phase 3: Store-channel EBIT margin %": "Store EBIT % of store rev.",
        "Phase 3: E-commerce EBIT margin %": "E-comm EBIT % of e-comm rev.",
        "Phase 3: Other-channels EBIT margin %": "Other EBIT % of other rev.",
        "Phase 4: Terminal marginal tax rate": "Statutory marginal rate (30%).",
        "Phase 2: R&D / software amortization period (years)": "R&D capitalization period (yrs).",
        "Reported operating income (EBIT)": "10-K EBIT / Scenarios forecast.",
        "+ Impairment / intangible amortization (add-back)": "Run-rate intangible amort add-back.",
        "+ Restructuring costs (add-back)": "Non-recurring restructuring add-back.",
        "+ Legal / M&A one-offs (add-back)": "One-off legal / M&A strip.",
        "+ Stock-based compensation (add-back only if flag = 1)": "SBC add-back (if flag = 1).",
        "= Adjusted EBIT (Phase 1)": "Reported EBIT plus non-recurring add-backs.",
        "+ Implied lease interest expense (reclass from rent)": "Lease interest reclass (memo).",
        "+ Capitalize R&D / software (current-year expense)": "R&D cap immaterial this yr.",
        "− Amortization of prior capitalized intangibles": "Prior capitalized intangible amort.",
        "= EBIT after lease & capitalization (Phase 2)": "Phase 1 plus lease / R&D adjustments.",
        "Store-channel revenue": "Store rev from Revenue Drivers.",
        "E-commerce revenue": "E-comm rev from Revenue Drivers.",
        "Other channels revenue": "Other rev from Revenue Drivers.",
        "Store-channel EBIT (= Rev × store margin)": "Store EBIT ÷ store rev only.",
        "E-commerce EBIT (= Rev × e-comm margin)": "E-comm EBIT ÷ e-comm rev only.",
        "Other-channels EBIT (= Rev × other margin)": "Other EBIT ÷ other rev only.",
        "Channel EBIT (sum of sector EBIT)": "Sum of sector EBIT (Phase 3).",
        "= EBIT after channel mix (Phase 3)": "Channel-mix EBIT output (Phase 3).",
        "= Normalized EBIT (tax base for NOPAT)": "FY25 walk-through Phases 1–2.",
        "Operating effective tax rate (t_operating)": "Operating ETR with interest shield.",
        "Unlevered tax on normalized EBIT": "Normalized EBIT × t_operating.",
        "NORMALIZED NOPAT": "EBIT_norm × (1 − t_operating).",
    },
    "Comps": {
        "lululemon (LULU)": "PitchBook EV/TTM EBITDA comp.",
        "Under Armour (UAA)": "PitchBook EV/TTM EBITDA comp.",
        "adidas (ADS)": "PitchBook EV/TTM EBITDA comp.",
        "Nike (NKE)": "PitchBook EV/TTM EBITDA comp.",
        "Deckers (DECK)": "PitchBook EV/TTM EBITDA comp.",
        "Williams-Sonoma (WSM)": "PitchBook EV/TTM EBITDA comp.",
        "Crocs (CROX)": "PitchBook EV/TTM EBITDA comp.",
        "Movado (MOV)": "PitchBook EV/TTM EBITDA comp.",
        "Levi Strauss (LEVI)": "PitchBook EV/TTM EBITDA comp.",
        "La-Z-Boy (LZB)": "PitchBook EV/TTM EBITDA comp.",
        "Kontoor (KTB)": "PitchBook EV/TTM EBITDA comp.",
        "Core athletic / apparel mean (LULU / NKE / ADS / DECK / CROX / LEVI / KTB)": (
            "PitchBook core set ex-UAA."
        ),
        "All positive-EBITDA mean (ex-UAA; includes WSM / MOV / LZB)": (
            "Wider tape incl. adjacent retail."
        ),
        "EV / EBITDA (FY2030E terminal)": (
            "Lo 5.0x FY30E EBITDA trough.\nHi 8.0x DECK high-end cap."
        ),
        "DCF exit method (Gordon-implied on FY2030E EBITDA)": (
            "Gordon TV / FY30 EBITDA exit."
        ),
        "P / E (FY2026E EPS)": (
            "Lo 10x FY26E EPS trough.\nHi 18x FY26E EPS recovery."
        ),
    },
}


def _word_count(text: str) -> int:
    return len(text.split())


def _lines_ok(text: str, max_words: int = 8) -> bool:
    for line in str(text).split("\n"):
        line = line.strip()
        if line and _word_count(line) > max_words:
            return False
    return True


def _is_justification_header(val: str) -> bool:
    return bool(re.match(r"^Justification\b", val.strip(), re.I))


def rewrite_workbook(
    source: Path = SOURCE, output: Path = OUTPUT
) -> tuple[Path, int, list[tuple[str, int, str, str]]]:
    """Copy source → output, rewrite Notes col B. Returns (path, count, samples)."""
    shutil.copy2(source, output)
    wb = openpyxl.load_workbook(output)
    rewritten = 0
    samples: list[tuple[str, int, str, str]] = []

    for sheet_name in SHEETS:
        ws = wb[sheet_name]
        sheet_map = NOTES_MAP.get(sheet_name, {})

        for row in range(1, ws.max_row + 1):
            label = ws.cell(row, 1).value
            cell = ws.cell(row, 2)
            old_val = cell.value

            # Header cells
            key = (sheet_name, row)
            if key in HEADER_CELLS:
                new_val = HEADER_CELLS[key]
                if old_val != new_val:
                    if len(samples) < 5:
                        samples.append((sheet_name, row, str(old_val or ""), new_val))
                    cell.value = new_val
                    rewritten += 1
                continue

            if label is None or not isinstance(old_val, str) or not old_val.strip():
                continue

            label_str = str(label).strip()
            new_val = sheet_map.get(label_str)

            if new_val is None:
                # Fallback: strip Justification headers missed above
                if _is_justification_header(old_val):
                    new_val = "Notes"
                else:
                    continue

            if old_val != new_val:
                if len(samples) < 5:
                    samples.append((sheet_name, row, old_val, new_val))
                cell.value = new_val
                rewritten += 1

    wb.save(output)
    return output, rewritten, samples


def verify(path: Path = OUTPUT) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    issues: list[str] = []

    for sheet_name in SHEETS:
        ws = wb[sheet_name]
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row, 2).value
            if not isinstance(val, str) or not val.strip():
                continue
            if "Justification" in val:
                issues.append(f"{sheet_name} R{row}: still contains 'Justification'")
            if ";" in val:
                issues.append(f"{sheet_name} R{row}: semicolon in Notes")
            for i, line in enumerate(val.split("\n"), 1):
                line = line.strip()
                if line and _word_count(line) > 8:
                    issues.append(
                        f"{sheet_name} R{row} L{i}: { _word_count(line)} words — {line!r}"
                    )

    if issues:
        raise AssertionError("Verification failed:\n" + "\n".join(issues))

    print(f"Verified {path}: {len(issues)} issues")


if __name__ == "__main__":
    out, count, samples = rewrite_workbook()
    verify(out)
    print(f"Rewrote {count} cells → {out}")
    print("\nSample before/after:")
    for sheet, row, before, after in samples:
        print(f"  [{sheet} R{row}]")
        print(f"    BEFORE: {before[:120]}")
        print(f"    AFTER:  {after[:120]}")
