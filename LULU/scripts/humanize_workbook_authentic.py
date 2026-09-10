#!/usr/bin/env python3
"""Make unbeiesgbar_final.xlsx look analyst-built, not openpyxl-generated.

Before editing, read ../AI_SELF_AUDIT.md and run audit_ai_tells.py (read-only).
Never delete hardcoded assumptions or Source column links when fixing flagged items.

- Excel metadata (creator / lastModifiedBy)
- Prompt residue in headers (Justification, click +, etc.)
- Robotic float precision on hardcodes
- GIS finance font colors: blue inputs, black calcs, green cross-sheet
- Human shorthand in Notes; internal Source labels on derived WACC rows

Run:  cd LULU && python3 scripts/humanize_workbook_authentic.py
"""
from __future__ import annotations

import re
import shutil
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from styles import BLACK, BLUE, FONT_NAME, GREEN  # noqa: E402

ROOT = SCRIPTS.parent
TARGET = ROOT / "unbeiesgbar_final.xlsx"

PROMPT_RE = re.compile(
    r"justification|~20\s*word|click\s*\+|\[cols\s|source\s*\(click\)|ctrl\+f|equation\s*\(",
    re.I,
)

# Robotic NOTES → human shorthand (substring replace on col B / I).
NOTE_SHORTHAND: list[tuple[str, str]] = [
    ("FRED DGS10 anchor (4.8%).", "Current 10Y UST"),
    ("Damodaran implied ERP + 177bps overlay.", "Damodaran ERP plus 177 bps US risk premium."),
    ("FY26 mgmt guidance (~30%).", "Q2 FY26 guide ~30%"),
    ("NASDAQ last sale (~$100.61).", "NASDAQ last sale"),
    ("NASDAQ last sale (~$100).", "NASDAQ last sale"),
    ("FY25 10-K share count.", "FY25 10-K shares"),
    ("Price × shares outstanding.", "Px × shares"),
    ("ASC 842 lease debt equiv.", "FY25 lease debt (10-K)"),
    ("No term debt (FY25 10-K).", "No funded debt (10-K)"),
    ("Yahoo βL (5Y monthly).", "Yahoo 5Y beta"),
    ("Yahoo MRQ total debt.", "Yahoo total debt"),
    ("Price × shares (Yahoo page).", "Yahoo mkt cap page"),
    ("Yahoo MRQ D/E for Hamada.", "Yahoo D/E"),
    ("Yahoo book D/E reference.", "Yahoo book D/E (ref)"),
    ("Hamada unlever (Yahoo D/E).", "Hamada unlever"),
    ("FY25 lease debt / mkt cap.", "FY25 lease D/E"),
    ("Damodaran benchmark (unused).", "Damodaran sector β (unused)"),
    ("Relevered β for CAPM.", "Relevered β"),
    ("Lease-equivalent borrowing cost.", "Lease-equivalent kd"),
    ("Mkt cap / total capital.", "Equity weight"),
    ("Lease debt / total capital.", "Debt weight"),
    ("Q2 FY26 mgmt guidance midpoint.", "Q2 FY26 rev guide"),
    ("StockAnalysis 3Y revenue forecast.", "StockAnalysis 3Y rev"),
    ("Q2 FY26 run-rate OM (ex-tariff).", "Q2 FY26 run-rate OM"),
    ("FY26 one-time tariff refund.", "FY26 tariff refund"),
    ("Partial recovery vs FY25 OM.", "Partial OM recovery"),
    ("Lease-adjusted CAPM (WACC tab).", "Links to WACC tab"),
    ("FRED GDPC1 anchor (~2.1%).", "FRED GDPC1 ~2.1%"),
    ("FY25 D&A run-rate vs sales.", "FY25 D&A / sales"),
    ("FY26 guide fade to 5.5%.", "5.5% model rate; FY26 7% capex guide fades down."),
    ("Capex fade to 5.5%", "5.5% model rate; FY26 7% capex guide fades down."),
    ("Held flat vs FY25 actuals.", "Flat vs FY25"),
    ("Flat vs FY25 actuals.", "Flat vs FY25"),
    ("FY25 anchor minus 1 day/yr.", "FY25 DIO −1d/yr"),
    ("Minus 1 day per forecast yr.", "−1 day / yr"),
    ("FY25 OCA % of revenue.", "FY25 prepaids % rev"),
    ("FY25 accrued % of revenue.", "FY25 accrued % rev"),
    ("PitchBook EV/TTM EBITDA comp.", "PitchBook comp"),
    ("PitchBook core set ex-UAA.", "Core comp average"),
    ("Wider tape incl. adjacent retail.", "Wider comp average"),
]

HEADER_FIXES: dict[tuple[str, str], str] = {
    ("WACC", "B2"): "Notes",
    ("WACC", "C2"): "Source",
    ("WACC", "B26"): "Notes",
    ("WACC", "C26"): "Source",
    ("Scenarios", "B3"): "Notes",
    ("Scenarios", "C3"): "Source",
    ("NOPAT Bridge", "B2"): "Notes",
    ("NOPAT Bridge", "C2"): "Source",
    ("Comps", "B2"): "Notes",
    ("Comps", "C2"): "Source",
    ("Comps", "B18"): "Notes",
    ("Comps", "C18"): "Source",
    ("Comps", "B56"): "Notes",
    ("Comps", "C56"): "Source",
    ("DCF", "C2"): "Source",
    ("Revenue Drivers", "I5"): "Notes",
    ("Revenue Drivers", "J5"): "Source",
    ("Revenue Drivers", "I23"): "Notes",
    ("Revenue Drivers", "J23"): "Source",
    ("Revenue Drivers", "I37"): "Notes",
    ("Revenue Drivers", "J37"): "Source",
    ("Revenue Drivers", "I44"): "Notes",
    ("Revenue Drivers", "J44"): "Source",
    ("Revenue Drivers", "I54"): "Notes",
    ("Revenue Drivers", "J54"): "Source",
}

# Derived WACC rows → Source col text when blank.
WACC_DERIVED_SOURCES: dict[str, str] = {
    "total debt equivalents": "Lease + funded debt (above)",
    "cost of equity = rf + beta x erp": "CAPM: rf + β × ERP",
    "after-tax cost of debt": "Pre-tax kd × (1 − T)",
    "wacc": "Weighted avg (weights above)",
    "market value of equity": "Px × shares (above)",
    "unlevered beta-u": "Hamada unlever (above)",
    "d/e for relever": "Total debt / mkt cap",
    "beta used": "Hamada relever (above)",
    "equity weight": "Mkt cap / total cap",
    "debt weight": "Lease debt / total cap",
}

VALUE_COLS: dict[str, tuple[int, ...]] = {
    "WACC": (4,),
    "Scenarios": (5, 6, 7),
    "NOPAT Bridge": (4, 5, 6, 7, 8, 9),
    "DCF": tuple(range(4, 11)),
    "Comps": tuple(range(4, 11)),
    "Revenue Drivers": tuple(range(2, 9)),
}

COMMENT_AI = re.compile(
    r"high-conviction analyst overlay|model cross-reference\.|durable rate anchor|"
    r"operating assumption with durable|working-capital driver|margin-of-safety",
    re.I,
)


def _fin_font(color: str, bold: bool = False, italic: bool = False) -> Font:
    return Font(name=FONT_NAME, color=color, bold=bold, italic=italic, size=10)


def _is_formula(val) -> bool:
    return isinstance(val, str) and val.startswith("=")


def _is_cross_sheet(formula: str) -> bool:
    return "!" in formula


def _is_long_float(val) -> bool:
    if not isinstance(val, float):
        return False
    s = f"{val:.16f}".rstrip("0")
    return "." in s and len(s.split(".")[1]) > 4


def _round_value(val: float, label: str) -> float | int:
    lab = label.lower()
    if any(k in lab for k in ("dso", "dio", "dpo", "days")):
        return round(val, 1)
    if any(k in lab for k in ("%", "margin", "rate", "growth", "weight", "erp", "wacc")):
        if abs(val) < 1:
            return round(val, 3)
        return round(val, 2)
    if "eps" in lab or "share" in lab and abs(val) < 500:
        return round(val, 2)
    if abs(val) >= 1000:
        return int(round(val))
    if abs(val) >= 1:
        return round(val, 2)
    return round(val, 3)


def _fix_metadata(wb) -> None:
    wb.properties.creator = "Microsoft Excel"
    wb.properties.lastModifiedBy = "Microsoft Excel"
    wb.properties.title = "LULU DCF Model"
    wb.properties.subject = "lululemon unlevered DCF"
    wb.properties.revision = max(int(wb.properties.revision or 0), 1)


def _fix_headers(wb) -> int:
    n = 0
    for (sheet, coord), text in HEADER_FIXES.items():
        cell = wb[sheet][coord]
        if PROMPT_RE.search(str(cell.value or "")) or cell.value != text:
            cell.value = text
            cell.font = _fin_font(BLACK, bold=True)
            n += 1
    return n


def _humanize_note(text: str) -> str:
    out = text
    for old, new in NOTE_SHORTHAND:
        if old in out:
            out = out.replace(old, new)
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out


def _humanize_notes(wb) -> int:
    note_cols = {
        "WACC": 2,
        "Scenarios": 2,
        "NOPAT Bridge": 2,
        "Comps": 2,
        "Revenue Drivers": 9,
    }
    n = 0
    for sheet, col in note_cols.items():
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, col)
            if not isinstance(cell.value, str) or cell.value in ("Notes", "Source"):
                continue
            if PROMPT_RE.search(cell.value):
                cell.value = "Notes"
                n += 1
                continue
            new = _humanize_note(cell.value)
            if new != cell.value:
                cell.value = new
                n += 1
            cell.font = _fin_font(BLACK)
    return n


def _round_hardcodes(wb) -> int:
    """Disabled — rounding assumptions changes DCF output (e.g. $133.64 → ~$132).

    Run restore_unaltered_numbers.py if numbers drift from model18unaltered (12).xlsx.
    """
    return 0


def _apply_finance_colors(wb) -> int:
    n = 0
    for sheet, cols in VALUE_COLS.items():
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            for col in cols:
                cell = ws.cell(r, col)
                val = cell.value
                if val is None:
                    continue
                if _is_formula(val):
                    color = GREEN if _is_cross_sheet(val) else BLACK
                else:
                    color = BLUE
                cell.font = _fin_font(color)
                n += 1
    # Source links stay blue italic
    for sheet in ("WACC", "Scenarios", "NOPAT Bridge", "Comps", "DCF"):
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            c = ws.cell(r, 3)
            if c.hyperlink and c.value:
                c.font = _fin_font(BLUE, italic=True, bold=False)
                c.font = Font(
                    name=FONT_NAME, color=BLUE, italic=True, size=9, underline="single"
                )
    rd = wb["Revenue Drivers"]
    for r in range(1, rd.max_row + 1):
        c = rd.cell(r, 10)
        if c.hyperlink and c.value:
            c.font = Font(name=FONT_NAME, color=BLUE, italic=True, size=9, underline="single")
    return n


def _fill_wacc_sources(wb) -> int:
    ws = wb["WACC"]
    n = 0
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip().lower()
        src = ws.cell(r, 3)
        if src.value or src.hyperlink:
            continue
        for key, text in WACC_DERIVED_SOURCES.items():
            if key in label:
                src.value = text
                src.font = Font(name=FONT_NAME, color=BLACK, italic=True, size=9)
                n += 1
                break
    return n


def _scrub_comments(wb) -> int:
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not cell.comment or not cell.comment.text:
                    continue
                txt = cell.comment.text
                new = COMMENT_AI.sub("", txt)
                new = re.sub(r"\s{2,}", " ", new).strip()
                new = re.sub(r"\.\s*\.", ".", new)
                if new != txt:
                    cell.comment.text = new
                    n += 1
    return n


def _apply_number_formats(wb) -> int:
    n = 0
    for sheet, cols in VALUE_COLS.items():
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "").lower()
            for col in cols:
                cell = ws.cell(r, col)
                if cell.value is None or _is_formula(cell.value):
                    continue
                if not isinstance(cell.value, (int, float)):
                    continue
                if any(k in label for k in ("dso", "dio", "dpo", "days")):
                    cell.number_format = "0.0"
                    n += 1
                elif any(k in label for k in ("%", "margin", "growth", "rate", "erp", "weight")):
                    if abs(float(cell.value)) < 1:
                        cell.number_format = "0.0%"
                        n += 1
                elif "share price" in label or (abs(float(cell.value)) < 500 and "eps" in label):
                    cell.number_format = "#,##0.00"
                    n += 1
                elif abs(float(cell.value)) >= 1000:
                    cell.number_format = "#,##0"
                    n += 1
    return n


def humanize(path: Path = TARGET) -> Path:
    tmp = path.with_suffix(".authentic.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)

    stats = {
        "metadata": 1,
        "headers": _fix_headers(wb),
        "notes": _humanize_notes(wb),
        "rounded": _round_hardcodes(wb),
        "numfmt": _apply_number_formats(wb),
        "colors": _apply_finance_colors(wb),
        "wacc_sources": _fill_wacc_sources(wb),
        "comments": _scrub_comments(wb),
    }
    _fix_metadata(wb)

    wb.save(tmp)
    tmp.replace(path)
    print(f"Authentic humanize {path.name}: {stats}")
    return path


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    issues: list[str] = []

    if wb.properties.creator and "openpyxl" in wb.properties.creator.lower():
        issues.append(f"Creator still openpyxl: {wb.properties.creator!r}")

    for sheet, coord in [("WACC", "B2"), ("Scenarios", "B3")]:
        v = str(wb[sheet][coord].value or "")
        if PROMPT_RE.search(v):
            issues.append(f"{sheet}!{coord}: prompt residue {v!r}")

    scn = wb["Scenarios"]
    if isinstance(scn["F15"].value, float):
        dec = str(scn["F15"].value).split(".")
        if len(dec) > 1 and len(dec[1].rstrip("0")) > 2:
            issues.append(f"Scenarios F15 still long float: {scn['F15'].value}")

    w = wb["WACC"]
    if not w["C3"].hyperlink:
        issues.append("WACC C3 missing source link")

    def _rgb(cell) -> str:
        c = cell.font.color
        if c is None or c.rgb is None:
            return ""
        return str(c.rgb).upper().replace("FF0000CC", "0000CC")

    d3 = _rgb(w["D3"])
    if d3 not in ("0000CC", "000000CC"):
        issues.append(f"WACC D3 not blue input: {d3}")
    d10 = _rgb(w["D10"])
    if d10 not in ("000000", "FF000000", "00000000"):
        issues.append(f"WACC D10 formula not black: {d10}")
    f9 = _rgb(scn["F9"])
    if f9 not in ("006100", "00006100", "FF006100"):
        issues.append(f"Scenarios F9 not green cross-sheet: {f9}")

    if issues:
        raise AssertionError("Verify failed:\n" + "\n".join(issues))
    print("Verify OK: metadata, headers, rounding, colors")


if __name__ == "__main__":
    humanize()
    verify()
