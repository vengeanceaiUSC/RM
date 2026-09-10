#!/usr/bin/env python3
"""Refactor model formulas to Wall Street Prep / Macabacus analyst standards.

Run: cd LULU && python3 scripts/refactor_formulas_wsp.py
"""
from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "model18_humanized (1).xlsx"
OUTPUT = ROOT / "model18_wsp_formulas.xlsx"
LOG_CSV = ROOT / "formula_refactor_log.csv"

FY25_REV = "DCF!$E$5"
FY25_NWC = "DCF!$E$32"
FY25_CASH = "DCF!$E$50"
FY25_NET_DEBT = "DCF!$E$51"
FY25_SHARES = "DCF!$E$55"

ANCHOR_LABELS: list[tuple[int, str, float | None]] = [
    (194, "FY25 total assets ($000)", 8456743),
    (195, "FY25 total liabilities ($000)", 3494903),
    (196, "FY25 total equity ($000)", 4961840),
    (197, "FY25 diluted shares (000)", 119068),
    (202, "FY25 stock-based compensation ($000)", 62203),
    (203, "FY25 implied lease interest ($000)", 1028),
]

LITERAL_TO_ANCHOR: list[tuple[str, str, str]] = [
    ("+1807202-1798441", f"+{FY25_CASH}+{FY25_NET_DEBT}", "Equity bridge: cash plus signed net debt from DCF FY25A"),
    ("/111380", f"/{FY25_SHARES}", "Link per-share bridge to DCF FY25A diluted shares"),
    ("11102600", FY25_REV, "Link FY26 revenue roll-forward to DCF FY25A net revenue"),
    ("1461096", FY25_NWC, "Link NWC bridge to DCF FY25A NWC balance"),
    ("1807202", FY25_CASH, "Link bridge cash to DCF FY25A cash"),
    ("8456743", "Scenarios!$B$194", "Scale balance sheet off FY25 total assets driver"),
    ("3494903", "Scenarios!$B$195", "Scale balance sheet off FY25 total liabilities driver"),
    ("4961840", "Scenarios!$B$196", "Scale balance sheet off FY25 total equity driver"),
    ("119068", "Scenarios!$B$197", "Link diluted share count to FY25 shares driver"),
    ("62203", "Scenarios!$B$202", "Link SBC add-back to FY25 SBC driver"),
    ("1028", "Scenarios!$B$203", "Link lease interest reclass to FY25 lease interest driver"),
]

TEXT_FORMULA_RE = re.compile(
    r"formula|calculated via|cross-check|black font\s*=\s*calculations|"
    r"Revenue Drivers:.*formula|beginning-stores formula|e-comm formula|store-channel formula",
    re.I,
)

FUNCTION_RE = re.compile(
    r"(?<![A-Za-z0-9_])(" + "|".join(
        [
            "sum", "if", "max", "min", "abs", "round", "iferror", "choose", "npv", "irr",
            "offset", "index", "match", "xlookup", "text", "and", "or", "not",
        ]
    )
    + r")\s*\(",
    re.I,
)


def uppercase_functions(formula: str) -> str:
    return FUNCTION_RE.sub(lambda m: m.group(1).upper() + "(", formula)


def strip_round_wrappers(formula: str) -> tuple[str, str | None]:
    """Remove ROUND(...,0) wrappers from core schedules."""
    reason = None
    m = re.fullmatch(r"=(-?)ROUND\((.+),0\)", formula, re.I)
    if m:
        sign, inner = m.group(1), m.group(2)
        reason = "Remove ROUND wrapper; expose raw repurchase math for audit"
        return f"={sign}{inner}", reason
    new = formula
    for pat, repl, why in [
        (r"ROUND\(([^,]+),0\)", r"\1", "Remove ROUND wrapper from repurchase/share schedule"),
    ]:
        if re.search(pat, new, re.I):
            new = re.sub(pat, repl, new, flags=re.I)
            reason = why
    return new, reason


def remove_same_sheet_prefix(formula: str, sheet: str) -> tuple[str, str | None]:
    norm_sheet = sheet.strip().lower()

    def repl(m: re.Match) -> str:
        quoted, plain = m.group(1), m.group(2)
        name = (quoted or plain).strip()
        if name.lower() != norm_sheet:
            return m.group(0)
        row = m.group(6)
        d1, col, d2 = m.group(3), m.group(4), m.group(5)
        return f"{d1}{col}{d2}{row}"

    pat = re.compile(
        r"(?:'([^']+)'|([A-Za-z][A-Za-z0-9_ ]*))!(\$?)([A-Z]{1,3})(\$?)(\d+)"
    )
    new = pat.sub(repl, formula)
    if new != formula:
        return new, f"Remove self-referencing '{sheet}!' prefix on active tab"
    return formula, None


def replace_literals(formula: str) -> tuple[str, list[str]]:
    reasons: list[str] = []
    out = formula
    for literal, anchor, reason in LITERAL_TO_ANCHOR:
        if literal in out:
            out = out.replace(literal, anchor)
            reasons.append(reason)
    return out, reasons


def refactor_sensitivity(dcf, changes: list[dict]) -> None:
    """Rewrite DCF sensitivity grid to reference WACC row and terminal-g row."""
    for row in range(99, 104):
        wacc_cell = dcf[f"A{row}"].value
        if not isinstance(wacc_cell, (int, float)):
            continue
        for col_idx in range(3, 9):  # C-H
            col = openpyxl.utils.get_column_letter(col_idx)
            cell = dcf[f"{col}{row}"]
            old = cell.value
            if not isinstance(old, str) or "NPV(" not in old.upper():
                continue
            new = (
                f"=(NPV($A{row},$F$35:$J$35)+($J$35*(1+{col}$98)/($A{row}-{col}$98))"
                f"/(1+$A{row})^$E$36+$E$50+$E$51)/$E$55"
            )
            if new != old:
                cell.value = new
                changes.append(
                    {
                        "sheet": "DCF",
                        "cell": f"{col}{row}",
                        "original": old,
                        "humanized": new,
                        "reason": "Link sensitivity NPV/Gordon to WACC row and terminal-g header row",
                    }
                )


def refactor_comps(comps, changes: list[dict]) -> None:
    cell = comps["E5"]
    old = cell.value
    if old == "=2210615+496228":
        new = "=DCF!$E$11+DCF!$E$16"
        cell.value = new
        changes.append(
            {
                "sheet": "Comps",
                "cell": "E5",
                "original": old,
                "humanized": new,
                "reason": "Replace inline EBIT+D&A literals with DCF FY25A links",
            }
        )


def refactor_revenue_drivers(rd, changes: list[dict]) -> None:
    rd["A3"] = "Unit conversions (drivers)"
    rd["B3"] = 1000000
    rd["C3"] = 1000
    for col in "CDEFG":
        cell = rd[f"{col}41"]
        old = cell.value
        if not isinstance(old, str) or "1000000" not in old:
            continue
        new = f"={col}38*$B$3*{col}39*{col}40/$C$3"
        cell.value = new
        changes.append(
            {
                "sheet": "Revenue Drivers",
                "cell": f"{col}41",
                "original": old,
                "humanized": new,
                "reason": "Extract sessions/output unit scalars to explicit driver cells (B3/C3)",
            }
        )


def setup_anchors(scn) -> None:
    scn["A193"] = "FY25 anchors (drivers — do not edit without source)"
    for row, label, val in ANCHOR_LABELS:
        scn.cell(row, 1).value = label
        if val is not None:
            scn.cell(row, 2).value = val


def clear_formula_meta(wb, changes: list[dict]) -> None:
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and not val.startswith("=") and TEXT_FORMULA_RE.search(val):
                    changes.append(
                        {
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "original": val,
                            "humanized": "",
                            "reason": "Delete formula/meta explanation text cell",
                        }
                    )
                    cell.value = None


def refactor_workbook(src: Path = INPUT, dst: Path = OUTPUT) -> list[dict]:
    tmp = dst.with_suffix(".tmp.xlsx")
    shutil.copy2(src, tmp)
    wb = openpyxl.load_workbook(tmp)
    changes: list[dict] = []

    setup_anchors(wb["Scenarios"])
    refactor_revenue_drivers(wb["Revenue Drivers"], changes)
    refactor_comps(wb["Comps"], changes)
    refactor_sensitivity(wb["DCF"], changes)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not isinstance(val, str) or not val.startswith("="):
                    continue
                original = val
                new = val
                reasons: list[str] = []

                up = uppercase_functions(new)
                if up != new:
                    new = up
                    reasons.append("Convert Excel functions to uppercase")

                no_self, r_self = remove_same_sheet_prefix(new, ws.title)
                if r_self:
                    new = no_self
                    reasons.append(r_self)

                no_round, r_round = strip_round_wrappers(new)
                if r_round:
                    new = no_round
                    reasons.append(r_round)

                new_lit, lit_reasons = replace_literals(new)
                if lit_reasons:
                    new = new_lit
                    reasons.extend(lit_reasons)

                if new != original:
                    cell.value = new
                    changes.append(
                        {
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "original": original,
                            "humanized": new,
                            "reason": "; ".join(dict.fromkeys(reasons)),
                        }
                    )

    clear_formula_meta(wb, changes)

    from restore_outline_groups import restore_outline_groups  # noqa: E402

    restore_outline_groups(wb)

    # Fix invalid label formulas and broken internal-link text before save.
    import re
    from openpyxl.worksheet.hyperlink import Hyperlink

    label_formula = re.compile(r"^=\s+(.+)$")
    if "NOPAT Bridge" in wb.sheetnames:
        ws = wb["NOPAT Bridge"]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and (m := label_formula.match(v)):
                    cell.value = m.group(1).strip()
    for sheet, coord, text, location in [
        ("DCF", "C4", "NOPAT Bridge tab", "'NOPAT Bridge'!A1"),
        ("DCF", "D4", "Revenue Drivers tab", "'Revenue Drivers'!A1"),
        ("Revenue Drivers", "K56", "Scenarios tab: base revenue", "'Scenarios'!G25"),
    ]:
        if sheet not in wb.sheetnames:
            continue
        cell = wb[sheet][coord]
        if isinstance(cell.value, str) and "!" in cell.value and not cell.value.startswith("="):
            cell.value = text
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location)

    wb.save(tmp)
    tmp.replace(dst)
    return changes


def write_log(changes: list[dict], path: Path = LOG_CSV) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            ["Sheet Name", "Cell Reference", "Original AI Formula", "Humanized WSP Formula", "Reason for Change"]
        )
        for c in changes:
            writer.writerow([c["sheet"], c["cell"], c["original"], c["humanized"], c["reason"]])


if __name__ == "__main__":
    changes = refactor_workbook()
    write_log(changes)
    print(f"Saved {OUTPUT}")
    print(f"Wrote {LOG_CSV} ({len(changes)} changes)")
