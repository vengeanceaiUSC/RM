#!/usr/bin/env python3
"""Fix DCF terminal-value reconciliation block — point at column B outputs (not empty D refs).

Rows 90–95 compare Gordon growth vs exit-multiple method. Outputs live in column B
after the Notes-column delete (B46 TV, B47 multiple, B57/B87 share prices, B84 exit TV).

Run:  cd LULU && python3 scripts/fix_dcf_tv_reconciliation.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# Row 89 = section title (col A only). Headers + data rows 90–95.
FIXES: dict[str, str | None] = {
    "B82": None,
    "A90": None,
    "B90": "Gordon Growth",
    "C90": "Exit Multiple",
    "D90": "Variance",
    "A91": "Terminal value ($)",
    "B91": "=B46",
    "C91": "=B84",
    "D91": "=B91-C91",
    "A92": "Exit EV/EBITDA (implied vs selected)",
    "B92": "=B47",
    "C92": "=B83",
    "D92": "=B92-C92",
    "A93": "Terminal value variance (%)",
    "B93": None,
    "C93": None,
    "D93": "=(B91-C91)/B91",
    "A94": "Implied share price",
    "B94": "=B57",
    "C94": "=B87",
    "D94": "=B94-C94",
    "A95": "Exit multiple check (±1.5x)",
    "B95": '=IF(ABS(D92)<=1.5,"PASS","REVIEW")',
    "C95": '=TEXT(D92,"0.0")&"x spread vs Gordon-implied exit"',
    "D95": "Selected exit is the Gordon identity, so spread should be 0",
    "B84": "=B45*B83",
    "B87": "=(B86+B51+B52)/B56",
}


def fix(path: Path = TARGET) -> list[str]:
    tmp = path.with_suffix(".tv_fix.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    ws = wb["DCF"]
    changes: list[str] = []

    for coord, val in FIXES.items():
        cell = ws[coord]
        if cell.value != val:
            changes.append(f"{coord}: {cell.value!r} → {val!r}")
            cell.value = val

    wb.save(tmp)
    tmp.replace(path)
    return changes


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["DCF"]
    issues: list[str] = []

    for coord, expected in [
        ("B91", "=B46"),
        ("C91", "=B84"),
        ("D91", "=B91-C91"),
        ("B94", "=B57"),
        ("C94", "=B87"),
        ("B87", "=(B86+B51+B52)/B56"),
        ("B84", "=B45*B83"),
    ]:
        if ws[coord].value != expected:
            issues.append(f"{coord} expected {expected!r}, got {ws[coord].value!r}")

    if ws["B90"].value != "Gordon Growth":
        issues.append(f"B90 header wrong: {ws['B90'].value!r}")

    if issues:
        raise AssertionError("Verify failed:\n" + "\n".join(issues))
    print("Verify OK: TV reconciliation refs point to column B")


if __name__ == "__main__":
    changed = fix()
    print(f"Fixed {len(changed)} cells in {TARGET.name}")
    for c in changed:
        print(f"  {c}")
    verify()
