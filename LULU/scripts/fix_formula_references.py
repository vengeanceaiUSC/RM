#!/usr/bin/env python3
"""Replace remaining hardcoded FY25 revenue anchors in formula strings.

Substitutes /11102600 and *11102600 with /DCF!$D$5 and *DCF!$D$5 so
formulas link to the FY25A net revenue cell (same numeric value).

Run:  cd LULU && python3 scripts/fix_formula_references.py
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"
REV_ANCHOR = "DCF!$D$5"

# Exact full-formula replacements (run first)
EXACT: list[tuple[str, str]] = [
    ("=11102600*(1+E4)", f"={REV_ANCHOR}*(1+E4)"),
    ("=11102600*(1+F4)", f"={REV_ANCHOR}*(1+F4)"),
    ("=11102600*(1+G4)", f"={REV_ANCHOR}*(1+G4)"),
    ("=2210615+496228", "=DCF!D11+496228"),
]


def _rewrite(formula: str) -> str:
    if "11102600" not in formula:
        return formula
    out = formula.replace("/11102600", f"/{REV_ANCHOR}")
    out = out.replace("*11102600", f"*{REV_ANCHOR}")
    out = out.replace("(11102600", f"({REV_ANCHOR}")
    return out


def fix(path: Path = TARGET) -> list[str]:
    tmp = path.with_suffix(".formula_refs.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    changes: list[str] = []

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                new = v
                for old, repl in EXACT:
                    if new == old:
                        new = repl
                        break
                else:
                    new = _rewrite(v)
                if new != v:
                    changes.append(f"{ws.title}!{cell.coordinate}")
                    cell.value = new

    wb.save(tmp)
    tmp.replace(path)
    return changes


if __name__ == "__main__":
    changed = fix()
    print(f"Updated {len(changed)} formula references")
    for c in changed[:20]:
        print(f"  {c}")
    if len(changed) > 20:
        print(f"  … +{len(changed) - 20} more")
