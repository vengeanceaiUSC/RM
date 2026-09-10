#!/usr/bin/env python3
"""Fix Excel +/- outline controls on model18 workbooks.

Run:  cd LULU && python3 scripts/fix_collapsible_outlines.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from restore_outline_groups import restore_outline_groups  # noqa: E402

ROOT = SCRIPTS.parent
TARGETS = [
    ROOT / "model18unaltered.xlsx",
    ROOT / "model18unaltered (12).xlsx",
    ROOT / "unbeiesgbar2model.xlsx",
    ROOT / "unbesiegbarmodel1.xlsx",
    ROOT / "model18altered.xlsx",
    ROOT / "model18_final.xlsx",
]


def fix_file(path: Path) -> None:
    wb = openpyxl.load_workbook(path)
    n = restore_outline_groups(wb)
    wb.save(path)
    print(f"Fixed {path.name} ({n} grouped sheets, showOutlineSymbols=True)")


if __name__ == "__main__":
    for p in TARGETS:
        if p.exists():
            fix_file(p)
        else:
            print(f"Skip missing: {p.name}")
