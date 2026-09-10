#!/usr/bin/env python3
"""Remove Excel +/- outline groups from model workbooks.

Run:  cd LULU && python3 scripts/remove_collapsible_outlines.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from restore_outline_groups import remove_outline_groups  # noqa: E402

ROOT = SCRIPTS.parent
TARGETS = [
    ROOT / "unbeiesgbar_final.xlsx",
    ROOT / "unbeiesgbar2model.xlsx",
]


def strip_file(path: Path) -> None:
    wb = openpyxl.load_workbook(path)
    n = remove_outline_groups(wb)
    wb.save(path)
    print(f"Removed outlines from {path.name} ({n} sheets cleared, showOutlineSymbols=False)")


if __name__ == "__main__":
    for p in TARGETS:
        if p.exists():
            strip_file(p)
        else:
            print(f"Skip missing: {p.name}")
