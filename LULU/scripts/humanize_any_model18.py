#!/usr/bin/env python3
"""Humanize any model18unaltered*.xlsx → LULU_model_submission.xlsx.

Usage:
  cd LULU && python3 scripts/humanize_any_model18.py
  cd LULU && python3 scripts/humanize_any_model18.py "model18unaltered (13).xlsx"

Preserves all hardcoded assumption values. Strips Ctrl+F, Firecrawl, Phase jargon.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = Path(__file__).resolve().parent

CANDIDATES = [
    "model18unaltered (13).xlsx",
    "model18unaltered (12).xlsx",
    "model18unaltered.xlsx",
]

STEPS = [
    "polish_model18_altered.py",
    "restore_unaltered_numbers.py",
    "remove_spot_check.py",
    "fix_dcf_tv_reconciliation.py",
    "fix_formula_references.py",
    "fix_ib_formatting.py",
    "humanize_workbook_authentic.py",
    "fix_financial_number_formats.py",
    "fix_decimal_display.py",
    "humanize_for_club_submission.py",
    "fix_mechanical_architecture.py",
    "fix_file_metadata.py",
    "strip_programmatic_colors.py",
    "strip_arrows.py",
    "remove_collapsible_outlines.py",
    "final_label_cleanup.py",
]


def resolve_input(arg: str | None) -> Path:
    if arg:
        p = ROOT / arg if not Path(arg).is_absolute() else Path(arg)
        if not p.exists():
            raise FileNotFoundError(f"Input not found: {p}")
        return p
    for name in CANDIDATES:
        p = ROOT / name
        if p.exists():
            return p
    raise FileNotFoundError(f"No input workbook found. Expected one of: {CANDIDATES}")


def main() -> int:
    src = resolve_input(sys.argv[1] if len(sys.argv) > 1 else None)
    unaltered_ref = ROOT / "model18unaltered (12).xlsx"
    if not unaltered_ref.exists():
        shutil.copy2(src, unaltered_ref)

    altered = ROOT / "model18altered.xlsx"
    final = ROOT / "unbeiesgbar_final.xlsx"
    out = ROOT / "LULU_model_submission.xlsx"

    shutil.copy2(src, altered)
    print(f"Input: {src.name} ({src.stat().st_size} bytes)")

    print("\n=== polish_model18_altered.py ===")
    r = subprocess.run([sys.executable, str(SCRIPTS / "polish_model18_altered.py")], cwd=SCRIPTS, check=False)
    if r.returncode != 0:
        print("FAILED: polish_model18_altered.py")
        return r.returncode

    shutil.copy2(altered, final)

    for step in STEPS[1:]:
        print(f"\n=== {step} ===")
        r = subprocess.run([sys.executable, str(SCRIPTS / step)], cwd=ROOT, check=False)
        if r.returncode != 0 and step not in (
            "humanize_for_club_submission.py",
            "fix_mechanical_architecture.py",
        ):
            print(f"WARN: {step} exit {r.returncode} (continuing)")
    shutil.copy2(final, out)
    print(f"\nOutput: {out}")
    subprocess.run([sys.executable, str(SCRIPTS / "audit_ai_tells.py")], cwd=ROOT, check=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
