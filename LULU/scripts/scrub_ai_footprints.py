#!/usr/bin/env python3
"""Run full AI-footprint scrub pipeline on unbeiesgbar_final.xlsx.

Order matters. Preserves hardcoded values; run audit_valuation_numbers.py after.

Run:  cd LULU && python3 scripts/scrub_ai_footprints.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = Path(__file__).resolve().parent

STEPS = [
    "backfill_source_comments.py",
    "fix_decimal_display.py",
    "fix_human_analyst_structure.py",
    "fix_stale_formula_refs.py",
    "fix_div_errors.py",
    "humanize_for_submission.py",
    "trim_model_labels.py",
    "fix_file_metadata.py",
    "strip_programmatic_colors.py",
    "audit_valuation_numbers.py",
]

# fix_dcf_valuation_refs.py runs inside fix_div_errors.py

# Superseded by fix_human_analyst_structure.py (draggable drivers + local P&L):
# rebuild_dcf_local_formulas.py

# One-time structural step (already applied on current workbook):
# move_sources_to_comments.py — delete cols B-C; use backfill on re-runs


def main() -> int:
    for step in STEPS:
        print(f"\n=== {step} ===")
        r = subprocess.run(
            [sys.executable, str(SCRIPTS / step)],
            cwd=ROOT,
            check=False,
        )
        if r.returncode != 0:
            print(f"FAILED: {step} (exit {r.returncode})")
            return r.returncode
    print("\nScrub pipeline complete.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
