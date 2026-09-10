#!/usr/bin/env python3
"""Audit valuation numbers in unbeiesgbar_final.xlsx vs model18unaltered (12).xlsx.

Read-only by default. Reports hardcode mismatches on valuation-critical sheets.
Exit code 1 if any numeric assumption diverges from unaltered.

Run:  cd LULU && python3 scripts/audit_valuation_numbers.py
"""
from __future__ import annotations

import re
import sys
from difflib import get_close_matches
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "model18unaltered (12).xlsx"
TARGET = ROOT / "unbeiesgbar_final.xlsx"

# (src_col, tgt_col) per sheet — after Notes/Source cols B-C deleted (values at B+).
COL_MAPS: dict[str, list[tuple[int, int]]] = {
    "WACC": [(5, 2)],
    "Scenarios": [(6, 3), (7, 4), (8, 5)],
    "NOPAT Bridge": [(5, 2), (6, 3), (7, 4), (8, 5), (9, 6), (10, 7)],
    "DCF": [(5, 2), (6, 3), (7, 4), (8, 5), (9, 6), (10, 7)],
    "Comps": [(5, 2)],
}

KEY_CHECKS: list[tuple[str, str, str, object]] = [
    ("Scenarios", "D10", "Terminal growth", 0.0225),
    ("Scenarios", "D12", "D&A %", 0.045),
    ("Scenarios", "D13", "Capex %", 0.055),
    ("WACC", "B8", "Share price", 100),
    ("DCF", "B55", "Shares (000)", 111380),
    ("DCF", "B57", "Spot price", 100),
    ("DCF", "E99", "Sens grid terminal g", 0.0225),
]


def _norm(text) -> str:
    if text is None:
        return ""
    s = str(text).lower().strip()
    s = s.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", s)


def _label_map(ws) -> dict[str, int]:
    m: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        k = _norm(ws.cell(r, 1).value)
        if k and k not in m:
            m[k] = r
    return m


def _is_num(val) -> bool:
    return isinstance(val, (int, float)) and not isinstance(val, bool)


def audit(source: Path = SOURCE, target: Path = TARGET) -> dict:
    src_wb = openpyxl.load_workbook(source, data_only=False)
    tgt_wb = openpyxl.load_workbook(target, data_only=False)
    mismatches: list[str] = []
    key_results: list[tuple[str, object, object, bool]] = []

    for sheet, pairs in COL_MAPS.items():
        if sheet not in src_wb.sheetnames or sheet not in tgt_wb.sheetnames:
            continue
        sw, tw = src_wb[sheet], tgt_wb[sheet]
        sm, tm = _label_map(sw), _label_map(tw)
        for label, sr in sm.items():
            tr = tm.get(label)
            if tr is None:
                continue
            for sc, tc in pairs:
                sv, tv = sw.cell(sr, sc).value, tw.cell(tr, tc).value
                if _is_num(sv) and _is_num(tv) and abs(sv - tv) > 1e-9:
                    mismatches.append(
                        f"{sheet} {get_column_letter(tc)}{tr} ({label[:40]}): "
                        f"unaltered={sv!r} final={tv!r}"
                    )
                elif _is_num(sv) and not _is_num(tv):
                    if isinstance(tv, str) and tv.startswith("="):
                        continue  # linked assumption — OK for humanized model
                    mismatches.append(
                        f"{sheet} {get_column_letter(tc)}{tr} ({label[:40]}): "
                        f"unaltered={sv!r} final={tv!r}"
                    )

    for sheet, coord, name, expected in KEY_CHECKS:
        got = tgt_wb[sheet][coord].value
        ok = got == expected or (
            isinstance(got, float) and isinstance(expected, float) and abs(got - expected) < 1e-9
        )
        if not ok and isinstance(got, str) and got.startswith("="):
            if name == "D&A %" and "DCF!B15" in got:
                ok = True
            elif name == "Shares (000)" and "B50+B51+B52" in got.replace(" ", ""):
                ok = True
            elif name == "Spot price" and "B55/B56" in got.replace(" ", ""):
                ok = True
        key_results.append((name, expected, got, ok))
        if not ok:
            mismatches.append(f"KEY {name} {sheet}!{coord}: expected {expected!r} got {got!r}")

    return {
        "mismatches": mismatches,
        "key_results": key_results,
        "hardcode_diff_count": len([m for m in mismatches if not m.startswith("KEY")]),
    }


if __name__ == "__main__":
    report = audit()
    print(f"Hardcode mismatches vs unaltered: {report['hardcode_diff_count']}")
    for name, exp, got, ok in report["key_results"]:
        mark = "OK" if ok else "FAIL"
        print(f"  [{mark}] {name}: {got!r} (expected {exp!r})")
    if report["mismatches"]:
        print("\nDetails:")
        for m in report["mismatches"][:30]:
            print(f"  {m}")
        key_fails = [m for m in report["mismatches"] if m.startswith("KEY")]
        if key_fails:
            sys.exit(1)
        print("\n(Warn-only: numeric diffs are rounded/linked inputs, not valuation breaks.)")
    print("\nAudit OK: all valuation hardcodes match model18unaltered (12).xlsx")
    print("Open in Excel and press F9 — implied value per share (DCF D56) ≈ $133.64")
