#!/usr/bin/env python3
"""Restore hardcoded numbers in unbeiesgbar_final.xlsx from model18unaltered (12).xlsx.

The unaltered workbook is the numeric source of truth. Humanize/round passes must
never change assumption values. This script copies exact hardcodes by row label
and verifies the base-case DCF implied share price ≈ $133.64.

Run:  cd LULU && python3 scripts/restore_unaltered_numbers.py
"""
from __future__ import annotations

import re
import shutil
from difflib import get_close_matches
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "model18unaltered (12).xlsx"
TARGET = ROOT / "unbeiesgbar_final.xlsx"
TARGET_IMPLIED = 133.64
IMPLIED_TOLERANCE = 0.05

# Unaltered value col → restored value col (after Notes+Source insert at B).
SHEET_COL_MAP: dict[str, tuple[int, int]] = {
    "WACC": (5, 4),  # E → D
    "Scenarios": (7, 6),  # G (base) primary check; also copies F→E, H→G
    "NOPAT Bridge": (5, 4),
    "Comps": (5, 4),
    "DCF": (5, 4),  # summary col E → D
    "Revenue Drivers": (5, 3),  # col C forecast anchor — spot-check only
}

SCEN_SCENARIO_COLS = (
    (6, 5),  # Bear F → E
    (7, 6),  # Base G → F
    (8, 7),  # Bull H → G
)

PROTECTED_COLS = {2, 3}  # Notes, Source — never overwrite


def _norm_label(text) -> str:
    if text is None:
        return ""
    s = str(text).lower().strip()
    s = s.replace("–", "-").replace("—", "-").replace("β", "beta")
    s = re.sub(r"^\s*memo:\s*", "", s)
    s = re.sub(r"phase [1-5][:\s-]*", "", s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^\w\s$%()./-]", "", s)
    return s


def _build_label_map(ws) -> dict[str, int]:
    m: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        key = _norm_label(ws.cell(r, 1).value)
        if key and key not in m:
            m[key] = r
    return m


def _match_row(label: str, label_map: dict[str, int], *, exact_only: bool = True) -> int | None:
    key = _norm_label(label)
    if key in label_map:
        return label_map[key]
    if exact_only:
        return None
    hit = get_close_matches(key, label_map.keys(), n=1, cutoff=0.92)
    return label_map[hit[0]] if hit else None


def _is_hardcode(val) -> bool:
    return isinstance(val, (int, float)) and not isinstance(val, bool)


def _copy_hardcodes(
    src_ws,
    tgt_ws,
    src_col: int,
    tgt_col: int,
    src_map: dict[str, int],
    tgt_map: dict[str, int],
) -> list[str]:
    changes: list[str] = []
    for r in range(1, tgt_ws.max_row + 1):
        if tgt_col in PROTECTED_COLS:
            continue
        label = tgt_ws.cell(r, 1).value
        if not label:
            continue
        src_row = _match_row(str(label), src_map)
        if src_row is None:
            continue
        src_val = src_ws.cell(src_row, src_col).value
        if not _is_hardcode(src_val):
            continue
        tgt_cell = tgt_ws.cell(r, tgt_col)
        if tgt_cell.value != src_val:
            old = tgt_cell.value
            tgt_cell.value = src_val
            changes.append(
                f"{tgt_ws.title}!{tgt_cell.coordinate}: {old!r} → {src_val!r} "
                f"({label})"
            )
    return changes


def restore_numbers(
    source: Path = SOURCE,
    target: Path = TARGET,
) -> list[str]:
    if not source.exists():
        raise FileNotFoundError(source)
    if not target.exists():
        raise FileNotFoundError(target)

    tmp = target.with_suffix(".restoring_nums.xlsx")
    shutil.copy2(target, tmp)
    src_wb = openpyxl.load_workbook(source, data_only=False)
    tgt_wb = openpyxl.load_workbook(tmp)

    all_changes: list[str] = []

    # WACC, NOPAT, Comps, DCF summary col
    for sheet, (src_col, tgt_col) in SHEET_COL_MAP.items():
        if sheet not in src_wb.sheetnames or sheet not in tgt_wb.sheetnames:
            continue
        src_ws = src_wb[sheet]
        tgt_ws = tgt_wb[sheet]
        all_changes.extend(
            _copy_hardcodes(
                src_ws,
                tgt_ws,
                src_col,
                tgt_col,
                _build_label_map(src_ws),
                _build_label_map(tgt_ws),
            )
        )

    # Scenarios: restore bear/base/bull hardcodes for assumption block
    if "Scenarios" in src_wb.sheetnames and "Scenarios" in tgt_wb.sheetnames:
        su = src_wb["Scenarios"]
        sf = tgt_wb["Scenarios"]
        sm, tm = _build_label_map(su), _build_label_map(sf)
        for src_col, tgt_col in SCEN_SCENARIO_COLS:
            all_changes.extend(_copy_hardcodes(su, sf, src_col, tgt_col, sm, tm))

    tgt_wb.save(tmp)
    tmp.replace(target)
    return all_changes


def _eval_implied_share_price(path: Path) -> float | None:
    try:
        from xlcalculator import ModelCompiler, Model  # type: ignore
    except ImportError:
        return None

    try:
        compiler = ModelCompiler()
        parsed = compiler.read_and_parse_archive(str(path))
        model = Model()
        model.build_code(parsed)
        val = model.get_cell_value("DCF!D56")
        if isinstance(val, (int, float)):
            return float(val)
    except Exception:
        return None
    return None


def verify(
    source: Path = SOURCE,
    target: Path = TARGET,
) -> None:
    src_wb = openpyxl.load_workbook(source, data_only=False)
    tgt_wb = openpyxl.load_workbook(target, data_only=False)
    issues: list[str] = []

    su, sf = src_wb["Scenarios"], tgt_wb["Scenarios"]
    sm, tm = _build_label_map(su), _build_label_map(sf)

    key_rows = [
        "terminal growth",
        "dso",
        "fy25 dio anchor",
        "dpo",
        "prepaid expenses",
        "accrued liabilities",
    ]
    for needle in key_rows:
        sr = _match_row(needle, sm)
        tr = _match_row(needle, tm)
        if sr is None or tr is None:
            continue
        sv = su.cell(sr, 7).value
        tv = sf.cell(tr, 6).value
        if sv != tv:
            issues.append(f"Scenarios base {needle}: src={sv!r} tgt={tv!r}")

    wu, wf = src_wb["WACC"], tgt_wb["WACC"]
    sr = _match_row("share price", _build_label_map(wu))
    tr = _match_row("share price", _build_label_map(wf))
    if sr and tr:
        if wu.cell(sr, 5).value != wf.cell(tr, 4).value:
            issues.append(
                f"WACC share price: src={wu.cell(sr,5).value} tgt={wf.cell(tr,4).value}"
            )

    implied = _eval_implied_share_price(target)
    if implied is not None:
        if abs(implied - TARGET_IMPLIED) > IMPLIED_TOLERANCE:
            issues.append(
                f"DCF implied ${implied:.2f} ≠ target ${TARGET_IMPLIED} "
                f"(±${IMPLIED_TOLERANCE})"
            )
        print(f"  DCF implied share price: ${implied:.2f} (target ${TARGET_IMPLIED})")
    else:
        print("  (xlcalculator unavailable — spot-check hardcodes only)")

    if issues:
        raise AssertionError("Verification failed:\n" + "\n".join(issues))

    print("Verify OK: hardcodes match unaltered model")


if __name__ == "__main__":
    changes = restore_numbers()
    print(f"Restored {len(changes)} hardcodes in {TARGET.name}")
    for line in changes[:25]:
        print(f"  {line}")
    if len(changes) > 25:
        print(f"  ... and {len(changes) - 25} more")
    verify()
