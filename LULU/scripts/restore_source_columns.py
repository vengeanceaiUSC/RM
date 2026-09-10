#!/usr/bin/env python3
"""Restore visible Notes (B) and Source links (C) on unbeiesgbar_final.xlsx.

Re-inserts the documentation columns removed by strip_model_manual.py, shifts
value/formula columns right, rebinds cross-sheet refs, and fills ~8-word Notes
(from rewrite_model_notes.NOTES_MAP) plus clickable Source hyperlinks from
model18unaltered (12).xlsx.

Run:  cd LULU && python3 scripts/restore_source_columns.py
"""
from __future__ import annotations

import re
import shutil
import sys
from difflib import get_close_matches
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

from polish_model18_altered import (  # noqa: E402
    _rewrite_all_formulas,
    _rewrite_sheet_internal_formulas,
)
from strip_model_manual import _exact_label_row, _label_row  # noqa: E402
from rewrite_model_notes import HEADER_CELLS, NOTES_MAP, _lines_ok, _word_count  # noqa: E402

ROOT = SCRIPTS.parent
ORIGINAL = ROOT / "model18unaltered (12).xlsx"
TARGET = ROOT / "unbeiesgbar_final.xlsx"

INSERT_BC_SHEETS = ("WACC", "Scenarios", "NOPAT Bridge", "Comps", "DCF")
SOURCE_HEADER = "Source"
NOTES_HEADER = "Notes"
BLUE = "0563C1"
CTRLF_RE = re.compile(r"Ctrl\+F[^\n]*", re.I)
URL_RE = re.compile(r"https?://[^\s\)\]\"']+", re.I)

# Original doc columns (B=justification, C=source) before strip.
ORIG_DOC_COLS = (2, 3)

# Revenue Drivers: append Notes/Source after data (no column shift).
RD_NOTES_COL = 9   # I
RD_SOURCE_COL = 10  # J
RD_HEADER_ROW = 5


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


def _col_shift_insert_at_b(max_col: int = 40) -> dict[str, str]:
    return {get_column_letter(c): get_column_letter(c + 2) for c in range(2, max_col + 1)}


def _truncate_to_words(text: str | None, max_words: int = 8) -> str | None:
    if not text:
        return None
    t = CTRLF_RE.sub("", str(text))
    t = re.sub(r"\s+", " ", t).strip()
    if not t:
        return None
    words = t.split()
    if len(words) <= max_words:
        return t
    return " ".join(words[:max_words]) + "."


def _note_for_row(sheet: str, label: str, orig_justification: str | None) -> str | None:
    sheet_map = NOTES_MAP.get(sheet, {})
    if label in sheet_map:
        note = sheet_map[label]
    elif orig_justification:
        note = _truncate_to_words(orig_justification)
    else:
        return None

    if note and "share price" in label.lower() and "~$100" in note:
        note = note.replace("~$100", "~$100.61")
    return note


def _clean_source_label(text: str | None) -> str | None:
    if not text:
        return None
    t = CTRLF_RE.sub("", str(text))
    t = re.sub(r"\s+", " ", t).strip()
    if not t or t.lower() in {"source (click)", "ctrl+f (prove number)"}:
        return None
    if "justification" in t.lower() and "click" in t.lower():
        return None
    # Keep first line for multi-line source labels.
    return t.split("\n")[0].strip() or None


def _rewrite_location(loc: str, sheet_shifts: dict[str, dict[str, str]]) -> str:
    """Legacy insert-shift rewriter (unused for original→restored links)."""
    m = re.match(
        r"^(?:'([^']+)'|([^!]+))!(\$?)([A-Z]{1,3})(\$?)(\d+)$",
        loc,
    )
    if not m:
        return loc
    sh_q, sh_p, d1, col, d2, row = m.groups()
    sh = sh_q or sh_p
    shift = sheet_shifts.get(sh, {})
    new_col = shift.get(col, col)
    prefix = f"'{sh}'" if sh_q else sh
    return f"{prefix}!{d1}{new_col}{d2}{row}"


def _value_col_for_sheet(sheet: str) -> str:
    """Primary hardcoded / linked value column after Notes+Source restore."""
    return "D"


def _resolve_internal_link(
    orig_wb, wb, location: str, sheet_shifts: dict[str, dict[str, str]]
) -> str:
    """Map original internal links to restored row/column coordinates."""
    m = re.match(
        r"^(?:'([^']+)'|([^!]+))!(\$?)([A-Z]{1,3})(\$?)(\d+)$",
        location,
    )
    if not m:
        return location

    sh_q, sh_p, d1, col, d2, row_s = m.groups()
    sh = sh_q or sh_p
    row = int(row_s)
    prefix = f"'{sh}'" if sh_q else sh

    if sh not in orig_wb.sheetnames or sh not in wb.sheetnames:
        return _rewrite_location(location, sheet_shifts)

    orig_ws = orig_wb[sh]
    restored_ws = wb[sh]
    label = orig_ws.cell(row, 1).value

    if isinstance(label, str) and label.strip():
        restored_row = _match_row(label.strip(), _build_label_map(restored_ws))
        if restored_row is None:
            restored_row = _match_row(label.strip(), _build_label_map(orig_ws)) or row
    else:
        restored_row = row

    # Original unaltered value columns (E+) map to restored D+ (one column left).
    col_idx = column_index_from_string(col)
    if col_idx >= 5:
        new_col = get_column_letter(col_idx - 1)
    else:
        new_col = col

    return f"{prefix}!{d1}{new_col}{d2}{restored_row}"


def _build_label_map(ws) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        key = _norm_label(ws.cell(r, 1).value)
        if key and key not in mapping:
            mapping[key] = r
    return mapping


def _match_row(label: str, orig_map: dict[str, int]) -> int | None:
    key = _norm_label(label)
    if key in orig_map:
        return orig_map[key]
    candidates = get_close_matches(key, orig_map.keys(), n=1, cutoff=0.82)
    return orig_map[candidates[0]] if candidates else None


def _apply_source_cell(cell, label: str | None, url: str | None, location: str | None) -> bool:
    if not label:
        return False
    cell.value = label
    if url and str(url).startswith("http"):
        cell.hyperlink = url
    elif location:
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location)
    cell.font = Font(color=BLUE, underline="single", italic=True, size=9)
    return True


def _collect_original_source(
    orig_ws, row: int, source_col: int = 3
) -> tuple[str | None, str | None, str | None]:
    c_cell = orig_ws.cell(row, source_col)
    label = _clean_source_label(c_cell.value)
    url = None
    location = None
    if c_cell.hyperlink:
        url = c_cell.hyperlink.target
        location = c_cell.hyperlink.location
    if not label and url:
        label = "Source"
    return label, url, location


def _insert_doc_columns(wb) -> dict[str, dict[str, str]]:
    """Insert Notes + Source at col B; return per-sheet column shift maps."""
    shifts: dict[str, dict[str, str]] = {}
    for name in INSERT_BC_SHEETS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        shift = _col_shift_insert_at_b(max_col=ws.max_column + 2)
        shifts[name] = shift
        ws.insert_cols(2, 2)

    for name in INSERT_BC_SHEETS:
        if name not in wb.sheetnames:
            continue
        shift = shifts[name]
        _rewrite_all_formulas(wb, name, shift)
        _rewrite_sheet_internal_formulas(wb[name], shift)

    return shifts


def _rebind_wacc_values(ws, value_col: str = "D") -> None:
    v = value_col

    def ref(row: int | None) -> str:
        if row is None:
            raise ValueError("Missing WACC label row")
        return f"{v}{row}"

    r_mkt = _label_row(ws, "market value of equity")
    r_px = _label_row(ws, "share price")
    r_sh = _label_row(ws, "shares outstanding")
    r_lease = _label_row(ws, "operating lease")
    r_funded = _label_row(ws, "funded debt")
    r_total = _label_row(ws, "total debt equivalents")
    r_obs = _label_row(ws, "observed beta")
    r_ydebt = _label_row(ws, "yahoo total debt")
    r_ymcap = _label_row(ws, "yahoo market cap")
    r_de_u = _label_row(ws, "d/e for unlever")
    r_bu = _label_row(ws, "unlevered")
    r_de_r = _label_row(ws, "d/e for relever")
    r_beta = _label_row(ws, "beta used")
    r_tax = _label_row(ws, "tax rate")
    r_pretax = _label_row(ws, "pre-tax cost of debt")
    r_after = _label_row(ws, "after-tax cost of debt")
    r_coe = _label_row(ws, "cost of equity = rf")
    r_eqw = _label_row(ws, "equity weight")
    r_dw = _label_row(ws, "debt weight")
    r_wacc = _exact_label_row(ws, "WACC")

    ws[f"{v}{r_mkt}"] = f"={v}{r_px}*{v}{r_sh}"
    ws[f"{v}{r_total}"] = f"={v}{r_lease}+{v}{r_funded}"
    ws[f"{v}{r_de_u}"] = f"={v}{r_ydebt}/{v}{r_ymcap}"
    ws[f"{v}{r_bu}"] = f"={v}{r_obs}/(1+(1-{v}{r_tax})*{v}{r_de_u})"
    ws[f"{v}{r_de_r}"] = f"={v}{r_total}/{v}{r_mkt}"
    ws[f"{v}{r_beta}"] = f"={v}{r_bu}*(1+(1-{v}{r_tax})*{v}{r_de_r})"
    ws[f"{v}{r_coe}"] = f"={v}3+{v}{r_beta}*{v}4"
    ws[f"{v}{r_after}"] = f"={v}{r_pretax}*(1-{v}{r_tax})"
    ws[f"{v}{r_eqw}"] = f"={v}{r_mkt}/({v}{r_mkt}+{v}{r_total})"
    ws[f"{v}{r_dw}"] = f"={v}{r_total}/({v}{r_mkt}+{v}{r_total})"
    ws[f"{v}{r_wacc}"] = f"={v}{r_eqw}*{v}{r_coe}+{v}{r_dw}*{v}{r_after}"


def _rebind_scenarios_wacc(wb) -> None:
    scn = wb["Scenarios"]
    wacc = wb["WACC"]
    r_wacc = _exact_label_row(wacc, "WACC")
    if r_wacc:
        scn["F9"] = f"=WACC!D{r_wacc}"


def _set_headers(wb) -> None:
    header_rows: dict[str, list[int]] = {
        "WACC": [2],
        "Scenarios": [3],
        "NOPAT Bridge": [2],
        "Comps": [2, 16, 41],
        "DCF": [2],
    }
    for sheet, rows in header_rows.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for r in rows:
            ws.cell(r, 2).value = HEADER_CELLS.get((sheet, r), NOTES_HEADER)
            ws.cell(r, 3).value = SOURCE_HEADER

    # Scrub any leftover prompt-style header text on col B.
    for sheet in INSERT_BC_SHEETS:
        ws = wb[sheet]
        for r in range(1, min(ws.max_row + 1, 50)):
            b = ws.cell(r, 2)
            if isinstance(b.value, str) and re.search(
                r"justification|~20\s*word|click\s*\+|\[cols", b.value, re.I
            ):
                b.value = NOTES_HEADER
            c = ws.cell(r, 3)
            if isinstance(c.value, str) and re.search(
                r"justification|source\s*\(click\)|ctrl\+f", c.value, re.I
            ):
                c.value = SOURCE_HEADER

    rd = wb["Revenue Drivers"]
    rd.cell(RD_HEADER_ROW, RD_NOTES_COL).value = NOTES_HEADER
    rd.cell(RD_HEADER_ROW, RD_SOURCE_COL).value = SOURCE_HEADER


def _fix_freeze_panes(wb) -> None:
    freezes = {
        "WACC": "E3",
        "Scenarios": "G3",
        "NOPAT Bridge": "G3",
        "DCF": "G3",
        "Comps": "G3",
    }
    for name, pane in freezes.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        ws.freeze_panes = pane


def _populate_notes_and_sources(wb, orig_wb, sheet_shifts: dict[str, dict[str, str]]) -> tuple[int, int]:
    notes_count = 0
    source_count = 0

    for sheet in INSERT_BC_SHEETS:
        if sheet not in wb.sheetnames or sheet not in orig_wb.sheetnames:
            continue
        ws = wb[sheet]
        orig_ws = orig_wb[sheet]
        orig_map = _build_label_map(orig_ws)

        for r in range(1, ws.max_row + 1):
            label_raw = ws.cell(r, 1).value
            if label_raw is None:
                continue
            label = str(label_raw).strip()
            if not label:
                continue

            orig_row = _match_row(label, orig_map)
            orig_b = orig_ws.cell(orig_row, 2).value if orig_row else None
            orig_label, orig_url, orig_loc = (
                _collect_original_source(orig_ws, orig_row) if orig_row else (None, None, None)
            )

            note = _note_for_row(sheet, label, str(orig_b) if orig_b else None)
            if note:
                ws.cell(r, 2).value = note
                notes_count += 1

            if orig_loc:
                orig_loc = _resolve_internal_link(orig_wb, wb, orig_loc, sheet_shifts)
            if _apply_source_cell(ws.cell(r, 3), orig_label, orig_url, orig_loc):
                source_count += 1

    # Revenue Drivers: append without shifting data columns.
    if "Revenue Drivers" in wb.sheetnames and "Revenue Drivers" in orig_wb.sheetnames:
        rd = wb["Revenue Drivers"]
        orig_rd = orig_wb["Revenue Drivers"]
        orig_map = _build_label_map(orig_rd)
        for r in range(1, rd.max_row + 1):
            label_raw = rd.cell(r, 1).value
            if label_raw is None:
                continue
            label = str(label_raw).strip()
            if not label or r == RD_HEADER_ROW:
                continue
            orig_row = _match_row(label, orig_map)
            if not orig_row:
                continue
            orig_j = orig_rd.cell(orig_row, 10).value  # col J justification
            note = _truncate_to_words(str(orig_j) if orig_j else None)
            if note:
                rd.cell(r, RD_NOTES_COL).value = note
                notes_count += 1
            src_label, src_url, src_loc = _collect_original_source(orig_rd, orig_row, source_col=11)
            if src_loc:
                src_loc = _resolve_internal_link(orig_wb, wb, src_loc, sheet_shifts)
            if _apply_source_cell(rd.cell(r, RD_SOURCE_COL), src_label, src_url, src_loc):
                source_count += 1
            elif src_label:
                rd.cell(r, RD_SOURCE_COL).value = src_label
                rd.cell(r, RD_SOURCE_COL).font = Font(
                    color=BLUE, underline="single", italic=True, size=9
                )
                source_count += 1

    return notes_count, source_count


def _fix_scenarios_scenario_headers(ws) -> None:
    """Bear / Base / Bull labels move from C-E to E-G after insert."""
    for old_col, text in ((3, "Bear"), (4, "Base"), (5, "Bull")):
        cell = ws.cell(2, old_col)
        if cell.value == text:
            cell.value = None
    for new_col, text in ((5, "Bear"), (6, "Base"), (7, "Bull")):
        if ws.cell(2, new_col).value is None:
            ws.cell(2, new_col).value = text


def _fix_comps_peer_formulas(ws) -> None:
    """Ensure peer EV/EBITDA formulas reference shifted EV/EBITDA columns."""
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        if not label or label.lower() == "company":
            continue
        val_cell = ws.cell(r, 4)
        if isinstance(val_cell.value, str) and val_cell.value.startswith("=IF("):
            if "I" in val_cell.value and "H" in val_cell.value:
                continue
            val_cell.value = f"=IF(I{r}<=0,\"n.m.\",H{r}/I{r})"


def _fix_comps_dcf_links(ws) -> None:
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, 4)
        if isinstance(cell.value, str) and "DCF!" in cell.value:
            cell.value = re.sub(r"DCF!B(\d+)", r"DCF!D\1", cell.value)


def restore_workbook(
    target: Path = TARGET,
    original: Path = ORIGINAL,
) -> Path:
    if not target.exists():
        raise FileNotFoundError(target)
    if not original.exists():
        raise FileNotFoundError(original)

    tmp = target.with_suffix(".restoring.xlsx")
    shutil.copy2(target, tmp)
    wb = openpyxl.load_workbook(tmp)
    orig_wb = openpyxl.load_workbook(original, data_only=False)

    sheet_shifts = _insert_doc_columns(wb)
    _rebind_wacc_values(wb["WACC"], "D")
    _rebind_scenarios_wacc(wb)
    _fix_scenarios_scenario_headers(wb["Scenarios"])
    _fix_comps_peer_formulas(wb["Comps"])
    _fix_comps_dcf_links(wb["Comps"])
    _set_headers(wb)
    notes, sources = _populate_notes_and_sources(wb, orig_wb, sheet_shifts)
    _fix_freeze_panes(wb)

    wb.save(tmp)
    tmp.replace(target)

    # Fix base-case assumption locks that landed on empty Scenarios col H.
    from repair_scenarios_refs import repair_workbook as _repair_scenarios_refs  # noqa: E402

    repaired, _ = _repair_scenarios_refs(target)
    if repaired:
        print(f"  Repaired {repaired} Scenarios!$H$→$F$ assumption refs")

    print(f"Restored {target.name}: {notes} notes, {sources} source links")
    return target


def verify(path: Path = TARGET) -> None:
    wb = openpyxl.load_workbook(path, data_only=False)
    issues: list[str] = []

    wacc = wb["WACC"]
    if not str(wacc["D10"].value or "").startswith("="):
        issues.append("WACC D10 market cap formula missing")
    if wacc["C8"].hyperlink is None:
        issues.append("WACC C8 share price source link missing")

    dcf = wb["DCF"]
    if "Scenarios!D" in str(dcf["E5"].value or ""):
        issues.append("DCF still references Scenarios!D (should be F)")
    if "Scenarios!$H$" in str(dcf["E15"].value or ""):
        issues.append("DCF E15 still references empty Scenarios col H (should be $F$12)")
    if str(dcf["E15"].value or "") != "=Scenarios!$F$12":
        issues.append(f"DCF E15 D&A % link wrong: {dcf['E15'].value}")
    if str(dcf["E17"].value or "") != "=Scenarios!$F$13":
        issues.append(f"DCF E17 Capex % link wrong: {dcf['E17'].value}")

    scn = wb["Scenarios"]
    if not str(scn["F9"].value or "").startswith("=WACC!D"):
        issues.append(f"Scenarios F9 WACC link wrong: {scn['F9'].value}")

    for sheet in INSERT_BC_SHEETS:
        ws = wb[sheet]
        links = sum(
            1
            for r in range(1, ws.max_row + 1)
            if ws.cell(r, 3).hyperlink is not None
        )
        if sheet == "WACC" and links < 10:
            issues.append(f"{sheet}: only {links} source links")

    for sheet in ("WACC", "Scenarios", "NOPAT Bridge", "Comps"):
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            val = ws.cell(r, 2).value
            if not isinstance(val, str) or not val.strip():
                continue
            if val in (NOTES_HEADER, SOURCE_HEADER):
                continue
            if not _lines_ok(val):
                issues.append(f"{sheet} B{r}: note exceeds 8 words: {val!r}")

    comment_count = sum(
        1
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if cell.comment
    )
    if comment_count < 100:
        issues.append(f"Analyst comments dropped: {comment_count}")

    if issues:
        raise AssertionError("Verification failed:\n" + "\n".join(issues))

    print(
        f"Verify OK: source columns restored; {comment_count} Analyst comments preserved"
    )


if __name__ == "__main__":
    restore_workbook()
    verify()
