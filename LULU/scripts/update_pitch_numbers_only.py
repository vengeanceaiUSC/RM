#!/usr/bin/env python3
"""Update ONLY numeric values in the pitch deck to match recalculated Excel DCF.

Preserves all original slide copy verbatim — no narrative rewrites.
Also standardizes model source labels (filename + Scenarios col C).

Run:  cd LULU && python3 scripts/update_pitch_numbers_only.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

from pptx import Presentation

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))

from recalc_model18 import read_dcf_outputs, recalc_workbook

MODEL = "LULU_DCF_Valuation_Model.xlsx"
FALLBACK = Path("/tmp/orig_deck.pptx")
OUT = ROOT / "LULU_Investment_Pitch_Deck.pptx"


def _replace_text(text: str, mapping: list[tuple[str, str]]) -> str:
    for old, new in mapping:
        text = text.replace(old, new)
    text = re.sub(r"\$134(?!\.\d)", "$133.52", text)
    return text


def _fmt_m(v: int) -> str:
    return f"{v:,}"


def _patch_range(text: str, label: str, vals: list[int], *, money: bool = True) -> str:
    """Patch 'Label: FY26 Xm to FY30 Ym' forecast bullets."""
    fy26, fy30 = vals[0], vals[-1]
    if money:
        repl = f"{label}: FY26 {_fmt_m(fy26)}M to FY30 {_fmt_m(fy30)}M"
        return re.sub(
            rf"{re.escape(label)}: FY26 [\d,]+M to FY30 [\d,]+M",
            repl,
            text,
            count=1,
        )
    return text


def _patch_eps_narratives(text: str, eps: list[float]) -> str:
    text = re.sub(
        r"(compounding EPS to )\$[\d.]+",
        rf"\g<1>${eps[4]:.2f}",
        text,
        count=1,
    )
    text = re.sub(
        r"(buybacks compound EPS to )\$[\d.]+",
        rf"\g<1>${eps[1]:.2f}",
        text,
        count=1,
    )
    text = re.sub(
        r"(Diluted EPS: FY26 )\$[\d.]+( to FY30 )\$[\d.]+",
        rf"\g<1>${eps[0]:.2f}\g<2>${eps[4]:.2f}",
        text,
        count=1,
    )
    return text


def _patch_bear_bull_line(text: str, bear_px: str, bull_px: str) -> str:
    text = re.sub(r"202\.17\.17+", bull_px, text)
    text = re.sub(
        r"Bear \$[\d.]+ / bull \$[\d.]+",
        f"Bear ${bear_px} / bull ${bull_px}",
        text,
    )
    return text


def _standardize_sources(text: str) -> str:
    """One model filename; Scenarios base case is col C (post layout compact)."""

    def _provided(m: re.Match[str]) -> str:
        detail = m.group(1)
        if detail:
            return f"{MODEL}, {detail}"
        return MODEL

    text = re.sub(
        r"Provided Valuation Model(?: \(([^)]+)\))?: LULU_DCF_Valuation_Model\.xlsx",
        _provided,
        text,
    )
    text = text.replace("Scenarios col G", "Scenarios col C")
    text = text.replace("Scenarios tab, col G", "Scenarios tab, col C")
    text = text.replace("base case (Scenarios col G)", "base case (Scenarios col C)")
    text = text.replace("pitch-bridge block, col G", "pitch-bridge block, col C")
    text = re.sub(r", col G\b", ", col C", text)
    return text


def _update_fin_table(table, row_map: dict[int, list], eps_row: int | None = None) -> None:
    for ri, vals in row_map.items():
        for ci, val in enumerate(vals, start=5):
            if ri == eps_row:
                table.cell(ri, ci).text = f"{val:.2f}"
            else:
                table.cell(ri, ci).text = _fmt_m(int(val))


def _patch_forecast_narratives(text: str, b: dict) -> str:
    text = _patch_range(text, "Net revenue", b["revenue_m"])
    text = _patch_range(text, "Gross profit", b["gp_m"])
    text = _patch_range(text, "Operating income", b["ebit_m"])
    text = _patch_range(text, "Net income", b["ni_m"])
    text = _patch_range(text, "Cash & equivalents", b["cash_m"])
    text = _patch_range(text, "Inventories", b["inv_m"])
    text = _patch_range(text, "Total assets", b["ta_m"])
    text = _patch_range(text, "Total liabilities", b["tl_m"])
    text = _patch_range(text, "Total equity", b["te_m"])
    text = _patch_range(text, "Cash from operations", b["cfo_m"])
    text = _patch_range(text, "D&A (add-back)", b["dna_m"])
    text = _patch_range(text, "Free cash flow", b["fcf_m"])
    return text


def main() -> None:
    wb = recalc_workbook(ROOT / "model18_wsp_formulas.xlsx")
    data = read_dcf_outputs(wb)
    base = data["scenarios"]["base"]
    bear = data["scenarios"]["bear"]
    bull = data["scenarios"]["bull"]
    bridge = data["bridge"]
    b = data["base"]

    px = f"{base['implied_px']:.2f}"
    bear_px = f"{bear['implied_px']:.2f}"
    bull_px = f"{bull['implied_px']:.2f}"
    upside = f"{base['upside_pct']:.1f}"
    eps = b["eps"]

    mapping = [
        ("$133.64", f"${px}"),
        ("133.64", px),
        ("(+33.6% upside)", f"(+{upside}% upside)"),
        ("33.6% upside", f"{upside}% upside"),
        ("$54.00 bear", f"${bear_px} bear"),
        ("$54 bear", f"${bear_px} bear"),
        ("$56.00 bear", f"${bear_px} bear"),
        ("$56 bear", f"${bear_px} bear"),
        ("Bear $54", f"Bear ${bear_px}"),
        ("Bear $56", f"Bear ${bear_px}"),
        ("bear $54", f"bear ${bear_px}"),
        ("bear $56", f"bear ${bear_px}"),
        ("$54 / bull", f"${bear_px} / bull"),
        ("Bear $54 / bull $208", f"Bear ${bear_px} / bull ${bull_px}"),
        ("Bear $56 / bull $202", f"Bear ${bear_px} / bull ${bull_px}"),
        ("bull $208", f"bull ${bull_px}"),
        ("$208", f"${bull_px}"),
        ("$14.57", f"${eps[4]:.2f}"),
        ("14.57", f"{eps[4]:.2f}"),
        ("$14.87", f"${eps[4]:.2f}"),
        ("14.87", f"{eps[4]:.2f}"),
        ("$10.09", f"${eps[1]:.2f}"),
        ("10.09", f"{eps[1]:.2f}"),
        ("$10.11", f"${eps[1]:.2f}"),
        ("10.11", f"{eps[1]:.2f}"),
        ("$9.61", f"${eps[0]:.2f}"),
        ("9.61", f"{eps[0]:.2f}"),
        ("$134", f"${px}"),
        ("gives $134", f"gives ${px}"),
        ("3,944", _fmt_m(bridge["pv_fcf_m"])),
        ("14,876", _fmt_m(bridge["ev_m"])),
        ("14,885", _fmt_m(bridge["equity_m"])),
        ("1,087", _fmt_m(b["ni_m"][0])),
        ("1,057", _fmt_m(b["ni_m"][1])),
        ("1,056", _fmt_m(b["fcf_m"][3])),
        # stale BS total assets (pre-recalc)
        ("7,691", _fmt_m(b["ta_m"][0])),
        ("7,624", _fmt_m(b["ta_m"][1])),
        ("7,560", _fmt_m(b["ta_m"][2])),
        ("7,502", _fmt_m(b["ta_m"][3])),
        ("7,447", _fmt_m(b["ta_m"][4])),
    ]

    if not OUT.exists() and FALLBACK.exists():
        shutil.copy2(FALLBACK, OUT)
    prs = Presentation(str(OUT))

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if not run.text:
                            continue
                        run.text = _replace_text(run.text, mapping)
                        run.text = _patch_eps_narratives(run.text, eps)
                        run.text = _patch_bear_bull_line(run.text, bear_px, bull_px)
                        run.text = _patch_forecast_narratives(run.text, b)
                        run.text = _standardize_sources(run.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text:
                            cell.text = _replace_text(cell.text, mapping)
                            cell.text = _standardize_sources(cell.text)

    # Slide 14 — income statement
    for shape in prs.slides[13].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            _update_fin_table(
                shape.table,
                {1: b["revenue_m"], 2: b["gp_m"], 3: b["ebit_m"], 5: b["ni_m"], 6: eps},
                eps_row=6,
            )

    # Slide 15 — balance sheet
    for shape in prs.slides[14].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            _update_fin_table(
                shape.table,
                {1: b["cash_m"], 2: b["inv_m"], 3: b["ta_m"], 4: b["tl_m"], 5: b["te_m"]},
            )

    # Slide 21 — comps implied P/E uses FY26 EPS
    for shape in prs.slides[20].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text.strip() == "Metric":
            shape.table.cell(4, 2).text = f"${eps[0]:.2f}"

    # Slide 16 — cash flow
    for shape in prs.slides[15].shapes:
        if shape.has_table and shape.table.rows[0].cells[0].text == "US$ M":
            t = shape.table
            _update_fin_table(t, {1: b["cfo_m"], 2: b["dna_m"], 4: b["fcf_m"]})
            for ci, val in enumerate(b["capex_m"], start=5):
                t.cell(3, ci).text = f"({val:,})"
            for ci, val in enumerate(b["buy_m"], start=5):
                t.cell(5, ci).text = f"({val:,})"

    prs.save(str(OUT))

    from fix_file_metadata import fix as fix_meta
    from fix_pitch_gis_compliance import main as fix_gis_layout
    from fix_table_header_colors import fix as fix_table_headers
    from repair_pitch_deck import repair_and_repackage
    from sync_sensitivity_from_dcf import sync as sync_sensitivity

    fix_gis_layout()
    sync_sensitivity(OUT)
    from fix_pitch_gis_compliance import _fix_slide20

    prs_align = Presentation(str(OUT))
    _fix_slide20(prs_align)
    prs_align.save(str(OUT))
    fix_table_headers(OUT)
    fix_meta(OUT)
    repair_and_repackage(OUT)
    from fix_pitch_gis_compliance import _rebuild_toc
    from fix_pitch_gis_formatting import _read_toc_entries

    prs_toc = Presentation(str(OUT))
    entries = _read_toc_entries(prs_toc.slides[0])
    if entries:
        _rebuild_toc(prs_toc.slides[0], entries)
        prs_toc.save(str(OUT))
    fix_meta(OUT)
    fix_meta(ROOT / "model18_wsp_formulas.xlsx")

    from apply_garamond_excel import apply as apply_garamond

    apply_garamond(ROOT / "model18_wsp_formulas.xlsx")

    print(f"Updated numbers + sources → {OUT}")
    print(f"  Base ${px} (+{upside}%) | Bear ${bear_px} | Bull ${bull_px}")
    print(f"  EPS FY26-FY30: {eps}")
    print(f"  BS TA FY26-FY30: {b['ta_m']}")


if __name__ == "__main__":
    main()
