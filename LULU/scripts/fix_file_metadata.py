#!/usr/bin/env python3
"""Scrub AI/tool metadata fingerprints from deliverable Excel and PowerPoint files.

Sets author/creator to the student analyst; removes openpyxl / python-pptx stamps.

Run:  cd LULU && python3 scripts/fix_file_metadata.py
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
AUTHOR = "Christian Gardner"
EXCEL_TITLE = "LULU DCF Valuation Model"
PPT_TITLE = "LULU Investment Pitch Deck"
TARGETS = [
    ROOT / "model18_wsp_formulas.xlsx",
    ROOT / "LULU_Investment_Pitch_Deck.pptx",
]


def _patch_app_xml(data: bytes, application: str) -> bytes:
    text = data.decode("utf-8")
    text = re.sub(r"<Application>[^<]*</Application>", f"<Application>{application}</Application>", text)
    if "<AppVersion>" in text:
        text = re.sub(r"<AppVersion>[^<]*</AppVersion>", "<AppVersion>16.0</AppVersion>", text)
    return text.encode("utf-8")


def fix_xlsx_metadata(path: Path) -> dict[str, str]:
    """Set XLSX metadata via openpyxl API — never zip-patch core.xml (breaks Excel)."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.properties.title = EXCEL_TITLE
    tmp = path.with_suffix(".meta.xlsx")
    wb.save(tmp)

    # app.xml Application tag is safe to regex-replace; core.xml is not.
    patched = path.with_suffix(".meta.patched")
    with zipfile.ZipFile(tmp, "r") as zin, zipfile.ZipFile(patched, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/app.xml":
                data = _patch_app_xml(data, "Microsoft Excel")
            zout.writestr(item, data)
    patched.replace(path)
    tmp.unlink(missing_ok=True)

    with zipfile.ZipFile(path) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
        app = zf.read("docProps/app.xml").decode("utf-8")
    bad = [n for n in ("openpyxl", "python-pptx", "Steve Canny") if n.lower() in (core + app).lower()]
    healthy = "cp:coreProperties" in core and "ns0:" not in core
    return {
        "file": path.name,
        "author": AUTHOR,
        "clean": str(not bad and healthy),
        "residual": ", ".join(bad) if bad else ("" if healthy else "bad core.xml namespaces"),
    }


def fix_pptx_metadata(path: Path) -> dict[str, str]:
    """Set PPTX metadata via python-pptx API — never zip-patch .pptx (breaks PowerPoint)."""
    from pptx import Presentation

    prs = Presentation(str(path))
    cp = prs.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.title = PPT_TITLE
    cp.comments = ""
    tmp = path.with_suffix(".meta.pptx")
    prs.save(str(tmp))
    tmp.replace(path)
    with zipfile.ZipFile(path) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
    bad = [n for n in ("openpyxl", "python-pptx", "Steve Canny") if n.lower() in core.lower()]
    healthy = "cp:coreProperties" in core and "ns0:" not in core
    return {
        "file": path.name,
        "author": AUTHOR,
        "clean": str(not bad and healthy),
        "residual": ", ".join(bad) if bad else ("" if healthy else "bad core.xml namespaces"),
    }


def fix(path: Path | None = None) -> list[dict[str, str]]:
    results = []
    files = [path] if path else TARGETS
    for f in files:
        if not f.exists():
            continue
        if f.suffix == ".xlsx":
            results.append(fix_xlsx_metadata(f))
        elif f.suffix == ".pptx":
            results.append(fix_pptx_metadata(f))
    return results


if __name__ == "__main__":
    for r in fix():
        print(f"{r['file']}: author={r['author']} clean={r['clean']} {r['residual']}")
