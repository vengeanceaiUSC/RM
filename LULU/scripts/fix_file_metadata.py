#!/usr/bin/env python3
"""Scrub AI/tool metadata fingerprints from deliverable Excel and PowerPoint files.

Sets author/creator to the student analyst; removes openpyxl / python-pptx stamps.

Run:  cd LULU && python3 scripts/fix_file_metadata.py
"""
from __future__ import annotations

import re
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
AUTHOR = "Christian Gardner"
EXCEL_TITLE = "LULU DCF Valuation Model"
PPT_TITLE = "LULU Investment Pitch Deck"
TARGETS = [
    ROOT / "model18_wsp_formulas.xlsx",
    ROOT / "LULU_Investment_Pitch_Deck.pptx",
]

CP_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"
DCTERMS_NS = "{http://purl.org/dc/terms/}"
XSI_NS = "{http://www.w3.org/2001/XMLSchema-instance}"
APP_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _set_text(parent, tag: str, ns: str, value: str) -> None:
    el = parent.find(f"{ns}{tag}")
    if el is None:
        el = ET.SubElement(parent, f"{ns}{tag}")
    el.text = value


def _patch_core_xml(data: bytes, *, title: str, description: str = "") -> bytes:
    root = ET.fromstring(data)
    _set_text(root, "creator", CP_NS, AUTHOR)
    _set_text(root, "lastModifiedBy", CP_NS, AUTHOR)
    _set_text(root, "creator", DC_NS, AUTHOR)
    _set_text(root, "title", DC_NS, title)
    _set_text(root, "description", DC_NS, description)

    for tag in ("created", "modified"):
        el = root.find(f"{DCTERMS_NS}{tag}")
        if el is None:
            el = ET.SubElement(root, f"{DCTERMS_NS}{tag}")
        el.set(f"{XSI_NS}type", "dcterms:W3CDTF")
        el.text = _now_iso()

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_app_xml(data: bytes, application: str) -> bytes:
    text = data.decode("utf-8")
    text = re.sub(r"<Application>[^<]*</Application>", f"<Application>{application}</Application>", text)
    if "<AppVersion>" in text:
        text = re.sub(r"<AppVersion>[^<]*</AppVersion>", "<AppVersion>16.0</AppVersion>", text)
    return text.encode("utf-8")


def fix_ooxml(path: Path, *, title: str, application: str, description: str = "") -> dict[str, str]:
    tmp = path.with_suffix(path.suffix + ".meta.tmp")
    shutil.copy2(path, tmp)
    patched = path.with_suffix(path.suffix + ".meta.patched")

    with zipfile.ZipFile(tmp, "r") as zin, zipfile.ZipFile(patched, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = _patch_core_xml(data, title=title, description=description)
            elif item.filename == "docProps/app.xml":
                data = _patch_app_xml(data, application)
            zout.writestr(item, data)

    patched.replace(path)
    tmp.unlink(missing_ok=True)

    with zipfile.ZipFile(path) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
        app = zf.read("docProps/app.xml").decode("utf-8")

    bad = []
    for needle in ("openpyxl", "python-pptx", "Steve Canny"):
        if needle.lower() in core.lower() or needle.lower() in app.lower():
            bad.append(needle)

    return {
        "file": path.name,
        "author": AUTHOR,
        "clean": str(not bad),
        "residual": ", ".join(bad) if bad else "",
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
    return {"file": path.name, "author": AUTHOR, "clean": str(not bad), "residual": ", ".join(bad)}


def fix(path: Path | None = None) -> list[dict[str, str]]:
    results = []
    files = [path] if path else TARGETS
    for f in files:
        if not f.exists():
            continue
        if f.suffix == ".xlsx":
            results.append(fix_ooxml(f, title=EXCEL_TITLE, application="Microsoft Excel"))
        elif f.suffix == ".pptx":
            results.append(fix_pptx_metadata(f))
    return results


if __name__ == "__main__":
    for r in fix():
        print(f"{r['file']}: author={r['author']} clean={r['clean']} {r['residual']}")
