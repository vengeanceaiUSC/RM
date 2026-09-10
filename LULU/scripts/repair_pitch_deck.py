#!/usr/bin/env python3
"""Repair PPTX for Microsoft PowerPoint compatibility.

python-pptx can open files that PowerPoint rejects when:
  - docProps were zip-patched with ElementTree (namespace breakage)
  - table rows were cloned via deepcopy (duplicate OOXML extension IDs)

This script round-trips the deck through python-pptx and sets author metadata
via the API (no raw zip surgery on .pptx).

Run:  cd LULU && python3 scripts/repair_pitch_deck.py
"""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
DECK = ROOT / "LULU_Investment_Pitch_Deck.pptx"
AUTHOR = "Christian Gardner"
TITLE = "LULU Investment Pitch Deck"


def _core_xml_healthy(data: bytes) -> bool:
    text = data.decode("utf-8")
    return "cp:coreProperties" in text and "ns0:" not in text and text.count("<dc:creator>") == 1


def repair(path: Path = DECK) -> Path:
    prs = Presentation(str(path))
    cp = prs.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.title = TITLE
    cp.comments = ""

    tmp = path.with_suffix(".repairing.pptx")
    prs.save(str(tmp))

    # python-pptx preserves broken ns0: namespaces from zip-patched inputs;
    # restore a known-good core.xml if the round-trip did not heal metadata.
    import zipfile

    with zipfile.ZipFile(tmp, "r") as zin:
        core = zin.read("docProps/core.xml")
    if not _core_xml_healthy(core):
        healthy = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"<dc:title>{TITLE}</dc:title><dc:subject/><dc:creator>{AUTHOR}</dc:creator>"
            "<cp:keywords/><dc:description></dc:description>"
            f"<cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy><cp:revision>1</cp:revision>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2013-01-27T09:14:16Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2013-01-27T09:15:58Z</dcterms:modified>'
            "<cp:category/></cp:coreProperties>"
        ).encode("utf-8")
        patched = path.with_suffix(".corefix.pptx")
        with zipfile.ZipFile(tmp, "r") as zin, zipfile.ZipFile(patched, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "docProps/core.xml":
                    data = healthy
                zout.writestr(item, data)
        patched.replace(tmp)

    tmp.replace(path)
    return path


def repackage(path: Path = DECK) -> Path:
    """Repack OPC zip the way Office expects (fixes many PowerPoint open failures)."""
    opc = shutil.which("opc")
    if opc is None:
        local = Path.home() / ".local/bin/opc"
        opc = str(local) if local.exists() else None
    if opc is None:
        raise RuntimeError("opc-diag not installed (pip install opc-diag)")

    with tempfile.TemporaryDirectory(prefix="pptx_opc_") as tmp:
        extract_dir = Path(tmp) / "extract"
        out = path.with_suffix(".repacked.pptx")
        subprocess.run([opc, "extract", str(path), str(extract_dir)], check=True)
        subprocess.run([opc, "repackage", str(extract_dir), str(out)], check=True)
        out.replace(path)
    return path


def repair_and_repackage(path: Path = DECK) -> Path:
    repair(path)
    repackage(path)
    return path


if __name__ == "__main__":
    out = repair()
    print(f"Repaired → {out} (author={AUTHOR})")
