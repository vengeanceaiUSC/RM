#!/usr/bin/env python3
"""Remove openpyxl / Microsoft Excel metadata fingerprint from docProps.

User should Save As from native Excel to set their author name.

Run:  cd LULU && python3 scripts/fix_file_metadata.py
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "unbeiesgbar_final.xlsx"

CP_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"


def _patch_core_xml(data: bytes) -> bytes:
    root = ET.fromstring(data)
    for tag in ("creator", "lastModifiedBy"):
        el = root.find(f"{CP_NS}{tag}")
        if el is not None:
            root.remove(el)
    # Remove dc:creator if present
    for el in root.findall(f"{DC_NS}creator"):
        root.remove(el)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def fix(path: Path = TARGET) -> dict[str, str]:
    tmp = path.with_suffix(".meta.xlsx")
    shutil.copy2(path, tmp)
    wb = openpyxl.load_workbook(tmp)
    wb.properties.creator = None
    wb.properties.lastModifiedBy = None
    wb.properties.title = "LULU DCF Model"
    wb.save(tmp)

    # openpyxl re-stamps creator on save — patch raw core.xml
    patched = path.with_suffix(".patched.xlsx")
    with zipfile.ZipFile(tmp, "r") as zin, zipfile.ZipFile(patched, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = _patch_core_xml(data)
            zout.writestr(item, data)
    patched.replace(path)
    tmp.unlink(missing_ok=True)

    # Verify
    with zipfile.ZipFile(path) as zf:
        core = zf.read("docProps/core.xml").decode("utf-8")
    has_creator = "creator" in core.lower()
    return {"creator_stripped": str(not has_creator)}


if __name__ == "__main__":
    print(f"Metadata fix: {fix()}")
