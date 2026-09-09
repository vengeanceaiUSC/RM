"""Geometry audit for LULU_Investment_Pitch_Deck.pptx.

Reports, per slide:
  * OVERLAP   - two shapes whose boxes intersect by more than a hairline
  * BLEED     - a shape crossing the slide edge or the footer line
  * ESTIMATED OVERFLOW - text taller than the box that holds it
  * WHITESPACE - fraction of the content band with nothing drawn in it

Run:  python3 audit_layout.py [--slides 1-13]
Exit code is always 0; this is a report, not a gate.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "GIS"))
from pptx import Presentation
from pptx.util import Emu
from gis_pitch import est_height

DECK = os.path.join(os.path.dirname(__file__), "..", "LULU_Investment_Pitch_Deck.pptx")

EPS = Emu(int(0.02 * 914400))  # ignore hairline touches from adjacent fills
FOOTER_Y_IN = 7.08
CONTENT_TOP_IN = 1.28


def _in(emu):
    return emu / 914400.0


def _boxes(slide):
    out = []
    for i, sh in enumerate(slide.shapes):
        if sh.top is None or sh.left is None:
            continue
        out.append((i, sh))
    return out


def _label(sh):
    txt = ""
    if sh.has_text_frame:
        txt = " ".join(sh.text_frame.text.split())[:40]
    return f"{sh.shape_type}:{txt!r}" if txt else f"{sh.shape_type}"


def _effective_bottom(sh):
    """Bottom edge as drawn: PowerPoint text spills past an undersized box."""
    bottom = sh.top + sh.height
    if sh.has_text_frame and sh.text_frame.text.strip():
        spill = int(max(_in(sh.height), _text_rows(sh)) * 914400)
        bottom = max(bottom, sh.top + spill)
    return bottom


def _intersect(a, b):
    x = min(a.left + a.width, b.left + b.width) - max(a.left, b.left)
    y = min(_effective_bottom(a), _effective_bottom(b)) - max(a.top, b.top)
    if x > EPS and y > EPS:
        return _in(x), _in(y)
    return None


def _text_rows(sh):
    """Rough wrapped-line count for a text frame."""
    if not sh.has_text_frame:
        return 0.0
    width_in = max(0.1, _in(sh.width) - 0.2)
    rows = 0.0
    for para in sh.text_frame.paragraphs:
        text = "".join(r.text for r in para.runs)
        sizes = [r.font.size.pt for r in para.runs if r.font.size]
        size = max(sizes) if sizes else 12.0
        rows += est_height(text, width_in, size)
        rows += (para.space_after.pt if para.space_after else 0) / 72.0
    return rows


def audit(slide, n, w_in, h_in):
    findings = []
    shapes = _boxes(slide)

    for ai in range(len(shapes)):
        for bi in range(ai + 1, len(shapes)):
            a, b = shapes[ai][1], shapes[bi][1]
            a_txt = a.has_text_frame and a.text_frame.text.strip()
            b_txt = b.has_text_frame and b.text_frame.text.strip()
            if not (a_txt and b_txt):
                continue  # a plain fill behind a label is intentional
            hit = _intersect(a, b)
            if hit:
                findings.append(
                    f"OVERLAP {hit[0]:.2f}x{hit[1]:.2f}in between {_label(a)} and {_label(b)}"
                )

    for _, sh in shapes:
        if _in(sh.left) < -0.01 or _in(sh.top) < -0.01:
            findings.append(f"BLEED off top/left: {_label(sh)}")
        if _in(sh.left + sh.width) > w_in + 0.01:
            findings.append(f"BLEED off right edge: {_label(sh)}")
        if _in(sh.top + sh.height) > h_in + 0.01:
            findings.append(f"BLEED off bottom edge: {_label(sh)}")
        if sh.has_text_frame and sh.text_frame.text.strip():
            need = _text_rows(sh)
            have = _in(sh.height)
            if need > have + 0.05:
                findings.append(
                    f"ESTIMATED OVERFLOW {need:.2f}in of text in {have:.2f}in box: {_label(sh)}"
                )

    band_top, band_bottom = CONTENT_TOP_IN, FOOTER_Y_IN
    filled = []
    for _, sh in shapes:
        plain_text = (
            sh.has_text_frame
            and not sh.has_table
            and str(sh.fill.type) in ("None", "BACKGROUND (5)")
        )
        if plain_text and not sh.text_frame.text.strip():
            continue
        top, bottom = _in(sh.top), _in(sh.top + sh.height)
        if bottom <= band_top or top >= band_bottom:
            continue
        if plain_text:
            # An empty tail inside a text box is whitespace, not content.
            bottom = min(bottom, top + max(_text_rows(sh), 0.2))
        filled.append((max(top, band_top), min(bottom, band_bottom)))
    filled.sort()
    merged = []
    for seg in filled:
        if merged and seg[0] <= merged[-1][1] + 0.02:
            merged[-1] = (merged[-1][0], max(merged[-1][1], seg[1]))
        else:
            merged.append(seg)
    used = sum(b - a for a, b in merged)
    band = band_bottom - band_top
    empty_pct = 100.0 * (1 - used / band)
    return findings, empty_pct


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--slides", default="1-22")
    args = ap.parse_args()
    lo, _, hi = args.slides.partition("-")
    lo, hi = int(lo), int(hi or lo)

    prs = Presentation(DECK)
    w_in, h_in = _in(prs.slide_width), _in(prs.slide_height)
    total = 0
    for n, slide in enumerate(prs.slides, 1):
        if not lo <= n <= hi:
            continue
        findings, empty_pct = audit(slide, n, w_in, h_in)
        total += len(findings)
        flag = "  <-- sparse" if empty_pct > 45 else ""
        print(f"SLIDE {n}: {len(findings)} issue(s), {empty_pct:.0f}% vertical whitespace{flag}")
        for f in findings:
            print(f"    {f}")
    print(f"\n{total} geometry issue(s) across slides {lo}-{hi}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
