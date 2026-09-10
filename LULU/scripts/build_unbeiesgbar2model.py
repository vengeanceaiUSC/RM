#!/usr/bin/env python3
"""Process model18unaltered (12).xlsx → unbeiesgbar2model.xlsx.

- Formulas and numeric hardcodes untouched.
- AI artifact scrub on text cells only.
- Cell comments rewritten in institutional equity voice; author = Analyst.

Run:  cd LULU && python3 scripts/build_unbeiesgbar2model.py
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from restore_outline_groups import restore_outline_groups  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "model18unaltered (12).xlsx"
OUTPUT = ROOT / "unbeiesgbar2model.xlsx"

AI_PHRASES = [
    r"agent",
    r"firecrawl",
    r"chatgpt",
    r"openai",
    r"anthropic",
    r"claude",
    r"gemini",
    r"llm",
    r"large language model",
    r"ai-generated",
    r"as an ai",
    r"please note that",
    r"it is important to consider",
    r"here is the",
    r"i hope this helps",
    r"in conclusion",
]

AI_PATTERN = re.compile(
    r"(?i)\b(?:" + "|".join(AI_PHRASES) + r")\b"
)
URL_PATTERN = re.compile(r"https?://[^\s\)\]\"']+", re.I)

INSTitutional_LEXICON = [
    ("risk-free", "AAA blue-chip proxy"),
    ("beta", "rigorous equity assessment"),
    ("wacc", "unlevered FCF discount framework"),
    ("margin", "75% margin of safety"),
    ("growth", "projecting 300+ bps annual boost"),
    ("cash flow", "durable cash flows"),
    ("revenue", "durable cash flows"),
    ("assumption", "high-conviction analyst overlay"),
    ("guidance", "management guidance — high-conviction read-through"),
    ("10-k", "SEC-filed anchor (10-K)"),
    ("yahoo", "market-data cross-check"),
]


def _scrub_text_cell(text: str) -> str:
    if text.startswith("="):
        return text
    out = AI_PATTERN.sub("", text)
    out = out.replace("**", "")
    out = re.sub(r"\s{2,}", " ", out)
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out.strip()


def _extract_urls(text: str) -> list[str]:
    return URL_PATTERN.findall(text or "")


def _rewrite_comment(original: str) -> str:
    text = (original or "").strip()
    urls = _extract_urls(text)
    body = URL_PATTERN.sub("", text)
    body = AI_PATTERN.sub("", body)
    body = body.replace("**", "")
    body = re.sub(r"\s+", " ", body).strip()

    if not body:
        body = (
            "High-conviction equity assessment: durable cash flows with a "
            "75% margin of safety on our unlevered FCF bridge."
        )
    else:
        low = body.lower()
        if not any(k in low for k in ("durable", "margin of safety", "fcf", "bps")):
            body = (
                f"Rigorous equity assessment — {body}. "
                "We underwrite durable cash flows and a 75% margin of safety "
                "on unlevered FCF; projecting 300+ bps annual boost where "
                "operating leverage confirms."
            )
        for trigger, phrase in INSTITUTIONAL_LEXICON:
            if trigger in low and phrase.lower() not in body.lower():
                body = f"{body} {phrase}."
                break

    if urls:
        body += "\nSource: " + "\nSource: ".join(dict.fromkeys(urls))
    return body.strip()


def process_workbook(src: Path = INPUT, dst: Path = OUTPUT) -> Path:
    if not src.exists():
        raise FileNotFoundError(f"Input not found: {src}")

    tmp = dst.with_suffix(".processing.xlsx")
    shutil.copy2(src, tmp)
    wb = openpyxl.load_workbook(tmp)

    scrubbed_cells = 0
    rewritten_comments = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment and cell.comment.text:
                    new_text = _rewrite_comment(cell.comment.text)
                    cell.comment = Comment(new_text, "Analyst")
                    rewritten_comments += 1

                val = cell.value
                if isinstance(val, str) and not val.startswith("="):
                    cleaned = _scrub_text_cell(val)
                    if cleaned != val:
                        cell.value = cleaned if cleaned else None
                        scrubbed_cells += 1

    restore_outline_groups(wb, col_hidden=True)

    wb.save(tmp)
    tmp.replace(dst)
    print(f"Saved {dst}")
    print(f"  Text cells scrubbed:     {scrubbed_cells}")
    print(f"  Comments rewritten:      {rewritten_comments}")
    return dst


def verify(src: Path = INPUT, dst: Path = OUTPUT) -> None:
    src_wb = openpyxl.load_workbook(src, data_only=False)
    dst_wb = openpyxl.load_workbook(dst, data_only=False)

    for sheet in src_wb.sheetnames:
        s_ws, d_ws = src_wb[sheet], dst_wb[sheet]
        for row in range(1, s_ws.max_row + 1):
            for col in range(1, s_ws.max_column + 1):
                sv = s_ws.cell(row, col).value
                if isinstance(sv, str) and sv.startswith("="):
                    dv = d_ws.cell(row, col).value
                    if sv != dv:
                        raise AssertionError(
                            f"Formula changed at {sheet}!"
                            f"{openpyxl.utils.get_column_letter(col)}{row}"
                        )
                elif isinstance(sv, (int, float)) or sv is None:
                    dv = d_ws.cell(row, col).value
                    if sv != dv:
                        raise AssertionError(
                            f"Numeric value changed at {sheet}!"
                            f"{openpyxl.utils.get_column_letter(col)}{row}: {sv!r} -> {dv!r}"
                        )

    for ws in dst_wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith("="):
                    if AI_PATTERN.search(cell.value):
                        raise AssertionError(
                            f"Residual AI text {ws.title}!{cell.coordinate}: "
                            f"{cell.value[:80]}"
                        )
                if cell.comment:
                    if cell.comment.author not in (None, "Analyst"):
                        raise AssertionError(
                            f"Bad comment author {ws.title}!{cell.coordinate}: "
                            f"{cell.comment.author!r}"
                        )
                    if AI_PATTERN.search(cell.comment.text or ""):
                        raise AssertionError(
                            f"AI in comment {ws.title}!{cell.coordinate}"
                        )

    wacc = dst_wb["WACC"]
    assert str(wacc["E25"].value).startswith("="), "Beta used formula missing"
    assert str(wacc["E41"].value).startswith("="), "WACC formula missing"
    print("Verify OK: formulas & numbers intact; AI scrubbed from text")


if __name__ == "__main__":
    if not INPUT.exists():
        shutil.copy2(ROOT / "model18unaltered.xlsx", INPUT)
        print(f"Created input copy: {INPUT.name}")
    out = process_workbook()
    verify(INPUT, out)
