#!/usr/bin/env python3
"""Read-only formula audit for model18_humanized workbook.

Scans every formula cell and classifies humanization/refactor issues.
Does NOT modify the workbook.

Run: cd LULU && python3 scripts/audit_formula_raw.py
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "model18_humanized (1).xlsx"
OUTPUT = ROOT / "formula_audit_raw.json"

EXCEL_FUNCTIONS = {
    "ABS", "ACCRINT", "ACOS", "ACOSH", "ADDRESS", "AGGREGATE", "AMORDEGRC", "AMORLINC", "AND",
    "ARABIC", "AREAS", "ASC", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "AVEDEV", "AVERAGE",
    "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BAHTTEXT", "BASE", "BESSELI", "BESSELJ", "BESSELK",
    "BESSELY", "BETA", "BETADIST", "BETAINV", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BINOM",
    "BINOMDIST", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "CEILING", "CELL",
    "CHAR", "CHIDIST", "CHIINV", "CHISQ", "CHITEST", "CHOOSE", "CLEAN", "CODE", "COLUMN",
    "COLUMNS", "COMBIN", "COMBINA", "COMPLEX", "CONCAT", "CONCATENATE", "CONFIDENCE", "CONVERT",
    "CORREL", "COS", "COSH", "COT", "COTH", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF",
    "COUNTIFS", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "COVAR",
    "CRITBINOM", "CSC", "CSCH", "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER", "CUBESET", "CUBESETCOUNT", "CUBEVALUE", "CUMIPMT", "CUMPRINC", "DATE",
    "DATEDIF", "DATEVALUE", "DAVERAGE", "DAY", "DAYS", "DAYS360", "DB", "DCOUNT", "DCOUNTA",
    "DDB", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DECIMAL", "DEGREES", "DELTA", "DEVSQ", "DGET",
    "DISC", "DMAX", "DMIN", "DOLLAR", "DOLLARDE", "DOLLARFR", "DPRODUCT", "DSTDEV", "DSTDEVP",
    "DSUM", "DURATION", "DVAR", "DVARP", "EDATE", "EFFECT", "ENCODEURL", "EOMONTH", "ERF",
    "ERFC", "ERROR", "EVEN", "EXACT", "EXP", "EXPON", "F", "FACT", "FACTDOUBLE", "FALSE",
    "FDIST", "FILTER", "FIND", "FINDB", "FINV", "FISHER", "FISHERINV", "FIXED", "FLOOR",
    "FORECAST", "FORMULATEXT", "FREQUENCY", "FTEST", "FV", "FVSCHEDULE", "GAMMA", "GAMMADIST",
    "GAMMALN", "GAUSS", "GCD", "GEOMEAN", "GESTEP", "GROWTH", "HARMEAN", "HEX2BIN", "HEX2DEC",
    "HEX2OCT", "HLOOKUP", "HOUR", "HYPERLINK", "HYPGEOM", "IF", "IFERROR", "IFNA", "IFS",
    "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC",
    "IMCSCH", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL",
    "IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT", "IMSUB", "IMSUM", "IMTAN", "INDEX",
    "INDIRECT", "INFO", "INT", "INTERCEPT", "INTRATE", "IPMT", "IRR", "ISBLANK", "ISERR",
    "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD",
    "ISREF", "ISTEXT", "ISO", "ISOWEEKNUM", "ISPMT", "JIS", "KURT", "LARGE", "LCM", "LEFT",
    "LEFTB", "LEN", "LENB", "LINEST", "LN", "LOG", "LOG10", "LOGEST", "LOGINV", "LOGNORM",
    "LOGNORMDIST", "LOOKUP", "LOWER", "MATCH", "MAX", "MAXA", "MAXIFS", "MDETERM", "MDURATION",
    "MEDIAN", "MID", "MIDB", "MIN", "MINA", "MINIFS", "MINUTE", "MINVERSE", "MIRR", "MMULT",
    "MOD", "MODE", "MODESNGL", "MONTH", "MROUND", "MULTINOMIAL", "MUNIT", "N", "NA", "NEGBINOM",
    "NEGBINOMDIST", "NETWORKDAYS", "NETWORKDAYSINTL", "NOMINAL", "NORM", "NORMDIST", "NORMINV",
    "NORMSDIST", "NORMSINV", "NOT", "NOW", "NPER", "NPV", "ODD", "ODDFPRICE", "ODDFYIELD",
    "ODDLPRICE", "ODDLYIELD", "OFFSET", "OR", "PDURATION", "PEARSON", "PERCENTILE",
    "PERCENTRANK", "PERMUT", "PERMUTATIONA", "PHI", "PI", "PMT", "POISSON", "POWER", "PPMT",
    "PRICE", "PRICEDISC", "PRICEMAT", "PROB", "PRODUCT", "PROPER", "PV", "QUARTILE", "QUOTIENT",
    "RADIANS", "RAND", "RANDBETWEEN", "RANK", "RATE", "RECEIVED", "REPLACE", "REPLACEB", "REPT",
    "RIGHT", "RIGHTB", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "RRI", "RSQ",
    "RTD", "SEARCH", "SEARCHB", "SEC", "SECH", "SECOND", "SERIESSUM", "SHEET", "SHEETS", "SIGN",
    "SIN", "SINH", "SKEW", "SLN", "SLOPE", "SMALL", "SORT", "SQRT", "SQRTPI", "STANDARDIZE",
    "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STEYX", "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF",
    "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "SWITCH", "SYD", "T",
    "TAN", "TANH", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "TDIST", "TEXT", "TEXTJOIN", "TIME",
    "TIMEVALUE", "TINV", "TODAY", "TRANSPOSE", "TREND", "TRIM", "TRIMMEAN", "TRUE", "TRUNC",
    "TTEST", "TYPE", "UNICHAR", "UNICODE", "UNIQUE", "UPPER", "VALUE", "VAR", "VARA", "VARP",
    "VARPA", "VDB", "VLOOKUP", "WEBSERVICE", "WEEKDAY", "WEEKNUM", "WEIBULL", "WEIBULLDIST",
    "WORKDAY", "WORKDAYINTL", "XIRR", "XLOOKUP", "XNPV", "XOR", "YEAR", "YEARFRAC", "YIELD",
    "YIELDDISC", "YIELDMAT", "ZTEST",
}

ISSUE_TYPES = (
    "self_referencing_sheet_name",
    "lowercase_function_names",
    "iferror_or_round_wrappers",
    "inline_numeric_hardcodes",
    "text_cells_explaining_formulas",
)

PROJECTION_SHEETS = {"Scenarios", "Revenue Drivers", "NOPAT Bridge", "DCF", "Comps", "WACC"}

SHEET_REF_RE = re.compile(r"(?:'([^']+)'|([^'!\s][^!]*))!", re.I)
FUNC_CALL_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
IFERROR_ROUND_RE = re.compile(r"\b(iferror|round|rounddown|roundup|mround)\s*\(", re.I)
FORMULA_NUM_RE = re.compile(
    r"(?<![A-Z$])(?<![A-Z]\$)(?<!\$)(?<![\w.])(-?\d+\.?\d*|-?\.\d+)(?![\w.])(?!\$)",
)
SKIP_NUMS = {
    "0", "1", "-1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "30", "100", "365", "360",
    "252", "1000",
}
TEXT_FORMULA_RE = re.compile(
    r"\bformula\b|\bequation\b|ctrl\+f|black formula|cell ref|hardcod|\b=sum\b|\b=if\b|every formula row",
    re.I,
)

PRIORITY = {
    "self_referencing_sheet_name": 100,
    "inline_numeric_hardcodes": 85,
    "lowercase_function_names": 60,
    "iferror_or_round_wrappers": 45,
    "text_cells_explaining_formulas": 25,
}

FIX_PATTERNS = {
    "self_referencing_sheet_name": (
        "Drop redundant sheet prefix on same-tab refs (e.g. use C5 instead of Scenarios!C5)."
    ),
    "lowercase_function_names": "Use Excel uppercase function names (SUM, IF, ROUND, IFERROR, etc.).",
    "iferror_or_round_wrappers": (
        "Remove IFERROR/ROUND wrapper; fix upstream precision or number format instead."
    ),
    "inline_numeric_hardcodes": (
        "Replace embedded numeric literal with assumption/input anchor cell reference."
    ),
    "text_cells_explaining_formulas": (
        "Replace meta/formula commentary with filing citation or delete instructional text."
    ),
}


def _is_formula(val) -> bool:
    return isinstance(val, str) and val.startswith("=")


def _normalize_sheet(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip()).lower()


def _extract_formula_nums(formula: str) -> list[str]:
    stripped = re.sub(r'"[^"]*"', '""', formula)
    nums: list[str] = []
    for match in FORMULA_NUM_RE.finditer(stripped):
        num = match.group(1)
        if num not in SKIP_NUMS:
            nums.append(num)
    return nums


def _meaningful_hardcodes(nums: list[str]) -> list[str]:
    meaningful = [
        n
        for n in nums
        if not (n.lstrip("-").replace(".", "").isdigit() and abs(float(n)) <= 10)
    ]
    if not meaningful and nums:
        big = [
            n
            for n in nums
            if n.lstrip("-").replace(".", "").isdigit()
            and len(n.replace(".", "").replace("-", "")) >= 5
        ]
        meaningful = big
    return meaningful


def _classify_formula(sheet: str, formula: str) -> list[str]:
    issues: list[str] = []

    for match in SHEET_REF_RE.finditer(formula):
        ref_sheet = (match.group(1) or match.group(2) or "").strip()
        if _normalize_sheet(ref_sheet) == _normalize_sheet(sheet):
            issues.append("self_referencing_sheet_name")
            break

    for match in FUNC_CALL_RE.finditer(formula):
        fn = match.group(1)
        fn_upper = fn.upper()
        if fn_upper in EXCEL_FUNCTIONS and fn != fn_upper:
            issues.append("lowercase_function_names")
            break

    if IFERROR_ROUND_RE.search(formula):
        issues.append("iferror_or_round_wrappers")

    nums = _meaningful_hardcodes(_extract_formula_nums(formula))
    if nums and sheet in PROJECTION_SHEETS:
        issues.append("inline_numeric_hardcodes")

    return issues


def _candidate_score(formula: str, issues: list[str]) -> int:
    score = sum(PRIORITY.get(issue, 0) for issue in issues)
    if "inline_numeric_hardcodes" in issues:
        nums = _meaningful_hardcodes(_extract_formula_nums(formula))
        if any(len(n.replace(".", "").replace("-", "")) >= 5 for n in nums):
            score += 20
    return score


def _suggested_fix(issues: list[str]) -> str:
    return " ".join(FIX_PATTERNS[issue] for issue in issues)


def audit_workbook(path: Path = WORKBOOK) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)

    formulas_per_sheet: dict[str, int] = {}
    issue_counts: dict[str, int] = {issue: 0 for issue in ISSUE_TYPES}
    all_formulas: list[dict] = []
    text_cells: list[dict] = []
    refactor_candidates: list[dict] = []

    for ws in wb.worksheets:
        formula_count = 0
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if _is_formula(value):
                    formula_count += 1
                    issues = _classify_formula(ws.title, value)
                    entry = {
                        "sheet": ws.title,
                        "cell": cell.coordinate,
                        "formula": value,
                        "issues": issues,
                    }
                    all_formulas.append(entry)
                    for issue in issues:
                        issue_counts[issue] += 1
                    if issues:
                        refactor_candidates.append(
                            {
                                "sheet": ws.title,
                                "cell": cell.coordinate,
                                "formula": value,
                                "issues": issues,
                                "score": _candidate_score(value, issues),
                                "suggested_fix_pattern": _suggested_fix(issues),
                                "reason": "; ".join(issues),
                            }
                        )
                elif isinstance(value, str) and TEXT_FORMULA_RE.search(value):
                    text_entry = {
                        "sheet": ws.title,
                        "cell": cell.coordinate,
                        "text": value,
                        "issues": ["text_cells_explaining_formulas"],
                    }
                    text_cells.append(text_entry)
                    issue_counts["text_cells_explaining_formulas"] += 1
                    refactor_candidates.append(
                        {
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "formula": value,
                            "issues": ["text_cells_explaining_formulas"],
                            "score": PRIORITY["text_cells_explaining_formulas"],
                            "suggested_fix_pattern": FIX_PATTERNS["text_cells_explaining_formulas"],
                            "reason": "text_cells_explaining_formulas",
                        }
                    )

        formulas_per_sheet[ws.title] = formula_count

    refactor_candidates.sort(key=lambda item: (-item["score"], item["sheet"], item["cell"]))
    top_50 = [
        {key: val for key, val in item.items() if key != "score"}
        for item in refactor_candidates[:50]
    ]

    return {
        "workbook": path.name,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_formulas": sum(formulas_per_sheet.values()),
        "formulas_per_sheet": formulas_per_sheet,
        "issue_counts": issue_counts,
        "formulas": all_formulas,
        "text_cells_explaining_formulas": text_cells,
        "top_50_refactor_candidates": top_50,
    }


def main() -> int:
    if not WORKBOOK.exists():
        raise FileNotFoundError(WORKBOOK)

    report = audit_workbook(WORKBOOK)
    with OUTPUT.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)
        handle.write("\n")

    print(f"Wrote {OUTPUT}")
    print(f"Total formulas: {report['total_formulas']}")
    print(f"Issue counts: {json.dumps(report['issue_counts'], indent=2)}")
    print(f"Refactor candidates in top 50: {len(report['top_50_refactor_candidates'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
