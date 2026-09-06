"""Shared styling helpers enforcing the GIS color convention:

    BLUE  font -> hard-coded numbers reported by the company (10-K / release)
    BLACK font -> calculations / formulas
    RED   font -> analyst assumptions / inputs

Uses openpyxl. Garamond is applied where available (matches the GIS deck).
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.hyperlink import Hyperlink

FONT_NAME = "Garamond"

BLUE = "0000CC"    # reported
BLACK = "000000"   # formula
RED = "C00000"     # assumption
WHITE = "FFFFFF"
NAVY = "1F2A44"
GREEN = "006100"

# GIS-style palette
DARK = "1F2A44"      # header navy
ACCENT = "C8102E"    # USC-ish cardinal accent
LIGHT = "D9E1F2"     # light blue band
GREY = "F2F2F2"

def font(color=BLACK, bold=False, size=10, italic=False, name=FONT_NAME, underline=None):
    return Font(name=name, color=color, bold=bold, size=size, italic=italic, underline=underline)

def fill(color):
    return PatternFill("solid", fgColor=color)

thin = Side(style="thin", color="BFBFBF")
medium = Side(style="medium", color="000000")
double = Side(style="double", color="000000")

def border(top=None, bottom=None, left=None, right=None):
    return Border(top=top, bottom=bottom, left=left, right=right)

top_border = border(top=thin)
bottom_border = border(bottom=thin)
top_bottom = border(top=thin, bottom=thin)
top_double = border(top=Side(style="thin", color="000000"),
                    bottom=Side(style="double", color="000000"))

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center")
right = Alignment(horizontal="right", vertical="center")
left_indent = Alignment(horizontal="left", vertical="center", indent=1)
left_indent2 = Alignment(horizontal="left", vertical="center", indent=2)

# number formats
NUM = '#,##0;(#,##0)'
NUM1 = '#,##0.0;(#,##0.0)'
PCT = '0.0%'
PCT0 = '0%'
MONEY = '$#,##0.00'
MULT = '0.0x'
EPSFMT = '$#,##0.00'

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def write(ws, cell, value, color=BLACK, bold=False, size=10, numfmt=None,
          align=None, fillc=None, bdr=None, italic=False, name=FONT_NAME):
    c = ws[cell]
    c.value = value
    c.font = font(color=color, bold=bold, size=size, italic=italic, name=name)
    if numfmt:
        c.number_format = numfmt
    if align:
        c.alignment = align
    if fillc:
        c.fill = fill(fillc)
    if bdr:
        c.border = bdr
    return c


def write_link(ws, cell, text, url, color=BLUE, bold=False, size=10, numfmt=None,
               align=None, fillc=None, bdr=None, italic=False, underline="single"):
    c = write(ws, cell, text, color=color, bold=bold, size=size, numfmt=numfmt,
              align=align, fillc=fillc, bdr=bdr, italic=italic)
    c.hyperlink = url
    c.font = font(color=color, bold=bold, size=size, italic=italic, underline=underline)
    return c


def write_reported(ws, cell, value, source_url=None, bold=False, size=10, numfmt=None,
                   align=None, bdr=None):
    """Blue-font reported figure; hyperlinks to source when URL provided."""
    c = write(ws, cell, value, BLUE, bold=bold, size=size, numfmt=numfmt,
              align=align or right, bdr=bdr)
    if source_url:
        c.hyperlink = source_url
        c.font = font(color=BLUE, bold=bold, size=size, underline="single")
    return c


def write_internal_link(ws, cell, text, location, color=BLUE, size=8, italic=True):
    """Clickable link to another cell/tab in this workbook (location e.g. 'WACC!C12')."""
    c = write(ws, cell, text, color, italic=italic, size=size, align=left)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=location)
    c.font = font(color=color, italic=italic, size=size, underline="single")
    return c


def write_assumption_docs(ws, row, justify_col, source_col, key, justify_dict, src_dict,
                          internal_location=None, extra_source_col=None, extra_label=None, extra_url=None):
    """Write ~20-word justification and clickable source link(s) for red assumptions."""
    if key in justify_dict:
        c = ws[f"{justify_col}{row}"]
        c.value = justify_dict[key]
        c.font = font(color=BLACK, italic=True, size=8)
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    src = src_dict.get(key)
    if internal_location:
        label = src[0] if src else "Model cross-reference"
        write_internal_link(ws, f"{source_col}{row}", f"↳ {label}", internal_location)
    elif src:
        label, url = src
        if url:
            # HYPERLINK formula renders reliably as blue underlined text in Excel
            c = ws[f"{source_col}{row}"]
            c.value = f'=HYPERLINK("{url}","↳ {label}")'
            c.font = font(color=BLUE, italic=True, size=8, underline="single")
            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        else:
            write(ws, f"{source_col}{row}", f"↳ {label}", BLACK, italic=True, size=8, align=left)
    if extra_source_col and extra_label:
        if extra_url:
            c = ws[f"{extra_source_col}{row}"]
            c.value = f'=HYPERLINK("{extra_url}","↳ {extra_label}")'
            c.font = font(color=BLUE, italic=True, size=8, underline="single")
        elif extra_label:
            write(ws, f"{extra_source_col}{row}", extra_label, BLACK, italic=True, size=8, align=left)
