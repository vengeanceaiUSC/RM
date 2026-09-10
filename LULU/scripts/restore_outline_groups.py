"""Restore Excel outline (+/-) controls for collapsible column/row groups."""
from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter

# Column groups: (start_col, end_col)
SHEET_COL_GROUPS: dict[str, list[tuple[str, str]]] = {
    "WACC": [("B", "D")],
    "Scenarios": [("B", "D")],
    "NOPAT Bridge": [("B", "D")],
    "Comps": [("B", "D")],
    "Revenue Drivers": [("I", "L")],
    "DCF": [("B", "D"), ("L", "M")],
}

# Row groups: (start_row, end_row)
SHEET_ROW_GROUPS: dict[str, list[tuple[int, int]]] = {
    "DCF": [(20, 32)],
    "NOPAT Bridge": [(15, 41)],
}


def enable_outline_symbols(ws) -> None:
    ws.sheet_view.showOutlineSymbols = True


def group_columns(
    ws,
    start_col: str,
    end_col: str,
    *,
    hidden: bool = True,
    outline_level: int = 1,
) -> None:
    start = column_index_from_string(start_col)
    end = column_index_from_string(end_col)
    if end < start:
        return
    for i in range(start, end + 1):
        col = get_column_letter(i)
        dim = ws.column_dimensions[col]
        dim.outline_level = outline_level
        dim.hidden = hidden
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.sheet_properties.outlinePr.applyStyles = True


def group_rows(
    ws,
    start_row: int,
    end_row: int,
    *,
    hidden: bool = True,
    outline_level: int = 1,
) -> None:
    if end_row < start_row:
        return
    ws.row_dimensions.group(start_row, end_row, outline_level=outline_level, hidden=hidden)
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_properties.outlinePr.applyStyles = True


def restore_outline_groups(wb, *, col_hidden: bool = True) -> int:
    """Re-apply outline metadata and show +/- controls on every grouped sheet."""
    restored = 0
    for sheet_name, col_ranges in SHEET_COL_GROUPS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        enable_outline_symbols(ws)
        for start, end in col_ranges:
            group_columns(ws, start, end, hidden=col_hidden)
        for start, end in SHEET_ROW_GROUPS.get(sheet_name, []):
            group_rows(ws, start, end, hidden=col_hidden)
        restored += 1

    for ws in wb.worksheets:
        enable_outline_symbols(ws)
    return restored
