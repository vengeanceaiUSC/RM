"""Restore Excel outline (+/-) controls for collapsible column/row groups."""
from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter

# Contiguous column groups (start_col, end_col)
SHEET_COL_GROUPS: dict[str, list[tuple[str, str]]] = {
    "Revenue Drivers": [("J", "K")],
}

# Row groups: (start_row, end_row, summary_below)
SHEET_ROW_GROUPS: dict[str, list[tuple[int, int, bool]]] = {
    "DCF": [(20, 32, False)],
}


def enable_outline_symbols(ws) -> None:
    ws.sheet_view.showOutlineSymbols = True


def clear_sheet_outlines(ws) -> None:
    """Remove all row/column outline metadata from a worksheet."""
    for dim in ws.column_dimensions.values():
        dim.outline_level = 0
        dim.hidden = False
        dim.collapsed = False
    for dim in ws.row_dimensions.values():
        dim.outline_level = 0
        dim.hidden = False
        dim.collapsed = False


def group_columns(
    ws,
    start_col: str,
    end_col: str,
    *,
    hidden: bool = True,
    outline_level: int = 1,
    summary_right: bool = True,
) -> None:
    ws.column_dimensions.group(start_col, end_col, outline_level=outline_level, hidden=hidden)
    ws.sheet_properties.outlinePr.summaryRight = summary_right
    ws.sheet_properties.outlinePr.applyStyles = True

    end_idx = column_index_from_string(end_col)
    if summary_right:
        summary_col = get_column_letter(end_idx + 1)
    else:
        summary_col = get_column_letter(column_index_from_string(start_col) - 1)
    ws.column_dimensions[summary_col].collapsed = True


def group_rows(
    ws,
    start_row: int,
    end_row: int,
    *,
    hidden: bool = True,
    outline_level: int = 1,
    summary_below: bool = True,
) -> None:
    ws.row_dimensions.group(start_row, end_row, outline_level=outline_level, hidden=hidden)
    ws.sheet_properties.outlinePr.summaryBelow = summary_below
    ws.sheet_properties.outlinePr.applyStyles = True

    summary_row = end_row + 1 if summary_below else start_row - 1
    if summary_row >= 1:
        ws.row_dimensions[summary_row].collapsed = True


def remove_outline_groups(wb) -> int:
    """Strip all Excel row/column outline levels and hide +/- controls."""
    cleared = 0
    for ws in wb.worksheets:
        touched = False
        for col, dim in ws.column_dimensions.items():
            if dim.outline_level or dim.hidden or dim.collapsed:
                dim.outline_level = 0
                dim.hidden = False
                dim.collapsed = False
                touched = True
        for row, dim in ws.row_dimensions.items():
            if dim.outline_level or dim.hidden or dim.collapsed:
                dim.outline_level = 0
                dim.hidden = False
                dim.collapsed = False
                touched = True
        ws.sheet_view.showOutlineSymbols = False
        if ws.sheet_properties.outlinePr is not None:
            ws.sheet_properties.outlinePr.showOutlineSymbols = False
        if touched:
            cleared += 1
    return cleared


def restore_outline_groups(wb, *, col_hidden: bool = True) -> int:
    """Re-apply outline metadata and show +/- controls on every grouped sheet."""
    for ws in wb.worksheets:
        clear_sheet_outlines(ws)

    restored = 0
    grouped_sheets = set(SHEET_COL_GROUPS) | set(SHEET_ROW_GROUPS)
    for sheet_name in grouped_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        enable_outline_symbols(ws)
        for start, end in SHEET_COL_GROUPS.get(sheet_name, []):
            group_columns(ws, start, end, hidden=col_hidden)
        for start, end, summary_below in SHEET_ROW_GROUPS.get(sheet_name, []):
            group_rows(ws, start, end, hidden=col_hidden, summary_below=summary_below)
        restored += 1

    for ws in wb.worksheets:
        enable_outline_symbols(ws)
    return restored
