from __future__ import annotations

from typing import Any, Optional
from text_cleaner import clean_table_preserving_format


def _cell_text(table: Any, row: int, col: int) -> str:
    try:
        return str(table.Cell(row, col).Shape.TextFrame.TextRange.Text or "").strip()
    except Exception:
        return ""


def _norm_header(value: str) -> str:
    return "".join(ch.lower() for ch in value if ch.isalnum())


def find_worktype_column(table: Any) -> Optional[int]:
    try:
        rows = int(table.Rows.Count)
        cols = int(table.Columns.Count)
    except Exception:
        return None

    for row in range(1, min(rows, 3) + 1):
        for col in range(1, cols + 1):
            if "worktype" in _norm_header(_cell_text(table, row, col)):
                return col
    return None


def first_data_row(table: Any) -> int:
    try:
        rows = int(table.Rows.Count)
        cols = int(table.Columns.Count)
    except Exception:
        return 2

    for row in range(1, min(rows, 3) + 1):
        headers = [_norm_header(_cell_text(table, row, col)) for col in range(1, cols + 1)]
        if any("worktype" in header for header in headers):
            return row + 1
    return 2


def row_is_empty(table: Any, row: int) -> bool:
    try:
        cols = int(table.Columns.Count)
    except Exception:
        return True
    values = [_cell_text(table, row, col).replace("\xa0", " ").strip() for col in range(1, cols + 1)]
    return all(not value for value in values)


def filter_table_by_worktype(table: Any, keep_worktype: str) -> int:
    """Delete nonmatching data rows without rewriting any retained cell text."""
    work_col = find_worktype_column(table)
    if not work_col:
        clean_table_preserving_format(table)
        return 0

    start = first_data_row(table)
    keep = keep_worktype.strip().upper()
    deleted = 0

    try:
        for row in range(int(table.Rows.Count), start - 1, -1):
            value = _cell_text(table, row, work_col).upper().strip()
            if row_is_empty(table, row) or value != keep:
                table.Rows(row).Delete()
                deleted += 1
    finally:
        clean_table_preserving_format(table)

    return deleted
