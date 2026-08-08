from __future__ import annotations

import re
from typing import Any

# Conservative business-rule removals. Route codes such as QSCHPRD and QSALPRD
# are protected because the patterns require a non-alphanumeric boundary.
_PATTERNS = [
    re.compile(r"(?i)(?<![A-Za-z0-9])(?:CHEC|ALCAT|SATURN|TTG|ALC)(?![A-Za-z0-9])[ \t]*"),
    re.compile(r"(?i)(?<![A-Za-z0-9])(?:CH|AL)[ \t]*-[ \t]*"),
    re.compile(r"(?i)(?<![A-Za-z0-9])(?:CH|AL)(?![A-Za-z0-9])[ \t]*"),
]


def replacement_spans(text: str) -> list[tuple[int, int, str]]:
    """Return non-overlapping text replacements, ordered from end to start."""
    if not text:
        return []

    candidates: list[tuple[int, int, str]] = []
    occupied: list[tuple[int, int]] = []
    for pattern in _PATTERNS:
        for match in pattern.finditer(text):
            start, end = match.span()
            if any(not (end <= a or start >= b) for a, b in occupied):
                continue
            occupied.append((start, end))
            candidates.append((start, end - start, ""))

    return sorted(candidates, key=lambda item: item[0], reverse=True)


def clean_text_range_preserving_format(text_range: Any) -> int:
    """
    Remove contractor tokens by editing only the matching character ranges.

    Reassigning TextRange.Text resets runs, fonts, paragraph spacing, highlight,
    and autofit behavior. Character-range edits preserve the source formatting.
    """
    try:
        text = str(text_range.Text or "")
    except Exception:
        return 0

    changed = 0
    for start_0, length, replacement in replacement_spans(text):
        try:
            text_range.Characters(start_0 + 1, length).Text = replacement
            changed += 1
        except Exception:
            pass
    return changed


def clean_shape_text_preserving_format(shape: Any) -> int:
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return clean_text_range_preserving_format(shape.TextFrame.TextRange)
    except Exception:
        pass
    return 0


def clean_table_preserving_format(table: Any) -> int:
    changed = 0
    try:
        for row in range(1, table.Rows.Count + 1):
            for col in range(1, table.Columns.Count + 1):
                tr = table.Cell(row, col).Shape.TextFrame.TextRange
                changed += clean_text_range_preserving_format(tr)
    except Exception:
        pass
    return changed
