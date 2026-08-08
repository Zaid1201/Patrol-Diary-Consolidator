from __future__ import annotations

import re
from typing import Any

from com_utils import iter_shapes, shape_text
from table_filter import filter_table_by_worktype
from text_cleaner import clean_shape_text_preserving_format, clean_table_preserving_format


def clean_slide_text(slide: Any) -> int:
    """Apply minimal contractor-token cleanup while preserving all formatting."""
    changed = 0
    for shape in list(iter_shapes(slide)):
        try:
            if shape.HasTable:
                changed += clean_table_preserving_format(shape.Table)
                continue
        except Exception:
            pass
        changed += clean_shape_text_preserving_format(shape)
    return changed


def filter_slide_tables(slide: Any, worktype: str) -> int:
    deleted = 0
    for shape in list(iter_shapes(slide)):
        try:
            if shape.HasTable:
                deleted += filter_table_by_worktype(shape.Table, worktype)
        except Exception:
            pass
    return deleted


def _direct_slide_shapes(slide: Any):
    try:
        for index in range(1, slide.Shapes.Count + 1):
            yield slide.Shapes.Item(index)
    except Exception:
        return


def set_reporting_title(slide: Any, section: str, numeric_date: str) -> bool:
    """Set only the section heading above a Maximo table to the canonical text."""
    canonical = {
        "PS": f"Patrolling during the reporting date of {numeric_date}",
        "CM": f"Corrective Maintenance during the reporting date of {numeric_date}",
        "FR": f"First response during the reporting date of {numeric_date}",
        "IN": f"IN during the reporting date of {numeric_date}",
    }.get(section.upper())
    if not canonical:
        return False

    candidates: list[tuple[float, Any]] = []
    for shape in _direct_slide_shapes(slide):
        text = shape_text(shape).strip()
        lower = text.lower()
        if not text:
            continue
        if (
            "reporting date" in lower
            or lower.startswith("maximo during")
            or lower.startswith("details of rectified defects")
            or lower.startswith("inspection during")
        ):
            try:
                candidates.append((float(shape.Top), shape))
            except Exception:
                candidates.append((0.0, shape))

    if not candidates:
        return False

    candidates.sort(key=lambda item: item[0])
    shape = candidates[0][1]
    try:
        tr = shape.TextFrame.TextRange
        old = str(tr.Text or "")
        # Section headings are uniformly formatted in the source. Replacing this
        # one short range does not touch the table or any body text.
        tr.Characters(1, max(1, len(old))).Text = canonical
        return True
    except Exception:
        return False


def _normalise_line(text: str) -> str:
    return " ".join(
        text.replace("\xa0", " ").replace("\x0b", " ").replace("–", "-").split()
    ).strip().lower()


def _missing_patrol_score(text: str) -> int:
    """Return evidence that a road/ODO block is genuinely empty.

    A route prefix such as ``QSCHSRN-`` is not considered empty when it is
    followed by a route number. This fixes the August case where the valid
    ``QSCHSRN- 1002`` middle block was removed by an earlier substring rule.
    """
    if not text:
        return 0

    score = 0
    lines = [_normalise_line(line) for line in re.split(r"[\r\n\x0b]+", text)]
    lines = [line for line in lines if line]

    for line in lines:
        if line in {"()", "( )", "-"} or re.fullmatch(r"(?:chec|alc|alcat)?\s*\(\s*\)", line):
            score += 2
            continue

        if "route reference" in line:
            remainder = line.split("route reference", 1)[1]
            # A valid route must contain at least one digit after the label.
            if not re.search(r"\d", remainder):
                score += 2
            continue

        if "length covered" in line:
            # Ignore digits that may occur in the label elsewhere; the line
            # itself needs a distance value.
            if not re.search(r"\d+(?:\.\d+)?", line):
                score += 2
            continue

        if "start" in line and "end" in line:
            if not re.search(r"\d+(?:\.\d+)?", line):
                score += 2
            continue

        if line.startswith("total") or " total:" in f" {line}":
            if not re.search(r"\d+(?:\.\d+)?", line):
                score += 1

    return score


def _horizontal_overlap_values(left_a: float, right_a: float, left_b: float, right_b: float) -> float:
    return max(0.0, min(right_a, right_b) - max(left_a, left_b))


def _vertical_overlap_values(top_a: float, bottom_a: float, top_b: float, bottom_b: float) -> float:
    return max(0.0, min(bottom_a, bottom_b) - max(top_a, top_b))


def delete_empty_patrol_blocks(slide: Any, source_media: Any | None = None) -> int:
    """Delete only road columns that are proven empty.

    The function is deliberately conservative:

    * a valid source image protects its whole column;
    * at least two independent missing-data signals are required; and
    * route prefixes are never treated as empty when followed by a number.

    This allows differently formatted daily reports while preventing populated
    SATURN blocks such as ``QSCHSRN- 1002`` from being removed.
    """
    body_top = 105.0
    evidence_shapes: list[tuple[float, float, int]] = []
    direct_shapes = list(_direct_slide_shapes(slide))

    for shape in direct_shapes:
        text = shape_text(shape)
        score = _missing_patrol_score(text)
        if score <= 0:
            continue
        try:
            top = float(shape.Top)
            if top < body_top:
                continue
            left = float(shape.Left)
            right = left + float(shape.Width)
            evidence_shapes.append((max(0.0, left - 35.0), right + 35.0, score))
        except Exception:
            continue

    if not evidence_shapes:
        return 0

    # Merge adjacent evidence spans into candidate road columns.
    evidence_shapes.sort(key=lambda item: item[0])
    merged: list[list[float]] = []  # left, right, score
    for left, right, score in evidence_shapes:
        if not merged or left > merged[-1][1] + 15.0:
            merged.append([left, right, float(score)])
        else:
            merged[-1][1] = max(merged[-1][1], right)
            merged[-1][2] += score

    # Source PPTX media is authoritative. Never remove a region containing a
    # genuine body image, regardless of how its text is formatted.
    protected_images = []
    if source_media is not None:
        for image in getattr(source_media, "images", ()):
            rect = getattr(image, "rect", None)
            if rect is not None and float(rect.top) >= body_top:
                protected_images.append(rect)

    empty_regions: list[tuple[float, float]] = []
    for left, right, score in merged:
        if score < 2:
            continue

        region_width = max(1.0, right - left)
        protected = False
        for rect in protected_images:
            horizontal = _horizontal_overlap_values(left, right, float(rect.left), float(rect.right))
            # Road images occupy the upper body; horizontal overlap is the key
            # column signal. A 25% threshold tolerates different crop sizes.
            if horizontal >= min(region_width, float(rect.width)) * 0.25:
                protected = True
                break
        if not protected:
            empty_regions.append((left, right))

    if not empty_regions:
        return 0

    deleted = 0
    for shape in reversed(direct_shapes):
        try:
            left = float(shape.Left)
            top = float(shape.Top)
            width = float(shape.Width)
            center = left + width / 2.0
            if top < body_top:
                continue
            if any(region_left <= center <= region_right for region_left, region_right in empty_regions):
                shape.Delete()
                deleted += 1
        except Exception:
            pass
    return deleted

