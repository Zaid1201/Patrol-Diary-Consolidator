from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

msoFalse = 0
msoTrue = -1
ppSaveAsOpenXMLPresentation = 24


def start_powerpoint(visible: bool = True):
    """Start PowerPoint safely.

    Some Microsoft 365 / Smart App Control configurations reject setting
    Application.Visible = False and raise:
    "Invalid request. Hiding the application window is not allowed."

    PowerPoint is therefore always made visible. The `visible` argument is kept
    for backward compatibility with older callers, but it is intentionally not
    used to hide the application.
    """
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = msoTrue
    except Exception:
        # If PowerPoint is already visible or managed by policy, continue.
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app


def open_presentation(app: Any, path: Path, readonly: bool = True, with_window: bool = False):
    # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
    return app.Presentations.Open(
        str(path.resolve()),
        msoTrue if readonly else msoFalse,
        msoFalse,
        msoTrue if with_window else msoFalse,
    )


def wait_for_clipboard(delay: float = 0.35):
    time.sleep(delay)
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def copy_slide_preserving_media(source_presentation: Any, source_slide_number: int, destination: Any) -> Any:
    """
    Copy a slide through PowerPoint's native clipboard pipeline.

    This is intentionally used instead of Slides.InsertFromFile. InsertFromFile
    can remap embedded image relationships when many presentations contain
    identically named media parts, which is what caused logos to appear inside
    work-order picture placeholders in earlier versions.
    """
    source_slide = source_presentation.Slides(int(source_slide_number))
    last_error: Exception | None = None

    for attempt in range(1, 6):
        before = int(destination.Slides.Count)
        try:
            source_slide.Copy()
            wait_for_clipboard(0.20 + attempt * 0.12)
            pasted_range = destination.Slides.Paste(before + 1)
            wait_for_clipboard(0.15)

            after = int(destination.Slides.Count)
            if after != before + 1:
                raise RuntimeError(f"PowerPoint pasted {after - before} slides instead of one.")

            try:
                return pasted_range.Item(1)
            except Exception:
                return destination.Slides(after)
        except Exception as exc:
            last_error = exc
            # Remove a partial paste before retrying.
            try:
                while int(destination.Slides.Count) > before:
                    destination.Slides(destination.Slides.Count).Delete()
            except Exception:
                pass
            wait_for_clipboard(0.35 * attempt)

    raise RuntimeError(
        f"Could not copy source slide {source_slide_number} after five attempts: {last_error}"
    )


def iter_shapes(container: Any):
    """Yield shapes recursively, including grouped shapes."""
    try:
        count = container.Shapes.Count
        shapes = container.Shapes
    except Exception:
        return

    for i in range(1, count + 1):
        shp = shapes.Item(i)
        yield shp
        try:
            if int(getattr(shp, "Type", 0)) == 6:  # msoGroup
                yield from iter_group_items(shp)
        except Exception:
            pass


def iter_group_items(group_shape: Any):
    try:
        for j in range(1, group_shape.GroupItems.Count + 1):
            child = group_shape.GroupItems.Item(j)
            yield child
            try:
                if int(getattr(child, "Type", 0)) == 6:
                    yield from iter_group_items(child)
            except Exception:
                pass
    except Exception:
        return


def shape_text(shape: Any) -> str:
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return str(shape.TextFrame.TextRange.Text or "")
    except Exception:
        pass
    return ""
