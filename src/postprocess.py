from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def remove_text_highlights(pptx_path: Path) -> int:
    """Remove accidental DrawingML text highlights without changing text runs."""
    pptx_path = pptx_path.resolve()
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_no_highlight_"))
    replacement = pptx_path.with_suffix(".cleaned.pptx")
    removed = 0

    try:
        with zipfile.ZipFile(pptx_path, "r") as archive:
            archive.extractall(temp_dir)

        for xml_path in (temp_dir / "ppt" / "slides").glob("slide*.xml"):
            tree = ET.parse(xml_path)
            root = tree.getroot()
            changed = False
            for parent in root.iter():
                for child in list(parent):
                    if child.tag == f"{{{A_NS}}}highlight":
                        parent.remove(child)
                        removed += 1
                        changed = True
            if changed:
                tree.write(xml_path, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(replacement, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in temp_dir.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_dir).as_posix())

        replacement.replace(pptx_path)
        return removed
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        try:
            replacement.unlink(missing_ok=True)
        except Exception:
            pass
