from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from com_utils import (
    copy_slide_preserving_media,
    open_presentation,
    ppSaveAsOpenXMLPresentation,
    start_powerpoint,
)
from media_restore import (
    cleanup_empty_picture_slots,
    ensure_standard_logos,
    extract_slide_media,
    extract_standard_logos,
)
from planner import build_slide_plan
from postprocess import remove_text_highlights
from ppt_actions import (
    clean_slide_text,
    delete_empty_patrol_blocks,
    filter_slide_tables,
    set_reporting_title,
)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"
LOG_DIR = PROJECT_ROOT / "logs"

SOURCE_KEYWORDS = {
    "ALCAT": ["ALCAT"],
    "SATURN": ["SATURN"],
    "CHEC": ["CHEC"],
    "TTG": ["TTG"],
}


def log(message: str):
    print(message, flush=True)


def detect_sources(input_dir: Path) -> dict[str, Path]:
    files = [path for path in input_dir.glob("*.pptx") if not path.name.startswith("~$")]
    result: dict[str, Path] = {}

    for key, keywords in SOURCE_KEYWORDS.items():
        matches: list[Path] = []
        for path in files:
            name = path.name.upper()
            if "CONSOLIDATED" in name:
                continue
            if any(keyword.upper() in name for keyword in keywords):
                matches.append(path)
        if matches:
            matches.sort(key=lambda item: (len(item.name), item.name.lower()))
            result[key] = matches[0]

    missing = [key for key in SOURCE_KEYWORDS if key not in result]
    if missing:
        raise FileNotFoundError(
            f"Missing source files: {', '.join(missing)}. "
            f"Use CHEC, ALCAT, SATURN, and TTG in the filenames and place them in {input_dir}."
        )
    return result


def extract_report_dates(files: dict[str, Path]) -> tuple[str, str]:
    joined = " ".join(path.name for path in files.values())
    match = re.search(r"(\d{1,2})[-_/ ](\d{1,2})[-_/ ](20\d{2})", joined)
    if match:
        day, month, year = map(int, match.groups())
        date = datetime(year, month, day)
        return date.strftime("%d %B %Y").lstrip("0"), date.strftime("%d-%m-%Y")

    match = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(20\d{2})", joined)
    if match:
        date = datetime.strptime(
            f"{int(match.group(1))} {match.group(2)} {match.group(3)}",
            "%d %B %Y",
        )
        return date.strftime("%d %B %Y").lstrip("0"), date.strftime("%d-%m-%Y")

    now = datetime.now()
    return now.strftime("%d %B %Y").lstrip("0"), now.strftime("%d-%m-%Y")


def _replace_range(text_range: Any, start_0: int, length: int, replacement: str) -> bool:
    try:
        text_range.Characters(start_0 + 1, length).Text = replacement
        return True
    except Exception:
        return False


def update_intro_date(slide: Any, long_date: str):
    """Replace only the date characters; never consume the Contractor line."""
    from com_utils import iter_shapes

    date_patterns = [
        re.compile(r"(?i)(D\s*a\s*t\s*e\s*:\s*)(\d{1,2}\s+[A-Za-z]+\s+20\d{2})"),
        re.compile(r"(?i)(Date\s*[–-]\s*)(\d{1,2}[/-]\d{1,2}[/-]20\d{2})"),
    ]

    for shape in iter_shapes(slide):
        try:
            if not (shape.HasTextFrame and shape.TextFrame.HasText):
                continue
            text_range = shape.TextFrame.TextRange
            text = str(text_range.Text or "")
        except Exception:
            continue

        for pattern in date_patterns:
            match = pattern.search(text)
            if match:
                _replace_range(text_range, match.start(2), len(match.group(2)), long_date)
                return


def build_consolidated(
    input_dir: Path | None = None,
    output_dir: Path | None = None,
    visible: bool = True,
) -> Path:
    """Create the consolidated Patrol Diary PowerPoint."""
    input_dir = (input_dir or DEFAULT_INPUT_DIR).resolve()
    output_dir = (output_dir or DEFAULT_OUTPUT_DIR).resolve()
    output_dir.mkdir(exist_ok=True, parents=True)
    LOG_DIR.mkdir(exist_ok=True, parents=True)

    sources = detect_sources(input_dir)
    long_date, numeric_date = extract_report_dates(sources)
    output_path = output_dir / f"Patrol Diary Consolidated - {numeric_date}.pptx"
    slide_plan = build_slide_plan(sources)

    log("Detected sources:")
    for key, path in sources.items():
        log(f"  {key}: {path.name}")
    log(f"Report date: {long_date} ({numeric_date})")
    log(f"Planned output slides: {len(slide_plan)}")

    # This is the V4 slide-generation pipeline. It has deliberately not been
    # replaced by the V5 slide transformations.
    temp_root = Path(tempfile.mkdtemp(prefix="patrol_diary_v4_"))
    app = start_powerpoint(visible=visible)
    opened: dict[str, Any] = {}
    destination = None

    try:
        for key, path in sources.items():
            opened[key] = open_presentation(app, path, readonly=True, with_window=visible)

        logos = extract_standard_logos(sources["CHEC"], temp_root, slide_number=2)
        logo_media = extract_slide_media(sources["CHEC"], 2, temp_root)

        destination = app.Presentations.Add()
        destination.PageSetup.SlideWidth = opened["CHEC"].PageSetup.SlideWidth
        destination.PageSetup.SlideHeight = opened["CHEC"].PageSetup.SlideHeight
        try:
            while destination.Slides.Count > 0:
                destination.Slides(1).Delete()
        except Exception:
            pass

        for output_number, step in enumerate(slide_plan, start=1):
            source_key = step["source"]
            source_slide = int(step["slide"])
            source_path = sources[source_key]
            log(f"{output_number:02d}: copy {source_key} slide {source_slide}")

            new_slide = copy_slide_preserving_media(
                opened[source_key], source_slide, destination
            )

            # V4 behavior: remove only placeholders proven empty in the source.
            media = extract_slide_media(source_path, source_slide, temp_root)
            empty_removed = cleanup_empty_picture_slots(new_slide, media)
            if empty_removed:
                log(f"    removed {empty_removed} empty-slot shapes")

            if output_number == 1:
                update_intro_date(new_slide, long_date)
            else:
                ensure_standard_logos(
                    new_slide,
                    logos,
                    slide_width=logo_media.slide_width,
                )

            worktype = step.get("worktype")
            if worktype:
                deleted_rows = filter_slide_tables(new_slide, worktype)
                set_reporting_title(new_slide, step.get("section", worktype), numeric_date)
                if deleted_rows:
                    log(f"    filtered {deleted_rows} non-{worktype} table rows")

            if step.get("clean", True):
                clean_slide_text(new_slide)

            if step.get("delete_empty_patrol_blocks"):
                extra_removed = delete_empty_patrol_blocks(new_slide, source_media=media)
                if extra_removed:
                    log(f"    removed {extra_removed} malformed empty-road shapes")

        destination.SaveAs(str(output_path.resolve()), ppSaveAsOpenXMLPresentation)
        log(f"PowerPoint saved: {output_path}")

    finally:
        try:
            if destination is not None:
                destination.Close()
        except Exception:
            pass
        for presentation in opened.values():
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        shutil.rmtree(temp_root, ignore_errors=True)

    highlights = remove_text_highlights(output_path)
    log(f"Removed accidental text highlights: {highlights}")

    log(f"Finished PowerPoint: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Create a consolidated Patrol Diary PowerPoint using Microsoft 365."
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help="Folder containing the CHEC, ALCAT, SATURN, and TTG PPTX files.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Folder where the consolidated PPTX will be saved.",
    )
    args = parser.parse_args()

    try:
        build_consolidated(
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            visible=True,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()
