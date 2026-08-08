# Patrol Diary Consolidator V4.4 — Terminal Edition

This version creates the consolidated **PowerPoint only**. The local web interface,
PDF conversion, and TTG PDF appendix functionality have been removed.

## What V4.4 fixes

V4.4 fixes the missing middle road diagram on the SATURN night-patrolling slide.
Earlier versions treated any text containing a route prefix such as `QSCHSRN-` as
an empty block. This incorrectly removed a valid block such as `QSCHSRN- 1002`.

The cleanup is now source-aware and conservative:

- A road column containing a real source image is never removed.
- A route reference is considered empty only when no route number follows it.
- Multiple missing-data signals are required before a road column is deleted.
- Truly empty columns from older reports are still removed.

## Requirements

- Windows 10 or Windows 11
- Microsoft PowerPoint / Microsoft 365 desktop application
- Python 3.10 or later

## First-time setup

Open PowerShell in this folder and run:

```powershell
python -m pip install -r requirements.txt
```

## Input files

Place exactly four `.pptx` files in the `input` folder. Their filenames must
contain these identifiers:

- `CHEC`
- `ALCAT`
- `SATURN`
- `TTG`

Do not put an old consolidated PowerPoint in the input folder.

Example filenames:

```text
PATROL DIARY 02-08-2026-CHEC.pptx
Patrol Diary - 2 August 2026-ALCAT.pptx
PATROL DIARY - 02-08-2026-SATURN.pptx
Patrol Diary-02-08-2026-TTG.pptx
```

The report date is read from the filenames. Use either `DD-MM-YYYY` or a written
English date such as `2 August 2026`.

## Run the program

From the project folder:

```powershell
python run.py
```

You can also run the main module directly:

```powershell
python src\main.py
```

The consolidated file is saved in `output` as:

```text
Patrol Diary Consolidated - DD-MM-YYYY.pptx
```

## Use different folders

```powershell
python src\main.py --input-dir "C:\DailyInput" --output-dir "C:\DailyOutput"
```

## Important notes

- PowerPoint opens visibly while the report is being created. Do not close it.
- Close any existing output PowerPoint with the same filename before running.
- Temporary Office files beginning with `~$` are ignored automatically.
- The planner identifies slides from their content, table work types, and section
  order rather than relying only on fixed slide numbers.
