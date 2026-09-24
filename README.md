# hopptx

Generate PowerPoint presentations from Calendly sign-up data.

## Setup (one time)

Install UV, a small tool that manages Python for you:

https://docs.astral.sh/uv/getting-started/installation/

## Usage

Open the `hopptx` folder and double-click:

- **Windows:** `run.bat`
- **Mac:** `run.command`

A small black window appears for a moment while the app starts and then closes by itself (the very first start takes a minute or two because it downloads what the app needs; an internet connection is required once). On a Mac, a Terminal window stays open; you can close it once the app is showing.

1. **Click "Select Files..." and pick your Excel report(s)**, you can select multiple `.xlsx` files at once.
2. **Pick a start and end date** from the calendar date pickers.
3. **Click "Generate Presentations"**, progress is shown right in the window.
4. When done, the output folder opens automatically with one PowerPoint per company.

If something goes wrong, an error message pops up. A detailed log file is always saved alongside the output (use the "View Log" button to open it).

## Excel file format

Each `.xlsx` file represents one company. The first sheet name becomes the company name in the presentation. If the sheet still has Excel's default name (e.g. "Sheet1" or "Tabelle1"), the presentation says "Unknown Company" instead; rename the sheet to fix it.

These columns are required (column order doesn't matter; all other columns are ignored):

| Column | Used for |
|--------|----------|
| Event Type Name | Program name: top programs, categories, session count |
| Start Date & Time | Date range filter, first/last attended session, session count |
| Event Created Date & Time | First/last sign-up date |
| Canceled | Withdrawn sign-ups (TRUE/FALSE or Yes/No) |
| Invitee Email | Number of unique attendees |

Completely empty rows are ignored. Rows where a required cell is empty or unreadable (e.g. a date that isn't a date) are left out of the statistics, and a warning in the window and the log names the Excel rows. A file is only skipped if none of its rows can be used. Skipped files and companies are listed when the run finishes.

If the app fails to start, the details are written to `output/hopptx-error.log`.

## Advanced usage

For more control, you can edit `config/runs.json` directly and run:

```
uv run hopptx default
```

See `config/runs.json` for all available options (date range, cost per training, topic grouping, etc.).

## Generate this README as PDF

```
uv run --extra docs python build_readme_pdf.py
```

This creates `README.pdf` in the project folder.

## Tests

```
uv run pytest
```
