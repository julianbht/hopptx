# hopptx

Generate PowerPoint presentations from Calendly sign-up data.

## Setup (one time)

Install UV — a small tool that manages Python for you:

https://docs.astral.sh/uv/getting-started/installation/

## Usage

Open the `hopptx` folder and double-click:

- **Windows:** `run.bat`
- **Mac:** `run.command`

A window opens where you:

1. **Click "Browse..." and pick the folder** containing your Excel reports — no need to select individual files, every `.xlsx` file in the folder is picked up automatically.
2. **Pick a start and end date** from the calendar date pickers.
3. **Click "Generate Presentations"** — progress is shown right in the window.
4. When done, the output folder opens automatically with one PowerPoint per company.

If something goes wrong, an error message pops up in plain language. A detailed log file is always saved alongside the output.

## Excel file format

Each `.xlsx` file represents one company. The first sheet name becomes the company name in the presentation.

The following columns are required (extra columns are fine and will be ignored):

| Column | Description |
|--------|-------------|
| User Name | Calendly user who owns the event |
| Team | Team the user belongs to |
| Invitee Name | Full name of the person who signed up |
| Invitee First Name | First name |
| Invitee Last Name | Last name |
| Invitee Email | Email address |
| Event Type Name | Name of the training/workshop |
| Start Date & Time | When the session starts |
| Event Created Date & Time | When the sign-up was created |
| Canceled | Whether the sign-up was canceled |

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
