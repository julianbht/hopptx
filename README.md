# hopptx

Generate PowerPoint presentations from Calendly sign-up data.

## Setup (one time)

Install UV — a small tool that manages Python for you:

https://docs.astral.sh/uv/getting-started/installation/

## Usage

Open the `hopptx` folder and double-click:

- **Windows:** `run.bat`
- **Mac:** `run.command`

The tool will:

1. **Ask you to pick your Excel file(s)** — a file picker window will open.
2. **Ask for a date range** — type a start and end date (e.g. `2026-04-01`).
3. **Generate the presentations** — one PowerPoint per company.
4. **Open the output folder** — your reports are ready.

If something goes wrong, the tool will show an error message in plain language. A detailed log file is always saved alongside the output.

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
