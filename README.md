# hopptx

Generate PowerPoint presentations from Calendly sign-up data.

## Setup (one time)

Install UV, a small tool that manages Python for you:

https://docs.astral.sh/uv/getting-started/installation/

## Usage

Open the `hopptx` folder and double-click:

- **Windows:** `run.bat`
- **Mac:** `run.command`

A black terminal window will briefly appear behind the app: that's normal, it's just how Windows launches the program, and you can ignore it (don't close it, or the app will close too). The actual window you'll use looks like a regular program window:

1. **Click "Select Files..." and pick your Excel report(s)**, you can select multiple `.xlsx` files at once.
2. **Pick a start and end date** from the calendar date pickers.
3. **Click "Generate Presentations"**, progress is shown right in the window.
4. When done, the output folder opens automatically with one PowerPoint per company.

If something goes wrong, an error message pops up in plain language. A detailed log file is always saved alongside the output (use the "View Log" button to open it).

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
