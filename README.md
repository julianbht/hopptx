# hopptx

Generate PowerPoint presentations from Calendly sign-up data.

## Setup (one time)

### 1. Install UV

UV is a small tool that manages Python for you. Install it from:

https://docs.astral.sh/uv/getting-started/installation/

You do not need to install Python separately — UV handles that automatically.

### 2. Create a desktop shortcut

Open the `hopptx` folder in a terminal:

**Windows:** Right-click the `hopptx` folder and select **"Open in Terminal"**.

**Mac:** Right-click the `hopptx` folder and select **"New Terminal at Folder"** (you may need to enable this in System Settings > Keyboard > Shortcuts > Services).

Then paste this command and press Enter:

```
uv run hopptx-setup-shortcut
```

This creates a shortcut on your Desktop. From now on, just double-click it.

## Usage

Double-click the **hopptx** shortcut on your Desktop (or run `uv run hopptx-easy` from a terminal in the project folder).

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

## Tests

```
uv run pytest
```
