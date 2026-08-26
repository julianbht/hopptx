# hopptx

Generate PowerPoint presentations from sign-up report data.

## Getting Started

### 1. Install UV

UV is a tool that manages Python and project dependencies for you. Install it by following the instructions at:

https://docs.astral.sh/uv/getting-started/installation/

You do not need to install Python separately — UV handles that automatically.

### 2. Open a Terminal

You need a terminal (command line) to run this tool.

**Windows:**
Press `Win + R`, type `cmd`, and press Enter. Or search for "Command Prompt" in the Start menu.

**Mac:**
Press `Cmd + Space`, type `Terminal`, and press Enter.

Then navigate to the project folder. For example, if the project is on your Desktop:

```
cd Desktop/hopptx
```

### 3. Install Dependencies

Run this once (or after updates):

```
uv sync
```

This downloads everything the tool needs to run. It only needs to be done once.

### 4. Prepare Your Data

Place your `.xlsx` Excel files in the folder specified by `input.path` in `config/runs.json` (by default, the `reports/` folder).

Each Excel file represents one company and must have:

- **One sheet**, named after the company (the sheet name becomes the company name in the presentation).
- **The following columns** (extra columns are fine and will be ignored):

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

If any of these columns are missing, the tool will tell you which ones are missing and stop.

### 5. Configure

Edit `config/runs.json` before each run. The key fields:

| Field | What to set |
|-------|-------------|
| `start` / `end` | Date range to filter sign-ups (format: `YYYY-MM-DD`) |
| `input.path` | Path to the directory or file containing the `.xlsx` report(s) |
| `input.type` | `"directory"` (load all `.xlsx` in a folder) or `"file"` (single file) |
| `cost_per_training` | Per-attendee training cost |
| `filename_suffix` | Text appended to each output filename, or `null` for none |
| `companies_inline` | List of company names to generate for, or `null` for all |

The other fields (template, topic grouping, fuzzy thresholds, etc.) are tuned and generally don't need changing.

### 6. Run

```
uv run hopptx default
```

This reads your Excel files, calculates training metrics, and generates one PowerPoint presentation per company.

The output is saved to `output/pptx/default/...` in a timestamped folder. A `run.log` file is also created there with details about what happened.

#### Shortcut

```
uv run hopptx-default
```

Does the same thing — runs the `default` configuration.

## Tests

```
uv run pytest
```
