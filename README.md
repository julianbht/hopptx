# hopptx

Generate PowerPoint presentations from sign-up report data.

## Prerequisites

- [UV](https://docs.astral.sh/uv/getting-started/installation/) — UV will install the correct Python version automatically if needed.

## Setup

```sh
uv sync
```

## Input

The program reads `.xlsx` Excel files. Each file represents one company and must have:

- **One sheet**, named after the company (the sheet name becomes the company name in the presentation).
- **These exact columns:**

| Column | Type | Description |
|--------|------|-------------|
| User Name | text | Calendly user who owns the event |
| Team | text | Team the user belongs to |
| Invitee Name | text | Full name of the person who signed up |
| Invitee First Name | text | First name |
| Invitee Last Name | text | Last name |
| Invitee Email | text | Email address |
| Event Type Name | text | Name of the training/workshop |
| Start Date & Time | datetime | When the session starts |
| Event Created Date & Time | datetime | When the sign-up was created |
| Canceled | boolean | Whether the sign-up was canceled |

No extra columns are allowed. The column names must match exactly.

Place the `.xlsx` files in the path specified by `input.path` in `config/runs.json`.

## Configuration

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

## Running

```sh
uv run hopptx default
```

Where `default` is the run name defined in `config/runs.json`. You can define multiple runs with different names and configurations.

Shortcut to always run the `default` run:

```sh
uv run hopptx-default
```

## Output

Presentations are written to `output/pptx/{run-name}/YYYY/MM/DD/YYYY-MM-DD_HH-MM-SS/` — one `.pptx` per company, plus a `run.log`.

## Tests

```sh
uv run pytest
```
