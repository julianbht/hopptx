import logging
from pathlib import Path

import pandas as pd

from hopptx.schemas.config import RunConfig

log = logging.getLogger(__name__)

EXPECTED_COLUMNS = {
    "User Name",
    "Team",
    "Invitee Name",
    "Invitee First Name",
    "Invitee Last Name",
    "Invitee Email",
    "Event Type Name",
    "Start Date & Time",
    "Event Created Date & Time",
    "Canceled",
}


def load_report_file(path: Path) -> tuple[str, pd.DataFrame]:
    """Load a single .xlsx report file and extract company name from sheet name."""
    if not path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {path}. "
            f"Please check that the file exists and the path in config/runs.json is correct."
        )

    xl = pd.ExcelFile(path)
    if not xl.sheet_names:
        raise ValueError(
            f"The Excel file '{path.name}' appears to be empty (no sheets found). "
            f"Please make sure it contains at least one sheet with your data."
        )

    company = xl.sheet_names[0]  # Company name is the sheet name
    df = xl.parse(company)

    # Validate that all required columns are present (extra columns are ignored)
    actual_cols = set(df.columns)
    missing = EXPECTED_COLUMNS - actual_cols

    if missing:
        raise ValueError(
            f"The Excel file '{path.name}' is missing required column(s): "
            f"{', '.join(sorted(missing))}. "
            f"Please make sure your file contains all of these columns: "
            f"{', '.join(sorted(EXPECTED_COLUMNS))}."
        )

    extra = actual_cols - EXPECTED_COLUMNS
    if extra:
        log.info(f"Ignoring extra column(s) in {path.name}: {sorted(extra)}")

    log.info(f"Loaded {len(df)} rows from {path.name} (company: {company})")
    return company, df


def _filter_reports_by_companies(
    reports: list[tuple[str, pd.DataFrame]],
    companies_inline: list[str] | None,
) -> list[tuple[str, pd.DataFrame]]:
    """Filter loaded reports to an inline allow-list of company names."""
    if companies_inline is None:
        return reports

    wanted = {name.strip().casefold() for name in companies_inline}
    filtered = [(company, df) for company, df in reports if company.strip().casefold() in wanted]
    found = {company.strip().casefold() for company, _ in reports}
    missing = [name for name in companies_inline if name.strip().casefold() not in found]

    if missing:
        raise ValueError(
            f"Requested companies not found in loaded reports: {missing}"
        )

    if not filtered:
        raise ValueError("No reports matched companies_inline")

    log.info(
        f"Filtered reports to {len(filtered)} company(ies) via companies_inline: "
        f"{[company for company, _ in filtered]}"
    )
    return filtered


def load_reports(config: RunConfig) -> list[tuple[str, pd.DataFrame]]:
    """Load report files based on input config (file or directory)."""
    input_path = Path(config.input.path)
    companies_inline = config.companies_inline

    if config.input.type == "file":
        if not input_path.exists():
            raise FileNotFoundError(
                f"Input file not found: {input_path}. "
                f"Please check the 'input.path' setting in config/runs.json."
            )
        return _filter_reports_by_companies([load_report_file(input_path)], companies_inline)

    elif config.input.type == "directory":
        if not input_path.exists():
            raise FileNotFoundError(
                f"Input directory not found: {input_path}. "
                f"Please check the 'input.path' setting in config/runs.json."
            )

        xlsx_files = sorted(input_path.glob("*.xlsx"))
        if not xlsx_files:
            raise ValueError(
                f"No .xlsx files found in '{input_path}'. "
                f"Make sure the folder contains Excel files with the .xlsx extension."
            )

        results = []
        skipped = []
        for file in xlsx_files:
            try:
                results.append(load_report_file(file))
            except (ValueError, Exception) as e:
                log.warning(f"Skipping '{file.name}': {e}")
                skipped.append(file.name)

        if skipped:
            log.warning(
                f"Skipped {len(skipped)} file(s) that could not be loaded: "
                f"{', '.join(skipped)}"
            )

        if not results:
            raise ValueError(
                f"None of the .xlsx files in '{input_path}' could be loaded. "
                f"Make sure your Excel files have the required columns."
            )

        log.info(f"Loaded {len(results)} report files from {input_path}")
        return _filter_reports_by_companies(results, companies_inline)

    else:
        raise ValueError(f"Invalid input type: {config.input.type}")
