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
        raise FileNotFoundError(f"Report file not found: {path}")

    xl = pd.ExcelFile(path)
    if not xl.sheet_names:
        raise ValueError(f"No sheets found in {path}")

    company = xl.sheet_names[0]  # Company name is the sheet name
    df = xl.parse(company)

    # Validate columns match expected schema
    actual_cols = set(df.columns)

    missing = EXPECTED_COLUMNS - actual_cols
    extra = actual_cols - EXPECTED_COLUMNS

    if missing or extra:
        error_msg = f"Column mismatch in {path.name}."
        if missing:
            error_msg += f"\n  Missing columns: {sorted(missing)}"
        if extra:
            error_msg += f"\n  Extra columns: {sorted(extra)}"
        raise ValueError(error_msg)

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
            raise FileNotFoundError(f"Input file not found: {input_path}")
        return _filter_reports_by_companies([load_report_file(input_path)], companies_inline)

    elif config.input.type == "directory":
        if not input_path.exists():
            raise FileNotFoundError(f"Input directory not found: {input_path}")

        xlsx_files = sorted(input_path.glob("*.xlsx"))
        if not xlsx_files:
            raise ValueError(f"No .xlsx files found in {input_path}")

        results = []
        for file in xlsx_files:
            results.append(load_report_file(file))

        log.info(f"Loaded {len(results)} report files from {input_path}")
        return _filter_reports_by_companies(results, companies_inline)

    else:
        raise ValueError(f"Invalid input type: {config.input.type}")
