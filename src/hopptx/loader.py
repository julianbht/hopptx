import logging
import re
from pathlib import Path

import pandas as pd

from hopptx.schemas.config import RunConfig
from hopptx.schemas.results import Skipped

log = logging.getLogger(__name__)

# Only the columns the statistics are computed from; every other column is ignored.
REQUIRED_COLUMNS = (
    "Event Type Name",  # program name: topics, top programs, categories, sessions
    "Start Date & Time",  # session date: date range filter, first/last attended, sessions
    "Event Created Date & Time",  # first/last sign-up date
    "Canceled",  # withdrawn sign-ups, excluding canceled sign-ups
    "Invitee Email",  # unique attendees
)
DATE_COLUMNS = ("Start Date & Time", "Event Created Date & Time")

_CANCELED_VALUES = {
    "true": True,
    "yes": True,
    "1": True,
    "false": False,
    "no": False,
    "0": False,
}
_MAX_ROWS_IN_MESSAGE = 5

UNKNOWN_COMPANY = "Unknown Company"
# Names Excel gives new sheets (English, German, French, Spanish, Italian, Dutch).
_DEFAULT_SHEET_NAME = re.compile(
    r"(sheet|tabelle|feuil|hoja|foglio|blad)\s*\d*", re.IGNORECASE
)


def _excel_rows(index: pd.Index) -> str:
    """Format DataFrame row labels as the row numbers the user sees in Excel (header is row 1)."""
    rows = [str(i + 2) for i in index[:_MAX_ROWS_IN_MESSAGE]]
    more = len(index) - len(rows)
    return ", ".join(rows) + (f" and {more} more" if more > 0 else "")


def _is_blank(series: pd.Series) -> pd.Series:
    return series.isna() | (series.astype(str).str.strip() == "")


def _parse_canceled(series: pd.Series, file_name: str) -> pd.Series:
    if series.dtype == bool:
        return series
    parsed = series.map(
        lambda value: value if isinstance(value, bool) else _CANCELED_VALUES.get(str(value).strip().casefold())
    )
    unknown = parsed.isna()
    if unknown.any():
        examples = sorted({str(v) for v in series[unknown]})[:_MAX_ROWS_IN_MESSAGE]
        raise ValueError(
            f"'{file_name}': the 'Canceled' column must contain TRUE/FALSE (or Yes/No), "
            f"but row(s) {_excel_rows(series.index[unknown])} contain {examples}."
        )
    return parsed.astype(bool)


def _clean_report(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    """Keep the required columns, drop empty rows, and turn bad cells into clear errors."""
    df = df[list(REQUIRED_COLUMNS)].copy()

    blank = df.apply(_is_blank)
    empty_rows = blank.all(axis=1)
    if empty_rows.any():
        log.info(f"Ignoring {int(empty_rows.sum())} empty row(s) in {file_name}")
        df = df[~empty_rows]
        blank = blank[~empty_rows]

    for column in REQUIRED_COLUMNS:
        if blank[column].any():
            raise ValueError(
                f"'{file_name}': the '{column}' column is empty in row(s) "
                f"{_excel_rows(df.index[blank[column]])}. "
                f"Please fill in these cells or delete the rows."
            )

    for column in DATE_COLUMNS:
        parsed = pd.to_datetime(df[column], errors="coerce")
        invalid = parsed.isna()
        if invalid.any():
            raise ValueError(
                f"'{file_name}': the '{column}' column has values that are not dates in row(s) "
                f"{_excel_rows(df.index[invalid])}."
            )
        df[column] = parsed

    df["Canceled"] = _parse_canceled(df["Canceled"], file_name)
    df["Event Type Name"] = df["Event Type Name"].astype(str).str.strip()
    return df


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

    company = str(xl.sheet_names[0]).strip()  # Company name is the sheet name
    if _DEFAULT_SHEET_NAME.fullmatch(company):
        log.warning(
            f"'{path.name}': sheet name '{company}' is not a company name, "
            f"using '{UNKNOWN_COMPANY}'. Rename the sheet to the company name to fix this."
        )
        company = UNKNOWN_COMPANY
    df = xl.parse(xl.sheet_names[0])
    # Stray spaces in headers are invisible in Excel, so don't let them fail validation.
    df.columns = [str(column).strip() for column in df.columns]

    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(
            f"The Excel file '{path.name}' (first sheet '{company}') is missing required "
            f"column(s): {', '.join(missing)}. "
            f"Required columns: {', '.join(REQUIRED_COLUMNS)}. "
            f"Columns found: {', '.join(df.columns)}."
        )

    df = _clean_report(df, path.name)
    log.info(f"Loaded {len(df)} rows from {path.name} (company: {company})")
    return company, df


def _load_report_files(
    xlsx_files: list[Path],
) -> tuple[list[tuple[str, pd.DataFrame]], list[Skipped]]:
    """Load each file, skipping (and reporting) the ones that fail instead of aborting."""
    results: list[tuple[str, pd.DataFrame]] = []
    skipped: list[Skipped] = []
    for file in xlsx_files:
        # Excel's lock files ("~$name.xlsx") appear next to files that are open in Excel.
        if file.name.startswith("~$"):
            continue
        try:
            results.append(load_report_file(file))
        except Exception as e:
            log.warning(f"Skipping '{file.name}': {e}")
            skipped.append(Skipped(name=file.name, reason=str(e)))

    if skipped:
        log.warning(
            f"Skipped {len(skipped)} file(s) that could not be loaded: "
            f"{', '.join(s.name for s in skipped)}"
        )
    if not results:
        reasons = "\n".join(f"- {s.name}: {s.reason}" for s in skipped)
        raise ValueError(f"None of the Excel files could be loaded.\n{reasons}")

    log.info(f"Loaded {len(results)} report file(s)")
    return results, skipped


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


def load_reports(
    config: RunConfig,
) -> tuple[list[tuple[str, pd.DataFrame]], list[Skipped]]:
    """Load report files based on input config (file, directory, or files).

    Returns the loaded (company, data) pairs and the files that were skipped.
    """
    companies_inline = config.companies_inline

    if config.input.type == "files":
        xlsx_files = [Path(p) for p in (config.input.paths or [])]
        reports, skipped = _load_report_files(xlsx_files)
        return _filter_reports_by_companies(reports, companies_inline), skipped

    input_path = Path(config.input.path)

    if config.input.type == "file":
        if not input_path.exists():
            raise FileNotFoundError(
                f"Input file not found: {input_path}. "
                f"Please check the 'input.path' setting in config/runs.json."
            )
        return _filter_reports_by_companies([load_report_file(input_path)], companies_inline), []

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

        reports, skipped = _load_report_files(xlsx_files)
        return _filter_reports_by_companies(reports, companies_inline), skipped

    else:
        raise ValueError(f"Invalid input type: {config.input.type}")
