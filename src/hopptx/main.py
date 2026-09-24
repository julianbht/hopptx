import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Annotated

import typer
from rich.logging import RichHandler

from hopptx.schemas.config import RunConfig, load_runs
from hopptx.loader import load_reports
from hopptx.processor import compute_metrics
from hopptx.generator import generate_presentation
from hopptx.paths import OUTPUT_DIR, RUNS_FILE
from hopptx.schemas.results import GenerationResult, Skipped

log = logging.getLogger(__name__)


def _setup_logging(output_dir: Path) -> logging.FileHandler:
    """Configure root logger with Rich console + file handler in output_dir."""
    output_dir.mkdir(parents=True, exist_ok=True)
    file_handler = logging.FileHandler(output_dir / "run.log")
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s %(levelname)-8s %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[RichHandler(show_time=False, show_path=False, rich_tracebacks=True), file_handler],
    )
    return file_handler


app = typer.Typer()


def _complete_run_name(incomplete: str) -> list[str]:
    try:
        runs = json.loads(RUNS_FILE.read_text())
        return [r["name"] for r in runs if r["name"].startswith(incomplete)]
    except Exception:
        return []


# Characters Windows forbids in file names; Excel sheet names may still contain some of them.
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def _output_path(output_dir: Path, company: str, filename_suffix: str | None) -> Path:
    stem = _INVALID_FILENAME_CHARS.sub("_", company).strip(" .") or "presentation"
    if filename_suffix is not None:
        stem = f"{stem} {filename_suffix}"
    path = output_dir / f"{stem}.pptx"
    # Two files with the same sheet name must not overwrite each other.
    counter = 2
    while path.exists():
        path = output_dir / f"{stem} ({counter}).pptx"
        counter += 1
    return path


def generate_presentations(run: RunConfig, output_dir: Path) -> GenerationResult:
    """Load reports, compute metrics, and generate one PowerPoint presentation per company."""
    reports, skipped_files = load_reports(run)
    log.info(f"Loaded {len(reports)} report(s)")

    result = GenerationResult(skipped=list(skipped_files))
    for company, df in reports:
        log.info(f"\n{'=' * 60}")
        log.info(f"Processing: {company}")
        log.info(f"{'=' * 60}")
        try:
            metrics = compute_metrics(
                df,
                run.start,
                run.end,
                run.exclude_canceled,
                run.cost_per_training,
                run.evaluation_rate_percentage,
                run.top_topics_count,
                run.top_topics_hard_cutoff,
                run.group_topics,
                run.clean_topic_names,
                run.topic_fuzzy_match_threshold,
                run.category_reference,
                run.category_fuzzy_match_threshold,
                strip_company_prefix=run.strip_company_prefix,
                company=company,
            )
            output_path = _output_path(output_dir, company, run.filename_suffix)
            log.info(f"Generating presentation: {output_path}")
            generate_presentation(
                company,
                run.start,
                run.end,
                metrics,
                output_path,
                run.top_topics_count,
                run.max_programs_per_table_page,
                run.single_page_soft_overflow,
                run.topic_max_chars,
                run.topic_truncation_suffix,
                template_path=run.template_path,
            )
            result.written.append(output_path)
            log.info(f"Done: {company}")
        except Exception as e:
            log.error(f"Skipping {company}: {e}")
            result.skipped.append(Skipped(name=company, reason=str(e)))

    if not result.written:
        log.warning("No presentations were produced.")

    return result


def _execute_run(name: str) -> None:
    runs = load_runs()
    matches = [r for r in runs if r.name == name]
    if not matches:
        raise SystemExit(
            f"No run named '{name}' found in {RUNS_FILE}. "
            f"Available runs: {[r.name for r in runs]}."
        )
    run = matches[0]

    now = datetime.now()
    output_dir = (
        OUTPUT_DIR
        / "pptx"
        / run.name
        / now.strftime("%Y")
        / now.strftime("%m")
        / now.strftime("%d")
        / now.strftime("%Y-%m-%d_%H-%M-%S")
    )
    file_handler = _setup_logging(output_dir)

    log.info(f"[bold]{run.name}[/bold]", extra={"markup": True})

    result = generate_presentations(run, output_dir)
    log.info(
        f"Done — {len(result.written)} file(s) written to {output_dir}, "
        f"{len(result.skipped)} skipped"
    )

    logging.root.removeHandler(file_handler)
    file_handler.close()


@app.command()
def main(
    run_name: Annotated[
        str,
        typer.Argument(
            help="Run name as defined in config/runs.json",
            autocompletion=_complete_run_name,
        ),
    ],
) -> None:
    _execute_run(run_name)


def run_pptx_generator() -> None:
    """Shortcut to run the 'default' run."""
    _execute_run("default")
