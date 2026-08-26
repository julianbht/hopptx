import json
import logging
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


def generate_presentations(run: RunConfig, output_dir: Path) -> list[Path]:
    """
    Load reports, compute metrics, and generate PowerPoint presentations.
    Returns the list of paths that were written.
    """
    reports = load_reports(run)
    log.info(f"Loaded {len(reports)} report(s)")

    written = []
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
            if run.filename_suffix is None:
                filename = f"{company}.pptx"
            else:
                filename = f"{company} {run.filename_suffix}.pptx"
            output_path = output_dir / filename
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
            written.append(output_path)
            log.info(f"Done: {company}")
        except Exception as e:
            log.error(f"Skipping {company}: {e}")
            continue

    if not written:
        log.warning("No presentations were produced.")

    return written


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

    written = generate_presentations(run, output_dir)
    log.info(f"Done — {len(written)} file(s) written to {output_dir}")

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
