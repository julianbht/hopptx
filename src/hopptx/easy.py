"""Interactive wizard for non-technical users.

Opens a file picker, asks for date range, generates presentations,
and opens the output folder when done.
"""

import os
import platform
import subprocess
import sys
import logging
from datetime import datetime
from pathlib import Path

from rich.console import Console

from hopptx.main import _setup_logging, generate_presentations
from hopptx.paths import OUTPUT_DIR
from hopptx.schemas.config import InputConfig, load_runs

console = Console()


def _open_folder(path: Path) -> None:
    """Open a folder in the system file explorer."""
    folder = str(path.resolve())
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(folder)
        elif system == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception:
        pass  # If it fails, user still has the printed path


def _pick_files_or_folder() -> tuple[str, str]:
    """Open a native file dialog. Returns (input_type, input_path)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        console.print(
            "\n[red]Could not open file picker (tkinter not available).[/red]\n"
            "You can enter the path manually instead."
        )
        return _manual_path_entry()

    root = tk.Tk()
    root.withdraw()
    # Bring dialog to front
    root.attributes("-topmost", True)

    console.print("\nA file picker will open. Choose one of the following:")
    console.print("  [bold]1[/bold] — Select one or more Excel files")
    console.print("  [bold]2[/bold] — Select a folder containing Excel files")

    choice = ""
    while choice not in ("1", "2"):
        choice = input("\nYour choice (1 or 2): ").strip()

    if choice == "1":
        files = filedialog.askopenfilenames(
            title="Select Excel report file(s)",
            filetypes=[("Excel files", "*.xlsx")],
            parent=root,
        )
        root.destroy()
        if not files:
            return "", ""
        if len(files) == 1:
            return "file", files[0]
        # Multiple files: find common parent directory, copy to temp approach
        # Actually — we only support file or directory mode.
        # For multiple files, use the common directory and filter later.
        # Simplest: if all in same directory, use directory mode.
        parents = {str(Path(f).parent) for f in files}
        if len(parents) == 1:
            return "directory", str(Path(files[0]).parent)
        # Files from different folders — use the first file only, warn
        console.print(
            "\n[yellow]Files from multiple folders selected. "
            "Using only the first file.[/yellow]"
        )
        return "file", files[0]
    else:
        folder = filedialog.askdirectory(
            title="Select folder containing Excel reports",
            parent=root,
        )
        root.destroy()
        if not folder:
            return "", ""
        return "directory", folder


def _manual_path_entry() -> tuple[str, str]:
    """Fallback: ask user to type the path."""
    console.print("\nEnter the path to your Excel file or folder:")
    path_str = input("> ").strip().strip('"').strip("'")
    if not path_str:
        return "", ""
    p = Path(path_str)
    if p.is_dir():
        return "directory", str(p)
    elif p.is_file():
        return "file", str(p)
    else:
        console.print(f"\n[red]Path not found: {path_str}[/red]")
        return "", ""


def _ask_date(label: str) -> str:
    """Prompt for a date in YYYY-MM-DD format with validation."""
    while True:
        value = input(f"\n{label} (YYYY-MM-DD): ").strip()
        try:
            datetime.strptime(value, "%Y-%m-%d")
            return value
        except ValueError:
            console.print(
                f"[red]'{value}' is not a valid date. "
                f"Please use the format YYYY-MM-DD (e.g. 2026-01-15).[/red]"
            )


def run_easy() -> None:
    """Interactive wizard entry point."""
    console.print("\n[bold]hopptx — Presentation Generator[/bold]")
    console.print("=" * 40)

    # Step 1: Pick files
    input_type, input_path = _pick_files_or_folder()
    if not input_path:
        console.print("\nNo file or folder selected. Exiting.")
        return

    selected = Path(input_path)
    console.print(f"\nSelected: [bold]{selected}[/bold]")

    # Step 2: Ask for date range
    console.print("\nEnter the date range for the report:")
    start = _ask_date("Start date")
    end = _ask_date("End date  ")

    # Step 3: Load default config and override input + dates
    console.print("\nLoading configuration...")
    try:
        runs = load_runs()
    except Exception as e:
        console.print(f"\n[red]Could not load configuration: {e}[/red]")
        _wait_for_key()
        return

    defaults = [r for r in runs if r.name == "default"]
    if not defaults:
        console.print(
            "\n[red]No 'default' run found in config/runs.json. "
            "Please make sure a run named 'default' exists.[/red]"
        )
        _wait_for_key()
        return

    base = defaults[0]
    run = base.model_copy(update={
        "start": start,
        "end": end,
        "input": InputConfig(type=input_type, path=input_path),
    })

    # Step 4: Set up output directory and logging
    now = datetime.now()
    output_dir = (
        OUTPUT_DIR
        / "pptx"
        / "easy"
        / now.strftime("%Y")
        / now.strftime("%m")
        / now.strftime("%d")
        / now.strftime("%Y-%m-%d_%H-%M-%S")
    )
    file_handler = _setup_logging(output_dir)

    # Step 5: Run
    console.print(f"\nGenerating presentations for [bold]{start}[/bold] to [bold]{end}[/bold]...\n")
    try:
        written = generate_presentations(run, output_dir)
    except Exception as e:
        console.print(f"\n[red]Something went wrong: {e}[/red]")
        console.print(f"\nFor details, see the log: [link=file://{output_dir / 'run.log'}]{output_dir / 'run.log'}[/link]")
        logging.root.removeHandler(file_handler)
        file_handler.close()
        _wait_for_key()
        return

    logging.root.removeHandler(file_handler)
    file_handler.close()

    # Step 6: Show results
    if not written:
        console.print("\n[yellow]No presentations were generated.[/yellow]")
        console.print("This can happen if no data matched the selected date range,")
        console.print("or if the Excel files didn't have the expected columns.")
        console.print(f"\nFor details, see the log: {output_dir / 'run.log'}")
        _wait_for_key()
        return

    console.print(f"\n[bold green]Done![/bold green] {len(written)} presentation(s) created.\n")
    for path in written:
        console.print(f"  {path.name}")

    console.print(f"\n[bold]Your reports are here:[/bold]  {output_dir.resolve()}")
    console.print(f"[bold]Full log:[/bold]              {(output_dir / 'run.log').resolve()}")

    # Try to open the output folder
    _open_folder(output_dir)

    _wait_for_key()


def _wait_for_key() -> None:
    """Wait for user to press Enter before closing (useful when run from a shortcut)."""
    console.print("\nPress [bold]Enter[/bold] to close.")
    input()
