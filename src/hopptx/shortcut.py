"""Create a desktop shortcut to run hopptx-easy."""

import os
import platform
import stat
from pathlib import Path

from rich.console import Console

console = Console()


def _get_desktop() -> Path:
    """Return the user's Desktop path."""
    system = platform.system()
    if system == "Windows":
        # USERPROFILE is always set on Windows
        return Path(os.environ.get("USERPROFILE", "~")) / "Desktop"
    elif system == "Darwin":
        return Path.home() / "Desktop"
    else:
        return Path.home() / "Desktop"


def _get_project_dir() -> Path:
    """Return the project root directory (where pyproject.toml lives)."""
    # This module lives at src/hopptx/shortcut.py, so project root is 3 levels up
    return Path(__file__).resolve().parent.parent.parent


def create_shortcut() -> None:
    """Create a desktop shortcut that runs hopptx-easy."""
    project_dir = _get_project_dir()
    desktop = _get_desktop()
    system = platform.system()

    if not desktop.exists():
        console.print(f"[red]Desktop folder not found: {desktop}[/red]")
        console.print("Please create the shortcut manually.")
        return

    if system == "Windows":
        shortcut_path = desktop / "hopptx.bat"
        shortcut_path.write_text(
            f'@echo off\r\n'
            f'cd /d "{project_dir}"\r\n'
            f'uv run hopptx-easy\r\n',
            encoding="utf-8",
        )
    else:
        shortcut_path = desktop / "hopptx.command"
        shortcut_path.write_text(
            f'#!/bin/bash\n'
            f'cd "{project_dir}"\n'
            f'uv run hopptx-easy\n',
            encoding="utf-8",
        )
        # Make executable
        shortcut_path.chmod(shortcut_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    console.print(f"\n[bold green]Shortcut created:[/bold green] {shortcut_path}")
    console.print("\nDouble-click it to generate presentations.")
