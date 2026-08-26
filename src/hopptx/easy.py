"""Graphical wizard for non-technical users.

Opens a small window with a multi-file picker and date pickers, generates
presentations in a background thread, and opens the output folder when done.
No terminal interaction is required.
"""

import logging
import os
import platform
import queue
import subprocess
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkcalendar import DateEntry

from hopptx.main import generate_presentations
from hopptx.paths import OUTPUT_DIR
from hopptx.schemas.config import InputConfig, RunConfig, load_runs

log = logging.getLogger(__name__)


def _open_path(path: Path) -> None:
    """Open a file or folder with the system default application."""
    target = str(path.resolve())
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(target)
        elif system == "Darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])
    except Exception:
        pass  # If it fails, user still has the path shown in the window


class _QueueLogHandler(logging.Handler):
    """Logging handler that pushes formatted records onto a thread-safe queue.

    The generation work runs on a background thread, but only the main
    thread is allowed to touch Tk widgets — the queue is the hand-off point.
    """

    def __init__(self, log_queue: "queue.Queue[str]") -> None:
        super().__init__()
        self._queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        self._queue.put(self.format(record))


class EasyApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("hopptx — Presentation Generator")
        self.root.geometry("640x480")
        self.root.minsize(560, 420)

        self.selected_files: list[Path] = []
        self.files_summary_var = tk.StringVar(value="No files selected")
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.output_dir: Path | None = None
        self.file_handler: logging.FileHandler | None = None
        self.queue_handler: _QueueLogHandler | None = None

        self._build_ui()
        self.root.after(100, self._poll_log_queue)

    # -- UI construction --------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}

        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        ttk.Label(top, text="Excel files:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.files_summary_var, state="readonly", width=40).grid(
            row=0, column=1, sticky="ew", padx=(6, 6)
        )
        ttk.Button(top, text="Select Files...", command=self._select_files).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        dates = ttk.Frame(self.root)
        dates.pack(fill="x", **pad)
        ttk.Label(dates, text="Start date:").grid(row=0, column=0, sticky="w")
        self.start_date = DateEntry(dates, date_pattern="yyyy-mm-dd")
        self.start_date.grid(row=0, column=1, sticky="w", padx=(6, 24))
        ttk.Label(dates, text="End date:").grid(row=0, column=2, sticky="w")
        self.end_date = DateEntry(dates, date_pattern="yyyy-mm-dd")
        self.end_date.grid(row=0, column=3, sticky="w", padx=(6, 0))

        today = date.today()
        self.end_date.set_date(today)
        self.start_date.set_date(today - timedelta(days=90))

        action = ttk.Frame(self.root)
        action.pack(fill="x", **pad)
        self.generate_btn = ttk.Button(
            action, text="Generate Presentations", command=self._start_generation
        )
        self.generate_btn.pack(side="left")
        self.progress = ttk.Progressbar(action, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(12, 0))

        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log_frame, state="disabled", wrap="word", height=14)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        bottom = ttk.Frame(self.root)
        bottom.pack(fill="x", **pad)
        self.open_folder_btn = ttk.Button(
            bottom,
            text="Open Output Folder",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_btn.pack(side="left")
        self.view_log_btn = ttk.Button(
            bottom,
            text="View Log",
            command=self._view_log,
            state="disabled",
        )
        self.view_log_btn.pack(side="left", padx=(6, 0))
        ttk.Button(bottom, text="Close", command=self.root.destroy).pack(side="right")

    # -- Actions ------------------------------------------------------------

    def _select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select Excel report file(s)",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self.root,
        )
        if not files:
            return
        self.selected_files = [Path(f) for f in files]
        self.files_summary_var.set(f"{len(self.selected_files)} file(s) selected")

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log(f"Selected {len(self.selected_files)} Excel file(s):")
        for f in self.selected_files:
            self._log(f"  {f.name}")

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _start_generation(self) -> None:
        if not self.selected_files:
            messagebox.showerror(
                "No files selected",
                "Please select the Excel report file(s) to process.",
            )
            return

        start_d = self.start_date.get_date()
        end_d = self.end_date.get_date()
        if end_d < start_d:
            messagebox.showerror(
                "Invalid date range", "The end date must be on or after the start date."
            )
            return

        try:
            runs = load_runs()
        except Exception as e:
            messagebox.showerror("Configuration error", f"Could not load configuration:\n{e}")
            return

        defaults = [r for r in runs if r.name == "default"]
        if not defaults:
            messagebox.showerror(
                "Configuration error",
                "No 'default' run found in config/runs.json. "
                "Please make sure a run named 'default' exists.",
            )
            return

        start = start_d.strftime("%Y-%m-%d")
        end = end_d.strftime("%Y-%m-%d")
        base = defaults[0]
        run = base.model_copy(
            update={
                "start": start,
                "end": end,
                "input": InputConfig(
                    type="files", paths=[str(f) for f in self.selected_files]
                ),
            }
        )

        now = datetime.now()
        self.output_dir = (
            OUTPUT_DIR
            / "pptx"
            / "easy"
            / now.strftime("%Y")
            / now.strftime("%m")
            / now.strftime("%d")
            / now.strftime("%Y-%m-%d_%H-%M-%S")
        )
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._attach_logging(self.output_dir)

        self.generate_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.view_log_btn.configure(state="normal")
        self.progress.start(12)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log(f"Generating presentations for {start} to {end}...")

        threading.Thread(target=self._run_generation, args=(run,), daemon=True).start()

    def _attach_logging(self, output_dir: Path) -> None:
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)

        self.file_handler = logging.FileHandler(output_dir / "run.log")
        self.file_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s %(levelname)-8s %(name)s: %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        self.queue_handler = _QueueLogHandler(self.log_queue)
        self.queue_handler.setFormatter(logging.Formatter("%(message)s"))

        root_logger.handlers.clear()
        root_logger.addHandler(self.file_handler)
        root_logger.addHandler(self.queue_handler)

    def _detach_logging(self) -> None:
        root_logger = logging.getLogger()
        if self.file_handler is not None:
            root_logger.removeHandler(self.file_handler)
            self.file_handler.close()
            self.file_handler = None
        if self.queue_handler is not None:
            root_logger.removeHandler(self.queue_handler)
            self.queue_handler = None

    def _run_generation(self, run: RunConfig) -> None:
        try:
            written = generate_presentations(run, self.output_dir)
            error = None
        except Exception as e:
            written = []
            error = str(e)
        self.root.after(0, self._on_generation_done, written, error)

    def _on_generation_done(self, written: list[Path], error: str | None) -> None:
        self._detach_logging()
        self.progress.stop()
        self.generate_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal")

        if error is not None:
            self._log(f"\nSomething went wrong: {error}")
            self._log(f"For details, see the log: {self.output_dir / 'run.log'}")
            messagebox.showerror("Generation failed", f"Something went wrong:\n{error}")
            return

        if not written:
            self._log("\nNo presentations were generated.")
            self._log("This can happen if no data matched the selected date range,")
            self._log("or if the Excel files didn't have the expected columns.")
            self._log(f"For details, see the log: {self.output_dir / 'run.log'}")
            messagebox.showwarning(
                "No presentations created",
                "No presentations were generated. See the log for details.",
            )
            return

        self._log(f"\nDone! {len(written)} presentation(s) created.")
        for path in written:
            self._log(f"  {path.name}")
        self._log(f"\nYour reports are here: {self.output_dir.resolve()}")

        messagebox.showinfo("Done", f"{len(written)} presentation(s) created.")
        _open_path(self.output_dir)

    def _open_output_folder(self) -> None:
        if self.output_dir is not None:
            _open_path(self.output_dir)

    def _view_log(self) -> None:
        if self.output_dir is None:
            return
        log_path = self.output_dir / "run.log"
        if log_path.exists():
            _open_path(log_path)
        else:
            messagebox.showinfo("No log yet", "The log file hasn't been created yet.")

    def _poll_log_queue(self) -> None:
        try:
            while True:
                self._log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def run(self) -> None:
        self.root.mainloop()


def run_easy() -> None:
    """GUI wizard entry point."""
    EasyApp().run()
