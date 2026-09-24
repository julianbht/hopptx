@echo off
cd /d "%~dp0"

rem First run downloads Python and libraries; keep this window visible for that and for errors.
echo Starting hopptx, please wait...
uv sync --quiet
if errorlevel 1 (
    echo.
    echo hopptx could not be set up. Is uv installed and is there an internet connection?
    pause
    exit /b 1
)

rem uvw is uv without a console window; this window closes and the app keeps running.
where uvw >nul 2>nul
if errorlevel 1 (
    start "" uv run --no-sync hopptx-easy
) else (
    start "" uvw run --no-sync hopptx-easy
)
