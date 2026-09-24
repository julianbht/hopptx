#!/bin/bash
cd "$(dirname "$0")"

# First run downloads Python and libraries; keep this window visible for that and for errors.
echo "Starting hopptx, please wait..."
if ! uv sync --quiet; then
    echo
    echo "hopptx could not be set up. Is uv installed and is there an internet connection?"
    read -r -p "Press Enter to close."
    exit 1
fi

# Detach the app so closing this Terminal window does not close it.
nohup uv run --no-sync hopptx-easy >/dev/null 2>&1 &
disown
echo "hopptx is open. You can close this window."
