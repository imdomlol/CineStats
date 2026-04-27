#!/bin/bash
# run.sh — Launch CineStats on Linux.
#
# Usage: ./run.sh
# Run setup.sh first if you haven't already.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV="$SCRIPT_DIR/.venv/bin/python3"

if [ ! -f "$VENV" ]; then
    echo "Virtual environment not found."
    echo "Please run:  bash setup.sh"
    exit 1
fi

"$VENV" "$SCRIPT_DIR/src/main.py"
