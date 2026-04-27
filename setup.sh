#!/bin/bash
# setup.sh — One-time setup for CineStats on Linux.
#
# Creates a local virtual environment and installs dependencies.
# No admin/sudo rights required.
#
# Usage: bash setup.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "Creating virtual environment..."
python3 -m venv "$SCRIPT_DIR/.venv"

echo "Installing dependencies..."
"$SCRIPT_DIR/.venv/bin/pip" install --quiet -r "$SCRIPT_DIR/requirements.txt"

echo ""
echo "Setup complete. Run the app with:  bash run.sh"
