"""
main.py - Entry point for CineStats.

Run this file directly:
  Windows: .venv\\Scripts\\python src\\main.py
  Linux:   .venv/bin/python3 src/main.py

Or build a standalone binary with PyInstaller:
  pyinstaller --onefile --windowed --name CineStats src/main.py
"""

import os
import sys

# Ensure the src/ directory is on the Python path so that all sibling
# modules (config, core.*, gui.*) can be imported without a package prefix.
# This is necessary when PyInstaller bundles the app or when the script is
# run with `python src/main.py` from the project root.
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

import tkinter as tk

from gui.app import App


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
