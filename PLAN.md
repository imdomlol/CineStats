# CineStats — Implementation Plan

## What This App Does

CineStats reads Cinemark-exported `.xlsx` report files (Occupancy and Transaction Detail),
lets a non-technical user select which data they want and how they want it shaped, and
exports a clean, formatted `.xlsx` file.

---

## Constraints Driving Every Decision

| Constraint | Impact |
|---|---|
| No admin rights on work computer | Cannot install system-level software; everything must live in a user-accessible virtual environment |
| No downloads allowed | All dependencies must be installable via `pip` into a local `.venv` — no MSI installers, no system packages |
| Windows 11 + Linux Mint | No OS-specific libraries; GUI must use cross-platform toolkit |
| No LibreOffice | xlsx must be handled by a pure-Python library (`openpyxl`) |
| Non-technical daily user | GUI must be self-explanatory: file pickers, clear labels, one button to export |
| Maintainability is top priority | Extensive inline comments, clear module boundaries, no clever code |

---

## Technology Choices

### Python
Python is typically pre-installed on managed Windows/Linux environments, or can be installed
to a user-local directory (e.g. `%APPDATA%\Python`) without admin rights. Version target: **3.9+**.

### openpyxl
Pure-Python library for reading and writing `.xlsx` files. No LibreOffice, no COM interop,
no system dependencies. Installed once into the local `.venv` via pip.

### tkinter
Python's built-in GUI toolkit — ships with the standard library on both Windows and Linux.
Zero additional installs. Sufficient for a simple form-based desktop tool.

---

## Distribution Strategy

Python is not assumed to be present on the work computer. The app is distributed as a
**pre-built standalone binary** — the developer builds it once, then hands the file to the user.

### How it works
**PyInstaller** packages the Python interpreter, all dependencies (openpyxl, tkinter), and the
app source into a single file. The user receives one `.exe` (Windows) or one binary (Linux).
They double-click it. Nothing to install, no admin required, no internet access needed.

### Developer build steps (run on your Linux Mint machine)

```sh
# Install PyInstaller into your dev venv (one-time)
pip install pyinstaller

# Build the Windows exe using Wine + a Windows Python, OR just build the Linux binary
# and build Windows separately on any Windows machine with Python available.

# Linux binary:
pyinstaller --onefile --windowed --name CineStats src/main.py

# Windows exe (run this on any Windows machine that has Python, even temporarily):
pyinstaller --onefile --windowed --name CineStats src\main.py
```

The output appears in `dist/CineStats` (Linux) or `dist/CineStats.exe` (Windows).
The `dist/` folder is gitignored. Ship the binary to the user however is convenient
(USB drive, network share, email attachment).

### `--windowed` flag
This suppresses the terminal/console window on Windows so the user only sees the GUI,
not a black command prompt box behind it.

### If the work computer DOES have Python (fallback)
If Python turns out to be available after all, a venv-based setup also works without admin:

```bat
:: Windows
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python src\main.py
```
```sh
# Linux
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python3 src/main.py
```

`pip install` into a venv is always user-scoped — no admin required.

---

## Project Structure

```
CineStats/
├── .data/                   # Local xlsx files — gitignored, never committed
├── src/
│   ├── main.py              # Entry point: boots tkinter and opens the main window
│   ├── gui/
│   │   ├── app.py           # Main application window (root Tk frame, menu, layout)
│   │   └── widgets.py       # Reusable widget helpers (labeled row, file picker, etc.)
│   ├── core/
│   │   ├── reader.py        # Parses raw xlsx files into normalised Python structures
│   │   ├── transformer.py   # Applies user-selected filters and reshaping to the data
│   │   └── writer.py        # Writes the processed data back out as a formatted xlsx
│   └── config.py            # Column name mappings, known report types, app constants
├── requirements.txt         # Pinned: openpyxl==3.x.x (used for dev venv and PyInstaller)
├── PLAN.md                  # This file
└── README.md
```

---

## Data Formats (from sample files)

### Occupancy Report
- **Report title row:** `"Occupancy"` (row 1)
- **Meta rows:** Start/End Date (row 3), Theater number (row 4)
- **Column headers (row 6):** Movie, House, Showtime, Seats Sold, Total Seats, Occupancy %, Box Gross
- **Data rows:** Grouped by date; subtotal rows mixed in (identified by label patterns like `"Movie Totals"`)
- **Key transform need:** Strip subtotal/deleted rows, flatten date grouping, filter by date range or movie

### Transaction Detail Report
- **Report title row:** `"Transaction Detail"` (row 1)
- **Meta rows:** Date (row 2), Theater number (row 3)
- **Column headers (row 5):** Time, Total, Transaction ID, Terminal, Employee
- **Sub-header rows (row 9):** Item, Product Name, Quantity, Unit Price, Sales Pre-tax, Tax, Sales Post-tax, Adjustment Type/Code
- **Data rows:** Each transaction has a top-level row then one or more item sub-rows
- **Key transform need:** Flatten transaction+item rows into one row per item, filter by employee/terminal/date

---

## Module Responsibilities

### `config.py`
Defines constants so that if Cinemark ever renames a column or changes a report layout,
there is exactly one place to update it. No magic strings anywhere else in the codebase.

```python
# The column header row number for the Occupancy report (1-indexed)
OCCUPANCY_HEADER_ROW = 6

# Column names as they appear in the Occupancy report
OCCUPANCY_COLUMNS = {
    "movie":       "Movie",
    "house":       "House",
    "showtime":    "Showtime",
    "seats_sold":  "Seats Sold",
    "total_seats": "Total Seats",
    "occupancy":   "Occupancy",
    "box_gross":   "Box Gross",
}
# ... same pattern for TransactionDetail
```

### `core/reader.py`
- Opens an xlsx file with `openpyxl` in read-only mode (faster, lower memory)
- Detects report type from the title cell (row 1, col 1)
- Skips meta/blank rows, locates the real header row using the row number from `config.py`
- Returns a list of plain Python dicts — one dict per data row, keys = normalised column names
- Raises a clear, human-readable `ValueError` if the file doesn't look like a known report type

### `core/transformer.py`
- Receives the list of row-dicts from `reader.py` and the user's filter/option selections
- Filters: date range, movie title substring, house number, employee name, terminal
- Aggregates if requested: group by movie and sum seats sold; group by employee and sum totals
- Returns a new list of row-dicts — never mutates the input
- Each function is small and does one thing; the GUI calls them in sequence

### `core/writer.py`
- Receives the processed row-dicts and a target file path
- Writes a header row using the column names, then one row per dict
- Applies basic formatting: bold headers, auto-column width, date cells formatted as `YYYY-MM-DD`
- Uses `openpyxl.styles` only — no dependency on external styling libraries

### `gui/app.py`
- Single window with three sections:
  1. **Input** — file picker button + path label, report type auto-detected on load
  2. **Options** — dynamically rendered form fields based on detected report type
  3. **Export** — output path picker + "Generate Report" button + status label
- All user-facing text lives in this file so it is easy to update wording
- On error, shows a `messagebox.showerror` dialog — never crashes silently

### `gui/widgets.py`
- `LabeledRow(frame, label, widget)` — consistent left-label / right-widget layout
- `FilePicker(frame, label, mode)` — button + entry + `filedialog` wired together
- `StatusBar(frame)` — bottom-of-window label that shows "Ready", "Processing…", or error text
- Keeping reusable widgets here prevents `app.py` from becoming unmanageable

### `src/main.py`
Minimal entry point:
```python
import tkinter as tk
from gui.app import App

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
```

---

## GUI Wireframe (text)

```
┌─ CineStats ────────────────────────────────────────┐
│                                                      │
│  Input File:  [________________________] [Browse]    │
│  Report Type: Occupancy (auto-detected)              │
│                                                      │
│  ── Filters ─────────────────────────────────────   │
│  Start Date:  [__________]  End Date: [__________]  │
│  Movie:       [__________]  House:    [__________]  │
│                                                      │
│  ── Output ──────────────────────────────────────   │
│  Save As:     [________________________] [Browse]    │
│                                                      │
│               [    Generate Report    ]              │
│                                                      │
│  Status: Ready                                       │
└──────────────────────────────────────────────────────┘
```

Filters shown depend on the auto-detected report type. Transaction Detail shows
Employee and Terminal fields instead of Movie and House.

---

## Error Handling Philosophy

- **User-facing errors** (wrong file, bad date format, no data after filter): caught and shown
  as a `messagebox.showerror` dialog with plain English. No stack traces visible to the user.
- **Developer-facing errors** (unexpected data shape, openpyxl API errors): re-raised with
  context added, so a stack trace in the terminal tells the developer exactly where and why.
- The GUI thread never blocks — file I/O runs on a background thread via `threading.Thread`
  so the window stays responsive and the status bar can show "Processing…".

---

## Commenting Standards

Every module begins with a docstring explaining what it does and what it does NOT do.
Every function has a docstring covering: purpose, parameters, return value, and any
gotchas (e.g. "returns empty list if no rows match — does not raise").
Inline comments explain the WHY, not the what. Example:

```python
# Skip rows where Movie is None — these are date-group separator rows inserted by Cinemark,
# not actual showtime records. They have no data and would produce empty output rows.
rows = [r for r in rows if r.get("movie")]
```

---

## Implementation Order

1. `config.py` — column maps and constants (no dependencies)
2. `core/reader.py` — load and parse both report types, return dicts
3. `core/transformer.py` — filter functions, one per concern
4. `core/writer.py` — write dicts to formatted xlsx
5. Manual smoke test with real sample files in `data/`
6. `gui/widgets.py` — reusable widget primitives
7. `gui/app.py` — full window wired to core functions
8. Build with PyInstaller, verify binary runs on a clean machine
9. User acceptance test: hand the binary to a non-technical person and watch them use it
