# CineStats

Cinemark statistic file optimizer for ideal human perception.

---

## Requirements

- **Python 3.8+** must be installed and available on your PATH.
  - Linux: install via your package manager (`sudo apt install python3 python3-venv`)
  - Windows: download from [python.org](https://www.python.org/downloads/) — check **"Add Python to PATH"** during installation
- No admin/sudo rights required after Python is installed.

---

## Setup & Launch

### Linux / macOS

**First-time setup** (run once):

```bash
bash setup.sh
```

This creates a `.venv/` folder in the project root and installs `openpyxl`.

**Launch the app:**

```bash
bash run.sh
```

Or directly:

```bash
.venv/bin/python3 src/main.py
```

---

### Windows

**First-time setup** (run once):

Double-click `setup.bat`, or run from PowerShell / Command Prompt:

```bat
setup.bat
```

This creates a `.venv\` folder in the project root and installs `openpyxl`.

**Launch the app:**

Double-click `run.bat`, or run from PowerShell / Command Prompt:

```bat
run.bat
```

Or directly from PowerShell:

```powershell
.venv\Scripts\python.exe src\main.py
```

---

## Building a Standalone Executable

To distribute CineStats as a single file with no Python requirement on the target machine:

```bash
pip install pyinstaller
pyinstaller CineStats.spec
```

Output:
- Linux/macOS: `dist/CineStats`
- Windows: `dist/CineStats.exe`

The binary bundles Python, tkinter, and openpyxl — the user just copies one file.

---

## Smoke-testing Without the GUI

```bash
# Occupancy report path
.venv/bin/python3 -c "
import sys; sys.path.insert(0, 'src')
from core.reader import read_file
from core.transformer import filter_occupancy, compute_grand_total_occupancy
from core.writer import write_occupancy

rows = read_file('data/Occupancy.xlsx')['rows']
gt   = compute_grand_total_occupancy(rows)
write_occupancy(rows, '/tmp/out.xlsx', grand_total=gt)
print(len(rows), 'rows written')
"
```

Swap in `read_file('data/TransactionDetail.xlsx')`, `filter_transactions`, `compute_grand_total_transactions`, and `write_transaction_detail` for the Transaction Detail path.

On Windows, replace `.venv/bin/python3` with `.venv\Scripts\python.exe` and `/tmp/out.xlsx` with a Windows path such as `C:\Temp\out.xlsx`.
