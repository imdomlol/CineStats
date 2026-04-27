# CineStats.spec — PyInstaller build specification.
#
# Produces a single-file executable with no external dependencies.
#
# Usage (run from the project root):
#   Linux:   pyinstaller CineStats.spec
#   Windows: pyinstaller CineStats.spec
#
# Output: dist/CineStats  (Linux)  or  dist/CineStats.exe  (Windows)
#
# First-time setup on a developer machine:
#   pip install pyinstaller
#   pyinstaller CineStats.spec

import sys
import os

block_cipher = None

a = Analysis(
    # Entry point — relative to the project root where this .spec lives.
    ['src/main.py'],

    # Tell PyInstaller to search src/ for imports, matching how we run the app.
    pathex=[os.path.abspath('src')],

    # No native DLL dependencies beyond tkinter (bundled with Python).
    binaries=[],

    # No data files needed — .xlsx inputs are provided by the user at runtime.
    datas=[],

    # All imports are either stdlib or openpyxl (pure Python).
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
    ],

    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],

    name='CineStats',

    # --windowed: suppress the console/terminal window on Windows so the user
    # only sees the GUI, not a black command-prompt box behind it.
    console=False,

    # --onefile: bundle everything into a single executable for easy distribution.
    # The user just copies one file to their machine.
    onefile=True,

    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,

    # Windows-specific: set the app icon if one is provided.
    # icon='assets/icon.ico',  # Uncomment and add an .ico file to enable
)
