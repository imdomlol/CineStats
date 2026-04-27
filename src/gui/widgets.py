"""
widgets.py — Reusable tkinter widget components for CineStats.

Each widget class is self-contained: it creates its own internal layout and
exposes a simple API (get/set/clear) so app.py never touches internal details.

What this file does NOT do: call reader/transformer/writer, or hold application state.
"""

import tkinter as tk
from tkinter import filedialog, ttk


# ── FilePicker ────────────────────────────────────────────────────────────────


class FilePicker(tk.Frame):
    """
    A row containing a label, a text entry showing the chosen path(s), and a
    Browse button that opens a file-chooser dialog.

    Supports two modes:
      "open"      — pick one existing file (returns a single path string)
      "open_multi"— pick multiple existing files (returns a list of paths)
      "save"      — choose where to save a new file (returns a single path string)
    """

    def __init__(self, parent, label_text, mode="open", filetypes=None, **kwargs):
        """
        Args:
            parent:     tkinter parent widget
            label_text: text shown to the left of the entry box
            mode:       "open", "open_multi", or "save"
            filetypes:  list of (description, pattern) tuples for the dialog filter
                        e.g. [("Excel files", "*.xlsx"), ("All files", "*.*")]
        """
        super().__init__(parent, **kwargs)

        self._mode      = mode
        self._paths     = []  # list of selected paths (always a list internally)
        self._filetypes = filetypes or [("Excel files", "*.xlsx"), ("All files", "*.*")]

        # Layout: label | entry (stretches) | button
        self.columnconfigure(1, weight=1)

        tk.Label(self, text=label_text, anchor="w", width=14).grid(
            row=0, column=0, sticky="w", padx=(0, 6)
        )

        self._entry_var = tk.StringVar()
        self._entry = tk.Entry(self, textvariable=self._entry_var, state="readonly")
        self._entry.grid(row=0, column=1, sticky="ew")

        tk.Button(self, text="Browse…", command=self._browse, width=9).grid(
            row=0, column=2, padx=(6, 0)
        )

    def _browse(self):
        """Opens the appropriate file dialog based on mode and updates the entry."""
        if self._mode == "open":
            path = filedialog.askopenfilename(filetypes=self._filetypes)
            if path:
                self._paths = [path]
                self._entry_var.set(path)

        elif self._mode == "open_multi":
            paths = filedialog.askopenfilenames(filetypes=self._filetypes)
            if paths:
                self._paths = list(paths)
                # Show all filenames (basename only) separated by semicolons so the
                # entry doesn't overflow with full paths.
                import os
                display = "; ".join(os.path.basename(p) for p in self._paths)
                self._entry_var.set(display)

        elif self._mode == "save":
            path = filedialog.asksaveasfilename(
                filetypes=self._filetypes,
                defaultextension=".xlsx",
            )
            if path:
                self._paths = [path]
                self._entry_var.set(path)

    def get(self):
        """
        Returns selected path(s).
          - "open" and "save" modes return a single string (or "" if nothing selected).
          - "open_multi" mode returns a list of strings (empty list if nothing selected).
        """
        if self._mode == "open_multi":
            return self._paths
        return self._paths[0] if self._paths else ""

    def set(self, value):
        """
        Programmatically sets the displayed path.
        Useful for restoring a previous selection on app startup.
        """
        if isinstance(value, list):
            self._paths = value
            import os
            self._entry_var.set("; ".join(os.path.basename(p) for p in value))
        else:
            self._paths = [value] if value else []
            self._entry_var.set(value or "")

    def clear(self):
        """Resets the widget to an empty state."""
        self._paths = []
        self._entry_var.set("")


# ── LabeledEntry ─────────────────────────────────────────────────────────────


class LabeledEntry(tk.Frame):
    """
    A single-line text entry with a fixed-width label to its left.

    Used for filter fields like "Movie contains:" or "Employee:".
    """

    def __init__(self, parent, label_text, width=25, **kwargs):
        """
        Args:
            parent:     tkinter parent widget
            label_text: text shown to the left of the entry
            width:      character width of the entry box
        """
        super().__init__(parent, **kwargs)
        self.columnconfigure(1, weight=1)

        tk.Label(self, text=label_text, anchor="w", width=18).grid(
            row=0, column=0, sticky="w", padx=(0, 6)
        )
        self._var = tk.StringVar()
        tk.Entry(self, textvariable=self._var, width=width).grid(
            row=0, column=1, sticky="ew"
        )

    def get(self):
        """Returns the current text value, stripped of surrounding whitespace."""
        return self._var.get().strip()

    def set(self, value):
        """Sets the entry text programmatically."""
        self._var.set(value or "")

    def clear(self):
        """Empties the entry."""
        self._var.set("")


# ── LabeledCheckbox ──────────────────────────────────────────────────────────


class LabeledCheckbox(tk.Frame):
    """A labelled checkbox (BooleanVar-backed)."""

    def __init__(self, parent, label_text, default=True, **kwargs):
        """
        Args:
            parent:     tkinter parent widget
            label_text: text shown next to the checkbox
            default:    initial checked state (True = checked)
        """
        super().__init__(parent, **kwargs)
        self._var = tk.BooleanVar(value=default)
        tk.Checkbutton(self, text=label_text, variable=self._var).grid(
            row=0, column=0, sticky="w"
        )

    def get(self):
        """Returns True if checked, False otherwise."""
        return self._var.get()

    def set(self, value):
        self._var.set(bool(value))


# ── SectionLabel ─────────────────────────────────────────────────────────────


class SectionLabel(tk.Label):
    """
    A bold, slightly larger label used as a visual section divider.

    Example: "── Filters ─────────────────"
    """

    def __init__(self, parent, text, **kwargs):
        super().__init__(
            parent,
            text=f"  {text}  ",
            font=("Calibri", 10, "bold"),
            anchor="w",
            **kwargs,
        )


# ── StatusBar ────────────────────────────────────────────────────────────────


class StatusBar(tk.Frame):
    """
    A slim bar along the bottom of the window showing the current app status.

    Status levels:
      "ready"   — neutral grey text: "Ready"
      "working" — blue text: custom message (shown while processing)
      "success" — green text: custom message (shown after successful export)
      "error"   — red text: error message
    """

    _COLOURS = {
        "ready":   "#555555",
        "working": "#1565C0",
        "success": "#2E7D32",
        "error":   "#C62828",
    }

    def __init__(self, parent, **kwargs):
        super().__init__(parent, relief="sunken", bd=1, **kwargs)
        self._var = tk.StringVar(value="Ready")
        self._label = tk.Label(
            self,
            textvariable=self._var,
            anchor="w",
            font=("Calibri", 9),
            fg=self._COLOURS["ready"],
            padx=6,
            pady=2,
        )
        self._label.pack(fill="x")

    def set_ready(self):
        """Resets to the idle 'Ready' state."""
        self._var.set("Ready")
        self._label.config(fg=self._COLOURS["ready"])

    def set_working(self, message="Processing…"):
        """Shows a blue 'working' message while a background task runs."""
        self._var.set(message)
        self._label.config(fg=self._COLOURS["working"])

    def set_success(self, message="Export complete."):
        """Shows a green success message."""
        self._var.set(message)
        self._label.config(fg=self._COLOURS["success"])

    def set_error(self, message):
        """Shows a red error message."""
        self._var.set(f"Error: {message}")
        self._label.config(fg=self._COLOURS["error"])


# ── ReportTypeLabel ───────────────────────────────────────────────────────────


class ReportTypeLabel(tk.Frame):
    """
    A read-only display showing the auto-detected report type.

    Shown as: "Report type:  Occupancy  (auto-detected)"
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.columnconfigure(1, weight=1)

        tk.Label(self, text="Report type:", anchor="w", width=14).grid(
            row=0, column=0, sticky="w", padx=(0, 6)
        )
        self._var = tk.StringVar(value="— (no file loaded)")
        tk.Label(self, textvariable=self._var, anchor="w", fg="#1565C0",
                 font=("Calibri", 10, "bold")).grid(row=0, column=1, sticky="w")

    def set(self, report_type_string):
        """Updates the displayed report type label."""
        if report_type_string:
            self._var.set(f"{report_type_string}  (auto-detected)")
        else:
            self._var.set("— (no file loaded)")

    def clear(self):
        self._var.set("— (no file loaded)")
