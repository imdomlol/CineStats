"""
config.py — All constants, column indices, and report-type definitions for CineStats.

This is the single place to update if Cinemark ever changes their export format.
No magic numbers or hardcoded strings should appear anywhere else in the codebase.

What this file does NOT do: read files, transform data, or interact with the GUI.
"""

# ── Report type identifiers ──────────────────────────────────────────────────
# These strings appear in cell A1 of each report type.
REPORT_TYPE_OCCUPANCY = "Occupancy"
REPORT_TYPE_TRANSACTION = "Transaction Detail"

# ── Occupancy report layout ──────────────────────────────────────────────────
# Row numbers are 1-indexed, matching how Excel and openpyxl describe them.

OCC_TITLE_ROW = 1          # "Occupancy"
OCC_META_DATE_ROW = 3      # Start Date / End Date
OCC_META_THEATER_ROW = 4   # Theater number
OCC_HEADER_ROW = 6         # Column labels (Movie, House, Showtime, …)
OCC_DATA_START_ROW = 7     # First row that can contain real data or a date separator

# 0-based column indices within each data row tuple.
OCC_COL_MOVIE = 1          # Movie name — only populated on the "Movie Totals" summary row
OCC_COL_HOUSE = 5          # Auditorium/house number
OCC_COL_SHOWTIME = 7       # Time string (e.g. "7:35 PM") or "Movie Totals"
OCC_COL_SEATS_SOLD = 8     # Integer seats sold for this showing
OCC_COL_TOTAL_SEATS = 10   # Integer total capacity of this auditorium
OCC_COL_OCCUPANCY_PCT = 12 # String occupancy percentage (e.g. "33.33%")
OCC_COL_BOX_GROSS = 13     # Currency string for gross box office (e.g. "$129.00")

# The sentinel value in the Showtime column that marks a movie-totals row.
OCC_MOVIE_TOTALS_SENTINEL = "Movie Totals"

# Substring that Cinemark appends to deleted/cancelled showtimes.
OCC_DELETED_MARKER = "(Deleted)"

# Human-readable column names written to the output xlsx header row.
OCC_OUTPUT_COLUMNS = [
    "Date",
    "Movie",
    "House",
    "Showtime",
    "Seats Sold",
    "Total Seats",
    "Occupancy %",
    "Box Gross ($)",
]

# ── Transaction Detail report layout ─────────────────────────────────────────

TXN_TITLE_ROW = 1          # "Transaction Detail"
TXN_META_DATE_ROW = 2      # Date of the report
TXN_META_THEATER_ROW = 3   # Theater number
TXN_DATA_START_ROW = 5     # First row that can contain transaction headers

# 0-based column indices for transaction header rows
# (the row that contains Time, Total, Transaction ID, Terminal, Employee).
TXN_COL_TIME = 3           # Time of transaction (e.g. "10:22:56 AM")
TXN_COL_TOTAL = 4          # Total dollar amount of the transaction
TXN_COL_SALE_TYPE = 8      # " - SALE" or similar — confirms it is a transaction row
TXN_COL_TXN_ID = 12        # Integer transaction ID
TXN_COL_TERMINAL = 15      # Terminal name (e.g. "0420BOX03")
TXN_COL_EMPLOYEE = 17      # Employee name (e.g. "Raygoza , Amanda")

# 0-based column indices for item detail rows
# (the rows inside each transaction listing individual products).
TXN_COL_ITEM_TYPE = 3      # Item label (e.g. "Matinee", "Md Drink")
TXN_COL_CATEGORY = 9       # Product category (e.g. "Popcorn", "Drink")
TXN_COL_QUANTITY = 13      # Numeric quantity (float) — also used to identify item rows
TXN_COL_UNIT_PRICE = 14    # Unit price string (e.g. "$10.00")
TXN_COL_PRETAX = 16        # Sales pre-tax string
TXN_COL_TAX = 19           # Tax amount string (or "-" for zero tax)
TXN_COL_POSTTAX = 24       # Sales post-tax string

# Sentinel strings used to detect label/separator rows that must be skipped.
TXN_HEADER_SENTINEL_TIME = "Time"      # Appears in col D of the transaction header LABEL row
TXN_HEADER_SENTINEL_TOTAL = "Total"    # Appears in col E of the transaction header LABEL row
TXN_ITEM_SENTINEL = "Transaction Detail:"
TXN_PAYMENT_SENTINEL = "Payment Detail:"
TXN_ITEM_HEADER_SENTINEL = "Item"      # Col D on the item column-header row

# Human-readable column names written to the output xlsx header row.
TXN_OUTPUT_COLUMNS = [
    "Date",
    "Time",
    "Transaction ID",
    "Terminal",
    "Employee",
    "Item Type",
    "Product Category",
    "Quantity",
    "Unit Price ($)",
    "Sales Pre-Tax ($)",
    "Tax ($)",
    "Sales Post-Tax ($)",
    "Transaction Total ($)",
]
