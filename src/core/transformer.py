"""
transformer.py — Filters and aggregates parsed row data for CineStats output.

All functions are pure: they take a list of row dicts and return a new list.
None of them mutate their input, open files, or touch the GUI.

Row dict structures are the same as those produced by reader.py:
  - Occupancy rows have keys: date, movie, house, showtime, seats_sold,
                              total_seats, occupancy_pct, box_gross
  - Transaction rows have keys: date, time, txn_id, terminal, employee,
                                item_type, category, quantity, unit_price,
                                pretax, tax, posttax, txn_total
"""

import datetime


# ── Occupancy filters ─────────────────────────────────────────────────────────


def filter_occupancy(rows, *, start_date=None, end_date=None, movie=None, house=None):
    """
    Filters a list of occupancy row dicts based on user-supplied criteria.

    All parameters are optional. Passing None (the default) means "no filter on
    this field". Multiple filters combine with AND logic (all must match).

    Args:
        rows:       list of occupancy row dicts from reader.py
        start_date: datetime.date — include only rows on or after this date
        end_date:   datetime.date — include only rows on or before this date
        movie:      str — include only rows whose movie name contains this string
                    (case-insensitive, partial match)
        house:      int — include only rows for this specific auditorium/house number

    Returns:
        A new list containing only the rows that pass all filters.
    """
    result = rows

    if start_date is not None:
        # Convert to date object if a datetime was passed (e.g. from a date-picker widget).
        start = _to_date(start_date)
        result = [r for r in result if r["date"] is not None and r["date"] >= start]

    if end_date is not None:
        end = _to_date(end_date)
        result = [r for r in result if r["date"] is not None and r["date"] <= end]

    if movie:
        # Case-insensitive substring match so users don't need exact titles.
        needle = movie.strip().lower()
        result = [r for r in result if needle in (r["movie"] or "").lower()]

    if house is not None:
        result = [r for r in result if r["house"] == int(house)]

    return result


def summarize_occupancy_by_movie(rows):
    """
    Aggregates occupancy rows by movie, summing seats sold, total seats, and box gross.

    Useful for seeing the overall performance of each film across all showtimes.
    Occupancy % is recalculated from the aggregated seat counts so it is accurate.

    Args:
        rows: list of (already-filtered) occupancy row dicts

    Returns:
        list of summary dicts sorted alphabetically by movie name, each with keys:
            movie, seats_sold, total_seats, occupancy_pct, box_gross
    """
    # Build an intermediate dict keyed by movie name.
    buckets = {}
    for row in rows:
        key = row["movie"] or "Unknown"
        if key not in buckets:
            buckets[key] = {"movie": key, "seats_sold": 0, "total_seats": 0, "box_gross": 0.0}
        buckets[key]["seats_sold"]  += row["seats_sold"]
        buckets[key]["total_seats"] += row["total_seats"]
        buckets[key]["box_gross"]   += row["box_gross"]

    # Recalculate occupancy % from totals to avoid averaging percentages.
    result = []
    for bucket in buckets.values():
        if bucket["total_seats"] > 0:
            pct = (bucket["seats_sold"] / bucket["total_seats"]) * 100
            bucket["occupancy_pct"] = f"{pct:.2f}%"
        else:
            bucket["occupancy_pct"] = "0.00%"
        result.append(bucket)

    return sorted(result, key=lambda r: (r["movie"] or "").lower())


def summarize_occupancy_by_date(rows):
    """
    Aggregates occupancy rows by date, summing across all movies and houses.

    Args:
        rows: list of occupancy row dicts

    Returns:
        list of summary dicts sorted by date ascending, each with keys:
            date, seats_sold, total_seats, occupancy_pct, box_gross
    """
    buckets = {}
    for row in rows:
        key = row["date"]
        if key not in buckets:
            buckets[key] = {"date": key, "seats_sold": 0, "total_seats": 0, "box_gross": 0.0}
        buckets[key]["seats_sold"]  += row["seats_sold"]
        buckets[key]["total_seats"] += row["total_seats"]
        buckets[key]["box_gross"]   += row["box_gross"]

    result = []
    for bucket in buckets.values():
        if bucket["total_seats"] > 0:
            pct = (bucket["seats_sold"] / bucket["total_seats"]) * 100
            bucket["occupancy_pct"] = f"{pct:.2f}%"
        else:
            bucket["occupancy_pct"] = "0.00%"
        result.append(bucket)

    return sorted(result, key=lambda r: (r["date"] or datetime.date.min))


# ── Transaction filters ───────────────────────────────────────────────────────


def filter_transactions(rows, *, employee=None, terminal=None, category=None):
    """
    Filters a list of transaction item dicts based on user-supplied criteria.

    All parameters are optional. Passing None (the default) means "no filter on
    this field". Multiple filters combine with AND logic.

    Args:
        rows:     list of transaction row dicts from reader.py
        employee: str — include only rows whose employee name contains this string
                  (case-insensitive, partial match)
        terminal: str — include only rows whose terminal contains this string
                  (case-insensitive, partial match)
        category: str — include only rows whose product category contains this string
                  (case-insensitive, partial match)

    Returns:
        A new list containing only the rows that pass all filters.
    """
    result = rows

    if employee:
        needle = employee.strip().lower()
        result = [r for r in result if needle in (r["employee"] or "").lower()]

    if terminal:
        needle = terminal.strip().lower()
        result = [r for r in result if needle in (r["terminal"] or "").lower()]

    if category:
        needle = category.strip().lower()
        result = [r for r in result if needle in (r["category"] or "").lower()]

    return result


def summarize_transactions_by_employee(rows):
    """
    Aggregates transaction rows by employee, counting transactions and summing totals.

    Each item row carries the transaction total from its parent transaction. To avoid
    counting the same transaction total multiple times (once per item), this function
    de-duplicates by transaction ID before summing totals.

    Args:
        rows: list of (already-filtered) transaction item dicts

    Returns:
        list of summary dicts sorted by employee name, each with keys:
            employee, transaction_count, item_count, total_sales
    """
    buckets = {}
    seen_txn_ids = {}  # maps employee → set of txn_ids already counted toward total

    for row in rows:
        emp = row["employee"] or "Unknown"
        if emp not in buckets:
            buckets[emp] = {
                "employee":         emp,
                "transaction_count": 0,
                "item_count":        0,
                "total_sales":       0.0,
            }
            seen_txn_ids[emp] = set()

        buckets[emp]["item_count"] += 1

        # Only count the transaction total once per unique transaction ID per employee.
        txn_id = row["txn_id"]
        if txn_id not in seen_txn_ids[emp]:
            seen_txn_ids[emp].add(txn_id)
            buckets[emp]["transaction_count"] += 1
            buckets[emp]["total_sales"]       += row["txn_total"]

    return sorted(buckets.values(), key=lambda r: (r["employee"] or "").lower())


def summarize_transactions_by_category(rows):
    """
    Aggregates transaction rows by product category, summing quantity and sales.

    Args:
        rows: list of transaction item dicts

    Returns:
        list of summary dicts sorted by category name, each with keys:
            category, item_count, total_quantity, total_sales
    """
    buckets = {}
    for row in rows:
        key = row["category"] or "Unknown"
        if key not in buckets:
            buckets[key] = {"category": key, "item_count": 0, "total_quantity": 0.0, "total_sales": 0.0}
        buckets[key]["item_count"]     += 1
        buckets[key]["total_quantity"] += row["quantity"]
        buckets[key]["total_sales"]    += row["posttax"]

    return sorted(buckets.values(), key=lambda r: (r["category"] or "").lower())


# ── Shared utilities ──────────────────────────────────────────────────────────


def compute_grand_total_occupancy(rows):
    """
    Returns a single summary dict totalling all occupancy rows.

    Used to append a grand-total row to the output xlsx.

    Args:
        rows: list of occupancy row dicts (post-filter)

    Returns:
        dict with keys: seats_sold, total_seats, occupancy_pct, box_gross
    """
    seats_sold  = sum(r["seats_sold"]  for r in rows)
    total_seats = sum(r["total_seats"] for r in rows)
    box_gross   = sum(r["box_gross"]   for r in rows)

    if total_seats > 0:
        pct = (seats_sold / total_seats) * 100
        pct_str = f"{pct:.2f}%"
    else:
        pct_str = "0.00%"

    return {
        "seats_sold":    seats_sold,
        "total_seats":   total_seats,
        "occupancy_pct": pct_str,
        "box_gross":     box_gross,
    }


def compute_grand_total_transactions(rows):
    """
    Returns a single summary dict totalling all transaction rows.

    De-duplicates by transaction ID so the total is per-transaction, not per-item.

    Args:
        rows: list of transaction item dicts (post-filter)

    Returns:
        dict with keys: transaction_count, item_count, total_sales
    """
    seen_txn_ids = set()
    total_sales  = 0.0
    item_count   = 0

    for row in rows:
        item_count += 1
        if row["txn_id"] not in seen_txn_ids:
            seen_txn_ids.add(row["txn_id"])
            total_sales += row["txn_total"]

    return {
        "transaction_count": len(seen_txn_ids),
        "item_count":        item_count,
        "total_sales":       total_sales,
    }


# ── Private helpers ───────────────────────────────────────────────────────────


def _to_date(value):
    """Normalises a value to a datetime.date, handling datetime objects as input."""
    if isinstance(value, datetime.datetime):
        return value.date()
    return value
