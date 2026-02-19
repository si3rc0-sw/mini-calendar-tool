"""Pure calendar calculations — no UI dependencies."""

import calendar
from datetime import date

DAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def month_grid(year: int, month: int) -> list[list[int | None]]:
    """Return a 6×7 grid for the given month.

    Each cell is a day number (1–31) or None for empty slots.
    Weeks start on Monday (ISO convention).
    Always 6 rows so the calendar height stays constant.
    """
    cal = calendar.Calendar(firstweekday=0)  # Monday
    days = cal.itermonthdays(year, month)

    grid: list[list[int | None]] = []
    row: list[int | None] = []
    for d in days:
        row.append(d if d != 0 else None)
        if len(row) == 7:
            grid.append(row)
            row = []
    # Pad to exactly 6 rows
    while len(grid) < 6:
        grid.append([None] * 7)
    return grid


def iso_week_numbers(year: int, month: int) -> list[str]:
    """Return ISO week numbers for each of the 6 grid rows.

    If a row is entirely empty, returns an empty string for that row.
    """
    grid = month_grid(year, month)
    weeks: list[str] = []
    for row in grid:
        # Find the first actual day in the row to derive the week number
        day = next((d for d in row if d is not None), None)
        if day is None:
            weeks.append("")
        else:
            weeks.append(str(date(year, month, day).isocalendar()[1]))
    return weeks


def day_of_year(d: date) -> int:
    """Return the 1-based day-of-year for the given date."""
    return d.timetuple().tm_yday


def prev_month(year: int, month: int) -> tuple[int, int]:
    """Return (year, month) for one month earlier."""
    if month == 1:
        return year - 1, 12
    return year, month - 1


def next_month(year: int, month: int) -> tuple[int, int]:
    """Return (year, month) for one month later."""
    if month == 12:
        return year + 1, 1
    return year, month + 1
