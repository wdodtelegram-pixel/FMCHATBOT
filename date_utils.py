
from __future__ import annotations

import math
import re
from datetime import date, datetime

# Regex that recognises text a user already typed in the target format.
# PURPOSE: if the cell already says "Week 2 March 2026" we should trust it and
# just tidy the capitalisation instead of trying to re-parse it as a date.
_ALREADY_FORMATTED = re.compile(
    r"^\s*week\s*([1-5])\s+([a-z]+)\s+(\d{4})\s*$", re.IGNORECASE
)

_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _week_of_month(day: int) -> int:
    """
    PURPOSE: convert a day-of-month number into a week number 1-5.

    Days 1-7 are week 1, 8-14 week 2, and so on. This is the simple
    "calendar block" definition facilities teams normally mean when they say
    'week 1', rather than the ISO week-number definition.
    """
    return min(5, (day - 1) // 7 + 1)


def _is_missing(value) -> bool:
    return value is None or (isinstance(value, float) and math.isnan(value))


def _parse_date(text: str) -> datetime | None:
    normalized = text.strip()
    if not normalized:
        return None

    # Try ISO-like and common numeric formats first.
    iso_like = normalized.replace("/", "-").replace(".", "-")
    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%d-%m",
    ):
        try:
            parsed = datetime.strptime(iso_like, fmt)
            if fmt == "%d-%m":
                return parsed.replace(year=datetime.now().year)
            if parsed.year < 100:
                parsed = parsed.replace(year=parsed.year + 2000)
            return parsed
        except ValueError:
            continue

    # Try written month names.
    for fmt in (
        "%d %B %Y",
        "%d %b %Y",
        "%B %d %Y",
        "%b %d %Y",
        "%d %B %y",
        "%d %b %y",
    ):
        try:
            parsed = datetime.strptime(normalized, fmt)
            if parsed.year < 100:
                parsed = parsed.replace(year=parsed.year + 2000)
            return parsed
        except ValueError:
            continue

    # Fallback parse for common day-first numeric dates.
    match = re.match(r"^(\d{1,2})[\-/](\d{1,2})[\-/](\d{2,4})$", normalized)
    if match:
        day, month, year = match.groups()
        year = int(year)
        if year < 100:
            year += 2000
        try:
            return datetime(year, int(month), int(day))
        except ValueError:
            return None

    return None


def to_week_format(value) -> str:
    """
    PURPOSE: the single public function of this module. Takes anything from an
    Excel cell and returns a display string.

    Returns "-" for blanks so the message never shows the ugly word "NaT" or
    "nan" to a user, and returns unrecognised text unchanged (so a deliberate
    "TBC" or "On hold" written by staff still shows through).
    """
    if _is_missing(value):
        return "-"

    if isinstance(value, (datetime, date)):
        return f"Week {_week_of_month(value.day)} {_MONTH_NAMES[value.month - 1]} {value.year}"

    text = str(value).strip()
    if not text:
        return "-"

    match = _ALREADY_FORMATTED.match(text)
    if match:
        week, month, year = match.groups()
        return f"Week {week} {month.capitalize()} {year}"

    parsed = _parse_date(text)
    if parsed:
        return f"Week {_week_of_month(parsed.day)} {_MONTH_NAMES[parsed.month - 1]} {parsed.year}"

    return text
