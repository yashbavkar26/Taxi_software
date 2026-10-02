"""Data normalization and sanitization utilities for Excel cell values."""

import datetime
import re
from dataclasses import dataclass
from decimal import Decimal
from typing import Any, Optional
import dateutil.parser


@dataclass
class ColumnDef:
    """Definition of a table column and its normalization rules."""

    db_column: str
    data_type: str  # 'string', 'integer', 'date', 'time', 'phone', 'boolean'
    is_required: bool = False
    is_unique: bool = False
    max_length: Optional[int] = None
    default_value: Any = None



def normalize_string(val: Any) -> Optional[str]:
    """Trim whitespace and convert empty strings to None."""
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def normalize_integer(val: Any) -> Optional[int]:
    """Normalize numeric cell to integer, preserving 0 and avoiding None for 0."""
    if val is None:
        return None
    if isinstance(val, bool):
        return 1 if val else 0
    if isinstance(val, int):
        return val
    if isinstance(val, (float, Decimal)):
        return int(val)
    s = str(val).strip()
    if not s:
        return None
    try:
        f = float(s)
        return int(f)
    except (ValueError, OverflowError) as e:
        raise ValueError(f"'{val}' is not a valid integer") from e


def normalize_phone(val: Any) -> Optional[str]:
    """Normalize phone/mobile numbers cleanly without trailing .0 from Excel floats."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    elif isinstance(val, Decimal) and val % 1 == 0:
        val = int(val)
    s = str(val).strip()
    if not s or not any(c.isdigit() for c in s) or "cancel" in s.lower():
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s



def normalize_date(val: Any) -> Optional[datetime.date]:
    """Normalize dates from datetime, date, or string representations."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if not s:
        return None

    # Replace asterisks or typos in date separators like '27*01-2026'
    s = s.replace("*", "-")

    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
        "%b %d, %Y",
        "%d %b %Y",
        "%d-%m%Y",
        "%d/%m%Y",
    ]
    for fmt in formats:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    try:
        dt = dateutil.parser.parse(s, dayfirst=True)
        return dt.date()
    except (ValueError, OverflowError) as e:
        raise ValueError(f"'{val}' is not a valid date") from e


def normalize_time(val: Any) -> Optional[datetime.time]:
    """Normalize time values from time, datetime, timedelta, or string representations."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.time()
    if isinstance(val, datetime.time):
        return val
    if isinstance(val, datetime.timedelta):
        total_seconds = int(val.total_seconds()) % 86400
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return datetime.time(hours, minutes, seconds)
    s = str(val).strip()
    if not s or s.lower() in ("-", "--", "---", "0", "00", "00:00:00", "na", "n/a", "nil", "none"):
        return None

    # Handle cancellation remarks or non-time status notes in time cells
    if "cancel" in s.lower() or s.lower() in ("cash", "paid to driver"):
        return None

    # Replace typo 'o'/'O' in minutes of time strings like '02:oo PM'
    s = re.sub(r":([oO0-9]{2})", lambda m: ":" + m.group(1).replace("o", "0").replace("O", "0"), s)

    # Normalize spaces around colons like '5 :30 am' or '10: 00 AM'
    s = re.sub(r"\s*:\s*", ":", s)

    # Handle formats like '22 hrs' or '22hrs'
    hrs_match = re.match(r"^(\d{1,2})\s*hrs?$", s, re.IGNORECASE)
    if hrs_match:
        h = int(hrs_match.group(1))
        if 0 <= h <= 23:
            return datetime.time(h, 0)

    # Handle redundant AM/PM appended to 24-hour time like '14:35 PM'
    s = re.sub(r"^(1[3-9]|2[0-3]):([0-5][0-9])\s*[ap]m$", r"\1:\2", s, flags=re.IGNORECASE)

    formats = [
        "%H:%M:%S",
        "%H:%M",
        "%I:%M %p",
        "%I:%M:%S %p",
        "%I:%M%p",
        "%I:%M:%S%p",
        "%I %p",
        "%I%p",
    ]
    for fmt in formats:
        try:
            return datetime.datetime.strptime(s, fmt).time()
        except ValueError:
            pass

    try:
        dt = dateutil.parser.parse(s)
        return dt.time()
    except (ValueError, OverflowError) as e:
        raise ValueError(f"'{val}' is not a valid time format") from e


def normalize_boolean(val: Any) -> bool:
    """Normalize boolean value into Python bool (True/False)."""
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        if val in (1, 1.0):
            return True
        if val in (0, 0.0):
            return False
        raise ValueError(f"'{val}' is not a valid boolean")
    s = str(val).strip().lower()
    if not s:
        return False
    if s in ("true", "1", "yes", "y", "received"):
        return True
    if s in ("false", "0", "no", "n"):
        return False
    raise ValueError(f"'{val}' is not a valid boolean")


def normalize_cell(val: Any, col_def: ColumnDef) -> Any:
    """Route value normalization according to target column definition."""
    res = None
    if col_def.data_type == "string":
        res = normalize_string(val)
        if res is not None and col_def.max_length and len(res) > col_def.max_length:
            res = res[: col_def.max_length]
    elif col_def.data_type == "integer":
        res = normalize_integer(val)
    elif col_def.data_type == "phone":
        res = normalize_phone(val)
        if res is not None and col_def.max_length and len(res) > col_def.max_length:
            res = res[: col_def.max_length]
    elif col_def.data_type == "date":
        res = normalize_date(val)
    elif col_def.data_type == "time":
        res = normalize_time(val)
    elif col_def.data_type == "boolean":
        res = normalize_boolean(val)
    else:
        res = val

    if res is None and col_def.default_value is not None:
        return col_def.default_value
    return res

