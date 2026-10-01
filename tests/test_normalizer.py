"""Unit tests for data normalization and sanitization."""

import datetime
from decimal import Decimal
import pytest

from app.common.normalizer import (
    ColumnDef,
    normalize_boolean,
    normalize_cell,
    normalize_date,
    normalize_integer,
    normalize_phone,
    normalize_string,
    normalize_time,
)


class TestNormalizer:
    def test_normalize_string(self):
        assert normalize_string("  hello world  ") == "hello world"
        assert normalize_string("   ") is None
        assert normalize_string("") is None
        assert normalize_string(None) is None
        assert normalize_string(123) == "123"

    def test_normalize_integer_preserves_zero(self):
        assert normalize_integer(0) == 0
        assert normalize_integer("0") == 0
        assert normalize_integer(" 0 ") == 0
        assert normalize_integer(0.0) == 0

    def test_normalize_integer_valid(self):
        assert normalize_integer(42) == 42
        assert normalize_integer("42") == 42
        assert normalize_integer(42.0) == 42
        assert normalize_integer(Decimal("42")) == 42
        assert normalize_integer(None) is None
        assert normalize_integer("") is None

    def test_normalize_integer_invalid(self):
        with pytest.raises(ValueError, match="not a valid integer"):
            normalize_integer("forty-two")

    def test_normalize_phone(self):
        assert normalize_phone("8830087698") == "8830087698"
        assert normalize_phone(8830087698) == "8830087698"
        assert normalize_phone(8830087698.0) == "8830087698"
        assert normalize_phone(" 9822150660 ") == "9822150660"
        assert normalize_phone(0) == "0"
        assert normalize_phone(None) is None
        assert normalize_phone("") is None

    def test_normalize_date(self):
        d = datetime.date(2026, 2, 1)
        dt = datetime.datetime(2026, 2, 1, 15, 30)
        assert normalize_date(d) == d
        assert normalize_date(dt) == d
        assert normalize_date("2026-02-01") == d
        assert normalize_date("01-02-2026") == d
        assert normalize_date("01/02/2026") == d
        assert normalize_date("29-122025") == datetime.date(2025, 12, 29)
        assert normalize_date("29/122025") == datetime.date(2025, 12, 29)
        assert normalize_date(None) is None
        assert normalize_date("  ") is None

    def test_normalize_date_invalid(self):
        with pytest.raises(ValueError, match="not a valid date"):
            normalize_date("invalid-date-xyz")

    def test_normalize_date_with_asterisk(self):
        assert normalize_date("27*01-2026") == datetime.date(2026, 1, 27)

    def test_normalize_time(self):
        t = datetime.time(12, 0)
        dt = datetime.datetime(2026, 2, 1, 12, 0)
        assert normalize_time(t) == t
        assert normalize_time(dt) == t
        assert normalize_time("12:00") == t
        assert normalize_time("12:00:00") == t
        assert normalize_time("12:00 PM") == t
        assert normalize_time("12:00 AM") == datetime.time(0, 0)
        assert normalize_time("14:35 PM") == datetime.time(14, 35)
        assert normalize_time("9 PM") == datetime.time(21, 0)
        assert normalize_time(None) is None
        assert normalize_time("-") is None
        assert normalize_time("--") is None
        assert normalize_time("00") is None
        assert normalize_time("0") is None
        assert normalize_time("nil") is None
        assert normalize_time("n/a") is None

    def test_normalize_time_quirks_and_cancellations(self):
        assert normalize_time("5 :30 am") == datetime.time(5, 30)
        assert normalize_time("10: 00 AM") == datetime.time(10, 0)
        assert normalize_time("02:oo PM") == datetime.time(14, 0)
        assert normalize_time("22 hrs") == datetime.time(22, 0)
        assert normalize_time("22hrs") == datetime.time(22, 0)
        assert normalize_time("Cancelled") is None
        assert normalize_time("Cab cancelled used for GIM") is None
        assert normalize_time("Cancelled after coming in campus") is None
        assert normalize_time("CASH") is None
        assert normalize_time("Paid to driver") is None

    def test_normalize_time_invalid(self):
        with pytest.raises(ValueError, match="not a valid time"):
            normalize_time("99:99")

    def test_normalize_boolean(self):
        assert normalize_boolean(True) is True
        assert normalize_boolean(False) is False
        assert normalize_boolean("true") is True
        assert normalize_boolean("false") is False
        assert normalize_boolean("yes") is True
        assert normalize_boolean("no") is False
        assert normalize_boolean("1") is True
        assert normalize_boolean("0") is False
        assert normalize_boolean(1) is True
        assert normalize_boolean(0) is False
        assert normalize_boolean(None) is False
        assert normalize_boolean("") is False
        assert normalize_boolean("received") is True

    def test_normalize_boolean_invalid(self):
        with pytest.raises(ValueError, match="not a valid boolean"):
            normalize_boolean("invalid_value")
        with pytest.raises(ValueError, match="not a valid boolean"):
            normalize_boolean(99)

    def test_normalize_cell_dispatch(self):
        col_str = ColumnDef(db_column="test_str", data_type="string", max_length=5)
        assert normalize_cell("  abcdefgh  ", col_str) == "abcde"

        col_int = ColumnDef(db_column="test_int", data_type="integer")
        assert normalize_cell("100", col_int) == 100

        col_bool = ColumnDef(db_column="test_bool", data_type="boolean")
        assert normalize_cell("yes", col_bool) is True
        assert normalize_cell(False, col_bool) is False
        assert normalize_cell(None, col_bool) is False

    def test_normalize_cell_default_value(self):
        col_with_default = ColumnDef(
            db_column="company", data_type="string", default_value="GIM"
        )
        assert normalize_cell(None, col_with_default) == "GIM"
        assert normalize_cell("", col_with_default) == "GIM"
        assert normalize_cell("   ", col_with_default) == "GIM"
        assert normalize_cell("Acme Corp", col_with_default) == "Acme Corp"

    def test_format_guest_message(self):
        from app.modules.bookings.schema import format_guest_message

        # Default company
        msg_default = format_guest_message(
            driver_name="Rahul",
            contact="8007029326",
            start_time_str="10:00 AM",
            pickup_location="Airport",
        )
        assert msg_default == (
            "From GIM\n"
            "Driver Rahul 8007029326 will come to pick you tomorrow at 10:00 AM at Airport"
        )

        # Custom company
        msg_custom = format_guest_message(
            driver_name="Abhishek",
            contact="8830087698",
            start_time_str="2:30 PM",
            pickup_location="Hotel Taj",
            company="Tata Sons",
        )
        assert msg_custom == (
            "From Tata Sons\n"
            "Driver Abhishek 8830087698 will come to pick you tomorrow at 2:30 PM at Hotel Taj"
        )

