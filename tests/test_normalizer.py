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
        assert normalize_date(None) is None
        assert normalize_date("  ") is None

    def test_normalize_date_invalid(self):
        with pytest.raises(ValueError, match="not a valid date"):
            normalize_date("invalid-date-xyz")

    def test_normalize_time(self):
        t = datetime.time(12, 0)
        dt = datetime.datetime(2026, 2, 1, 12, 0)
        assert normalize_time(t) == t
        assert normalize_time(dt) == t
        assert normalize_time("12:00") == t
        assert normalize_time("12:00:00") == t
        assert normalize_time("12:00 PM") == t
        assert normalize_time("12:00 AM") == datetime.time(0, 0)
        assert normalize_time(None) is None

    def test_normalize_time_invalid(self):
        with pytest.raises(ValueError, match="not a valid time"):
            normalize_time("99:99")

    def test_normalize_boolean(self):
        assert normalize_boolean(True) == "True"
        assert normalize_boolean(False) == "False"
        assert normalize_boolean("true") == "True"
        assert normalize_boolean("false") == "False"
        assert normalize_boolean("yes") == "True"
        assert normalize_boolean("no") == "False"
        assert normalize_boolean("1") == "True"
        assert normalize_boolean("0") == "False"
        assert normalize_boolean(None) == "False"

    def test_normalize_cell_dispatch(self):
        col_str = ColumnDef(db_column="test_str", data_type="string", max_length=5)
        assert normalize_cell("  abcdefgh  ", col_str) == "abcde"

        col_int = ColumnDef(db_column="test_int", data_type="integer")
        assert normalize_cell("100", col_int) == 100
