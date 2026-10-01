"""Unit tests for worksheet header and row validation."""

from app.modules.bookings.schema import BOOKINGS_SCHEMA
from app.modules.drivers.schema import DRIVERS_SCHEMA
from app.modules.importers.excel_validator import WorksheetValidator


class TestValidator:
    def test_validate_valid_booking_headers(self, sample_booking_headers):
        val = WorksheetValidator()
        res = val.validate_headers("Feb 01", sample_booking_headers)
        assert res.is_valid
        assert len(res.missing_columns) == 0
        assert len(res.unexpected_columns) == 0
        assert len(res.duplicate_headers) == 0

    def test_validate_valid_contacts_headers(self, sample_contacts_headers):
        val = WorksheetValidator()
        res = val.validate_headers("Contacts", sample_contacts_headers)
        assert res.is_valid
        assert len(res.missing_columns) == 0
        assert len(res.unexpected_columns) == 0

    def test_detect_duplicate_headers(self):
        val = WorksheetValidator()
        headers = ["Received", "Name", "Name", "Guest", "Travel Date", "Driver"]
        res = val.validate_headers("Feb 01", headers)
        assert not res.is_valid
        assert "Name" in res.duplicate_headers
        assert "duplicate header" in res.error_message.lower()

    def test_detect_missing_and_unexpected_columns(self):
        val = WorksheetValidator()
        headers = ["Received", "Guest", "unknown_col"]
        res = val.validate_headers("Feb 01", headers)
        assert not res.is_valid
        assert "unknown_col" in res.unexpected_columns
        assert any(c in res.missing_columns for c in ["travel_date", "driver_name"])

    def test_schema_mismatch_between_booking_sheets(self, sample_booking_headers):
        val = WorksheetValidator()
        res1 = val.validate_headers("Feb 01", sample_booking_headers)
        assert res1.is_valid

        sheet2_headers = [h for h in sample_booking_headers if h != "Reason"] + ["Department"]
        res2 = val.validate_headers("Feb 02", sheet2_headers)
        assert not res2.is_valid
        assert "mismatch against first booking sheet" in res2.error_message.lower()

    def test_is_blank_row(self, sample_booking_headers):
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", sample_booking_headers)

        assert WorksheetValidator.is_blank_row([None] * len(sample_booking_headers), h_res.col_index_to_db_col)

        empty_template_row = [False] + [None] * (len(sample_booking_headers) - 1)
        assert WorksheetValidator.is_blank_row(empty_template_row, h_res.col_index_to_db_col)

        data_row = [False, 1, "Staff", "Guest A", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326"]
        assert not WorksheetValidator.is_blank_row(data_row, h_res.col_index_to_db_col)

    def test_validate_and_normalize_row_valid(self, sample_booking_headers):
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", sample_booking_headers)
        raw_row = [False, 1, "Staff", "Guest A", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326"]

        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Feb 01", 2, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert len(errors) == 0
        assert norm_row is not None
        assert norm_row["guest_name"] == "Guest A"
        assert norm_row["sl_no"] == 1
        assert norm_row["driver_name"] == "Rahul"
        assert norm_row["received"] is False
        assert norm_row["coordinator"] == "Staff"

    def test_validate_and_normalize_row_with_coordinator_and_vehicle_type(self):
        headers = [
            "Received",
            "Sl.No.",
            "Coordinator",
            "Guest",
            "Guest Mobile No.",
            "Travel Date",
            "From",
            "To",
            "Start Time",
            "Return Time",
            "Reason",
            "Driver",
            "Contact",
            "Vehicle Type",
        ]
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", headers)
        assert h_res.is_valid
        raw_row = [True, 5, "Manager", "Guest B", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326", "Sedan"]

        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Feb 01", 2, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert len(errors) == 0
        assert norm_row is not None
        assert norm_row["coordinator"] == "Manager"
        assert norm_row["vehicle_type"] == "Sedan"
        assert norm_row["received"] is True

    def test_validate_and_normalize_row_received_truthy(self, sample_booking_headers):
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", sample_booking_headers)
        raw_row = ["received", 1, "Staff", "Guest A", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326"]

        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Feb 01", 2, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert len(errors) == 0
        assert norm_row is not None
        assert norm_row["received"] is True

    def test_validate_and_normalize_row_invalid_boolean_reports_location(self, sample_booking_headers):
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", sample_booking_headers)
        raw_row = ["MAYBE", 1, "Staff", "Guest A", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326"]

        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Feb 01", 15, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert norm_row is None
        assert len(errors) == 1
        err = errors[0]
        assert err.sheet_name == "Feb 01"
        assert err.row_number == 15
        assert err.column_name == "received"
        assert "not a valid boolean" in err.message

    def test_validate_and_normalize_row_invalid_data_reports_location(self, sample_booking_headers):
        val = WorksheetValidator()
        h_res = val.validate_headers("Feb 01", sample_booking_headers)
        raw_row = [False, "NOT_A_NUMBER", "Staff", "Guest A", "9876543210", "2026-02-01", "A", "B", "10:00", "12:00", "Duty", "Rahul", "8007029326"]

        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Feb 01", 42, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert norm_row is None
        assert len(errors) == 1
        err = errors[0]
        assert err.sheet_name == "Feb 01"
        assert err.row_number == 42
        assert err.column_name == "sl_no"
        assert "not a valid integer" in err.message

    def test_validate_headers_with_company(self, sample_booking_headers):
        headers_with_company = ["Company"] + sample_booking_headers
        val = WorksheetValidator()
        res = val.validate_headers("Feb 01", headers_with_company)
        assert res.is_valid
        assert "company" in res.col_index_to_db_col.values()

    def test_company_is_optional_across_sheets(self, sample_booking_headers):
        val = WorksheetValidator()
        # Sheet 1 has Company
        headers1 = ["Company"] + sample_booking_headers
        res1 = val.validate_headers("Feb 01", headers1)
        assert res1.is_valid

        # Sheet 2 lacks Company (should still pass because company is in optional_cols)
        res2 = val.validate_headers("Feb 02", sample_booking_headers)
        assert res2.is_valid

    def test_jan_2026_canonical_headers(self):
        from app.modules.bookings.schema import CANONICAL_BOOKING_HEADERS
        assert CANONICAL_BOOKING_HEADERS[0] == "Received"
        assert CANONICAL_BOOKING_HEADERS[1] == "Sl.No."
        assert CANONICAL_BOOKING_HEADERS[-1] == "Company"

    def test_jan_2026_vech_no_and_vendor_aliases(self):
        val = WorksheetValidator()
        jan_headers = [
            "Received",
            "Sl.No.",
            "Coordinator",
            "Guest",
            "Guest Mobile No.",
            "Travel Date",
            "From",
            "To",
            "Start Time",
            "Return Time",
            "Reason",
            "Driver",
            "Contact",
            "Vech No",
            "Vendor",
        ]
        res = val.validate_headers("Jan 01", jan_headers)
        assert res.is_valid
        # "Vech No" maps to "vehicle_type"
        assert res.col_index_to_db_col[13] == "vehicle_type"
        # "Vendor" maps to "company"
        assert res.col_index_to_db_col[14] == "company"

    def test_validate_and_normalize_row_with_vech_no_and_vendor(self):
        val = WorksheetValidator()
        jan_headers = [
            "Received",
            "Sl.No.",
            "Coordinator",
            "Guest",
            "Guest Mobile No.",
            "Travel Date",
            "From",
            "To",
            "Start Time",
            "Return Time",
            "Reason",
            "Driver",
            "Contact",
            "Vech No",
            "Vendor",
        ]
        h_res = val.validate_headers("Jan 01", jan_headers)
        raw_row = [
            True,
            1,
            "Manager",
            "Guest C",
            "9876543210",
            "2026-01-01",
            "Airport",
            "GIM",
            "10:00",
            "12:00",
            "Meeting",
            "Sadiq",
            "8296274034",
            "GA 07 T 1586",
            "Lalu",
        ]
        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Jan 01", 2, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert len(errors) == 0
        assert norm_row is not None
        assert norm_row["vehicle_type"] == "GA 07 T 1586"
        assert norm_row["company"] == "Lalu"
        assert norm_row["received"] is True

    def test_validate_and_normalize_row_combines_vech_no_and_type(self):
        val = WorksheetValidator()
        headers = [
            "Received",
            "Sl.No.",
            "Coordinator",
            "Guest",
            "Guest Mobile No.",
            "Travel Date",
            "From",
            "To",
            "Start Time",
            "Return Time",
            "Reason",
            "Driver",
            "Contact",
            "Vehicle Number",
            "Type",
        ]
        h_res = val.validate_headers("Jan 01", headers)
        assert h_res.is_valid
        raw_row = [
            True,
            1,
            "Manager",
            "Guest C",
            "9876543210",
            "2026-01-01",
            "Airport",
            "GIM",
            "10:00",
            "12:00",
            "Meeting",
            "Sadiq",
            "8296274034",
            "GA 07 T 1586",
            "Ertiga",
        ]
        norm_row, errors = WorksheetValidator.validate_and_normalize_row(
            "Jan 01", 2, raw_row, h_res, BOOKINGS_SCHEMA
        )
        assert len(errors) == 0
        assert norm_row is not None
        assert norm_row["vehicle_type"] == "GA 07 T 1586 Ertiga"

