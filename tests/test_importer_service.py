"""Integration tests for ExcelImportService covering all 15 required scenarios."""

import datetime
from pathlib import Path
import pytest

from app.core.config import Config
from app.modules.bookings.repository import BookingRepository
from app.modules.drivers.repository import DriverRepository
from app.modules.importers.models import ImportResult
from app.modules.importers.service import ExcelImportService
from tests.conftest import MockDatabaseConnection


class TestExcelImportService:
    def _create_service(self, should_fail_db: bool = False, duplicate_mode: str = "skip", batch_size: int = 1000):
        config = Config(duplicate_mode=duplicate_mode, batch_size=batch_size)
        mock_db = MockDatabaseConnection(should_fail_on_execute=should_fail_db)
        driver_repo = DriverRepository(duplicate_mode=duplicate_mode, batch_size=batch_size)
        booking_repo = BookingRepository(duplicate_mode=duplicate_mode, batch_size=batch_size)
        service = ExcelImportService(
            config=config,
            db_conn=mock_db,
            driver_repo=driver_repo,
            booking_repo=booking_repo,
        )
        return service, mock_db

    # 1. One worksheet with valid data
    def test_single_sheet_valid_data(self, make_workbook, sample_booking_headers):
        rows = [
            sample_booking_headers,
            [False, 1, "Yash", "Guest 1", "9998887770", "2026-02-01", "Airport", "Hotel", "10:00 AM", "12:00 PM", "Meeting", "Abhishek", "8830087698"],
            [True, 2, "Yash", "Guest 2", "9998887771", "2026-02-01", "Hotel", "GIM", "02:00 PM", "04:00 PM", "Conference", "Rahul", "8007029326"],
        ]
        file_path = make_workbook({"Feb 01": rows})
        service, mock_db = self._create_service()

        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert res.rows_read == 2
        assert res.rows_validated == 2
        assert res.rows_inserted == 2
        assert res.rows_skipped == 0
        assert mock_db.mock_conn.committed
        assert not mock_db.mock_conn.rolled_back

    # 2. Multiple worksheets with identical schema (ALL go into the ONE bookings table)
    def test_multiple_sheets_identical_schema_one_table(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "LocA", "LocB", "10:00", "11:00", "R1", "Abhishek", "8830087698"],
                ],
                "Feb 02": [
                    sample_booking_headers,
                    [False, 2, "B", "G2", "222", "2026-02-02", "LocC", "LocD", "11:00", "12:00", "R2", "Rahul", "8007029326"],
                ],
                "Feb 03": [
                    sample_booking_headers,
                    [False, 3, "C", "G3", "333", "2026-02-03", "LocE", "LocF", "12:00", "13:00", "R3", "Arvind", "9822150660"],
                ],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert len(res.sheets_processed) == 3
        assert all(s.target_table == "bookings" for s in res.sheets_processed)
        assert res.rows_inserted == 3
        executed = [q for cur in mock_db.mock_conn.cursors for q in cur.executed_queries]
        assert any("INSERT" in q and "`bookings`" in q for q in executed)
        assert mock_db.mock_conn.committed

    # Contacts sheet -> drivers table, other sheets -> bookings table
    def test_contacts_and_bookings_together(self, make_workbook, sample_booking_headers, sample_contacts_headers):
        file_path = make_workbook(
            {
                "Contacts": [
                    sample_contacts_headers,
                    [1, "Abhishek", 8830087698],
                    [2, "Rahul", 8007029326],
                ],
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "Guest A", "999", "2026-02-01", "FromA", "ToA", "09:00", "10:00", "Trip", "Abhishek", "8830087698"],
                ],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert len(res.sheets_processed) == 2
        contacts_summary = next(s for s in res.sheets_processed if s.sheet_name == "Contacts")
        assert contacts_summary.target_table == "drivers"
        assert contacts_summary.rows_inserted == 2

        booking_summary = next(s for s in res.sheets_processed if s.sheet_name == "Feb 01")
        assert booking_summary.target_table == "bookings"
        assert booking_summary.rows_inserted == 1

        assert res.rows_inserted == 3

    # 3. Multiple worksheets where one sheet has a schema mismatch
    def test_multiple_worksheets_schema_mismatch_aborts(self, make_workbook, sample_booking_headers):
        mismatched_headers = [h for h in sample_booking_headers if h != "Reason"] + ["ExtraColumn"]
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "LocA", "LocB", "10:00", "11:00", "R1", "Abhishek", "8830087698"],
                ],
                "Feb 02": [
                    mismatched_headers,
                    [False, 2, "B", "G2", "222", "2026-02-02", "LocC", "LocD", "11:00", "12:00", "Abhishek", "8830087698", "ExtraVal"],
                ],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "FAILED"
        assert "Schema mismatch" in res.message or "unexpected column" in res.message.lower()
        assert res.rows_inserted == 0
        assert not mock_db.mock_conn.committed

    # 4. Empty worksheet safely skipped
    def test_empty_worksheet_safely_skipped(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "LocA", "LocB", "10:00", "11:00", "R1", "Abhishek", "8830087698"],
                ],
                "EmptySheet": [],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert len(res.sheets_processed) == 1
        assert res.rows_inserted == 1

    # 5. Blank rows skipped
    def test_blank_rows_skipped(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "LocA", "LocB", "10:00", "11:00", "R1", "Abhishek", "8830087698"],
                    [None] * len(sample_booking_headers),
                    [False] + [None] * (len(sample_booking_headers) - 1),
                ],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert res.rows_read == 3
        assert res.rows_validated == 1
        assert res.rows_skipped == 2
        assert res.rows_inserted == 1

    # 6. Empty cells handled safely
    def test_empty_cells_handled_safely(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, None, None, "Guest Only", None, "2026-02-01", None, None, None, None, None, "Abhishek", None],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert res.rows_inserted == 1

    # 7. Excel dates normalization
    def test_excel_dates_normalization(self, make_workbook, sample_booking_headers):
        py_date = datetime.date(2026, 2, 1)
        py_datetime = datetime.datetime(2026, 2, 2, 9, 30)
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", py_date, "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                    [False, 2, "B", "G2", "222", py_datetime, "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                    [False, 3, "C", "G3", "333", "03-02-2026", "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert res.rows_inserted == 3

    # 8. Numeric values & 0 preservation
    def test_numeric_values_and_preserve_zeros(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 0, "A", "G1", "111", "2026-02-01", "A", "B", "10:00", "11:00", "R", "Abhishek", "0"],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "SUCCESS"
        assert res.rows_inserted == 1
        cur = mock_db.mock_conn.cursors[-1]
        booking_params = cur.get_params_for_query("bookings")
        assert len(booking_params) == 1
        inserted_params = booking_params[0]
        assert inserted_params[1] == 0

    # 9. Invalid numeric value reports exact sheet and row
    def test_invalid_numeric_reports_sheet_and_row(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, "NOT_AN_INT", "A", "G1", "111", "2026-02-01", "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "FAILED"
        assert len(res.errors) >= 1
        err = res.errors[0]
        assert err.sheet_name == "Feb 01"
        assert err.row_number == 2
        assert err.column_name == "sl_no"
        assert mock_db.mock_conn.rolled_back

    # 10. Missing required value reports exact sheet and row
    def test_missing_required_value_reports_sheet_and_row(self, make_workbook, sample_contacts_headers):
        file_path = make_workbook(
            {
                "Contacts": [
                    sample_contacts_headers,
                    [1, "", "8830087698"],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "FAILED"
        err = res.errors[0]
        assert err.sheet_name == "Contacts"
        assert err.row_number == 2
        assert err.column_name == "driver"

    # 11. Duplicate headers detected
    def test_duplicate_headers_rejected(self, make_workbook):
        dup_headers = ["Received", "Name", "Name", "Guest", "Travel Date", "Driver", "Contact"]
        file_path = make_workbook(
            {
                "Feb 01": [
                    dup_headers,
                    [False, "Yash", "Yash", "Guest", "2026-02-01", "Abhishek", "8830087698"],
                ]
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path)

        assert res.status == "FAILED"
        assert "duplicate header" in res.message.lower()

    # 12. Large number of rows to verify batching
    def test_large_batching(self, make_workbook, sample_booking_headers):
        rows = [sample_booking_headers]
        for i in range(1, 1251):
            rows.append([False, i, f"Name{i}", f"Guest{i}", "9999999999", "2026-02-01", "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"])

        file_path = make_workbook({"Feb 01": rows})
        service, mock_db = self._create_service(batch_size=500)
        res = service.import_workbook(file_path, batch_size=500)

        assert res.status == "SUCCESS"
        assert res.rows_inserted == 1250
        booking_cur = [c for c in mock_db.mock_conn.cursors if c.executed_queries]
        assert len(booking_cur) >= 1
        booking_params = booking_cur[0].get_params_for_query("bookings")
        assert len(booking_params) == 1250

    # 13. Database failure causes rollback
    def test_database_failure_causes_rollback(self, make_workbook, sample_booking_headers):
        file_path = make_workbook(
            {
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                ]
            }
        )
        service, mock_db = self._create_service(should_fail_db=True)
        res = service.import_workbook(file_path)

        assert res.status == "FAILED"
        assert mock_db.mock_conn.rolled_back
        assert not mock_db.mock_conn.committed

    # 14. Dry-run causes zero database writes
    def test_dry_run_zero_database_writes(self, make_workbook, sample_booking_headers, sample_contacts_headers):
        file_path = make_workbook(
            {
                "Contacts": [
                    sample_contacts_headers,
                    [1, "Abhishek", 8830087698],
                ],
                "Feb 01": [
                    sample_booking_headers,
                    [False, 1, "A", "G1", "111", "2026-02-01", "A", "B", "10:00", "11:00", "R", "Abhishek", "8830087698"],
                ],
            }
        )
        service, mock_db = self._create_service()
        res = service.import_workbook(file_path, dry_run=True)

        assert res.status == "DRY_RUN"
        assert res.rows_validated == 2
        assert res.database_writes == 0
        assert res.rows_inserted == 0
        assert len(mock_db.mock_conn.cursors) == 0

    # 15. Duplicate / unique-key behavior (skip, upsert, fail)
    def test_duplicate_key_modes(self):
        repo_skip = DriverRepository(duplicate_mode="skip")
        repo_upsert = DriverRepository(duplicate_mode="upsert")
        repo_fail = DriverRepository(duplicate_mode="fail")

        from tests.conftest import MockCursor
        cur_skip = MockCursor()
        repo_skip.insert_batch(cur_skip, [{"driver": "Rahul", "contact": "8007029326"}])
        assert "ON DUPLICATE KEY UPDATE `sr_no` = `sr_no`" in cur_skip.executed_queries[0]

        cur_upsert = MockCursor()
        repo_upsert.insert_batch(cur_upsert, [{"driver": "Rahul", "contact": "8007029326"}])
        assert "ON DUPLICATE KEY UPDATE `contact` = VALUES(`contact`)" in cur_upsert.executed_queries[0]

        cur_fail = MockCursor()
        repo_fail.insert_batch(cur_fail, [{"driver": "Rahul", "contact": "8007029326"}])
        assert "ON DUPLICATE" not in cur_fail.executed_queries[0]
