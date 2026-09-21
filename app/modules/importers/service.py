"""Excel import service orchestrating schema validation, row normalization, and transactional database writes."""

import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

from app.core.config import Config
from app.core.database import DatabaseConnection
from app.modules.bookings.repository import BookingRepository
from app.modules.bookings.schema import BOOKINGS_SCHEMA
from app.modules.drivers.repository import DriverRepository
from app.modules.drivers.schema import DRIVERS_SCHEMA, is_contacts_sheet
from app.modules.importers.excel_reader import ExcelWorkbookReader
from app.modules.importers.excel_validator import HeaderValidationResult, WorksheetValidator
from app.modules.importers.models import ImportResult, RowError, SheetSummary

logger = logging.getLogger("app.modules.importers.service")


class ExcelImportService:
    """Service for validating and importing multi-sheet Excel workbooks into MySQL."""

    def __init__(
        self,
        config: Optional[Config] = None,
        db_conn: Optional[DatabaseConnection] = None,
        driver_repo: Optional[DriverRepository] = None,
        booking_repo: Optional[BookingRepository] = None,
    ):
        self.config = config or Config.from_env()
        self.db = db_conn or DatabaseConnection(self.config)
        self.driver_repo = driver_repo or DriverRepository(
            duplicate_mode=self.config.duplicate_mode,
            batch_size=self.config.batch_size,
        )
        self.booking_repo = booking_repo or BookingRepository(
            duplicate_mode=self.config.duplicate_mode,
            batch_size=self.config.batch_size,
        )

    def import_workbook(
        self,
        file_path: Union[str, Path, Any],
        dry_run: bool = False,
        batch_size: Optional[int] = None,
        duplicate_mode: Optional[str] = None,
        reverse_sheets: Optional[bool] = None,
        strict_validation: bool = True,
    ) -> ImportResult:
        """Import an Excel workbook into MySQL."""
        start_time = time.perf_counter()
        effective_batch_size = batch_size or self.config.batch_size
        effective_duplicate_mode = duplicate_mode or self.config.duplicate_mode
        effective_reverse_sheets = self.config.reverse_sheets if reverse_sheets is None else reverse_sheets

        if effective_batch_size != self.driver_repo.batch_size or effective_duplicate_mode != self.driver_repo.duplicate_mode:
            self.driver_repo = DriverRepository(
                duplicate_mode=effective_duplicate_mode,
                batch_size=effective_batch_size,
            )
            self.booking_repo = BookingRepository(
                duplicate_mode=effective_duplicate_mode,
                batch_size=effective_batch_size,
            )

        filename_str = str(file_path)
        logger.info("Starting Excel import for '%s' (dry_run=%s)", filename_str, dry_run)

        result = ImportResult(
            filename=filename_str,
            status="SUCCESS",
        )

        try:
            with ExcelWorkbookReader(file_path) as reader:
                sheet_names = reader.sheet_names
                result.total_sheets = len(sheet_names)

                if not sheet_names:
                    result.status = "FAILED"
                    result.message = "Workbook contains no worksheets."
                    result.duration_seconds = time.perf_counter() - start_time
                    return result

                # -----------------------------------------------------------
                # 1. Categorize sheets & Pre-validate all headers
                # -----------------------------------------------------------
                contacts_sheet_name: Optional[str] = None
                booking_sheet_names: List[str] = []

                for name in sheet_names:
                    if is_contacts_sheet(name):
                        if contacts_sheet_name is None:
                            contacts_sheet_name = name
                        else:
                            logger.warning("Multiple contacts sheets found; using first: '%s'", contacts_sheet_name)
                    else:
                        booking_sheet_names.append(name)

                # Process booking sheets in reverse order of the workbook (e.g. Day 1 -> Day 31)
                if effective_reverse_sheets:
                    booking_sheet_names.reverse()

                validator = WorksheetValidator()
                header_validations: Dict[str, HeaderValidationResult] = {}
                preval_errors: List[RowError] = []

                if contacts_sheet_name:
                    if not reader.is_sheet_empty(contacts_sheet_name):
                        c_headers = reader.read_headers(contacts_sheet_name)
                        c_val = validator.validate_headers(contacts_sheet_name, c_headers)
                        header_validations[contacts_sheet_name] = c_val
                        if not c_val.is_valid:
                            preval_errors.append(
                                RowError(
                                    sheet_name=contacts_sheet_name,
                                    row_number=1,
                                    error_type="SCHEMA_MISMATCH",
                                    message=c_val.error_message,
                                )
                            )

                for b_sheet in booking_sheet_names:
                    if reader.is_sheet_empty(b_sheet):
                        logger.info("Sheet '%s' is empty, skipping.", b_sheet)
                        continue
                    b_headers = reader.read_headers(b_sheet)
                    b_val = validator.validate_headers(b_sheet, b_headers)
                    header_validations[b_sheet] = b_val
                    if not b_val.is_valid:
                        preval_errors.append(
                            RowError(
                                sheet_name=b_sheet,
                                row_number=1,
                                error_type="SCHEMA_MISMATCH",
                                message=b_val.error_message,
                            )
                        )

                if preval_errors:
                    result.status = "FAILED"
                    result.errors.extend(preval_errors)
                    result.message = f"Schema validation failed: {preval_errors[0].message}"
                    result.duration_seconds = time.perf_counter() - start_time
                    logger.error("Pre-validation failed with %d error(s). Aborting import.", len(preval_errors))
                    return result

                # -----------------------------------------------------------
                # 2. DRY-RUN MODE
                # -----------------------------------------------------------
                if dry_run:
                    return self._execute_dry_run(
                        reader=reader,
                        contacts_sheet=contacts_sheet_name,
                        booking_sheets=booking_sheet_names,
                        header_validations=header_validations,
                        result=result,
                        start_time=start_time,
                    )

                # -----------------------------------------------------------
                # 3. LIVE TRANSACTIONAL IMPORT
                # -----------------------------------------------------------
                return self._execute_live_import(
                    reader=reader,
                    contacts_sheet=contacts_sheet_name,
                    booking_sheets=booking_sheet_names,
                    header_validations=header_validations,
                    result=result,
                    start_time=start_time,
                    strict_validation=strict_validation,
                )

        except Exception as e:
            logger.exception("Unexpected error during import: %s", e)
            result.status = "FAILED"
            result.message = str(e)
            result.errors.append(
                RowError(
                    sheet_name="N/A",
                    row_number=0,
                    error_type="SYSTEM_ERROR",
                    message=str(e),
                )
            )
            result.duration_seconds = time.perf_counter() - start_time
            return result

    def _execute_dry_run(
        self,
        reader: ExcelWorkbookReader,
        contacts_sheet: Optional[str],
        booking_sheets: List[str],
        header_validations: Dict[str, HeaderValidationResult],
        result: ImportResult,
        start_time: float,
    ) -> ImportResult:
        """Stream and validate all rows across all sheets without touching the database."""
        result.status = "DRY_RUN"
        result.database_writes = 0

        if contacts_sheet and contacts_sheet in header_validations:
            c_val = header_validations[contacts_sheet]
            summary = SheetSummary(sheet_name=contacts_sheet, target_table="drivers")
            for row_num, raw_row in reader.stream_rows(contacts_sheet, min_row=2):
                summary.rows_read += 1
                if WorksheetValidator.is_blank_row(raw_row, c_val.col_index_to_db_col):
                    summary.rows_skipped += 1
                    continue
                norm_row, errors = WorksheetValidator.validate_and_normalize_row(
                    contacts_sheet, row_num, raw_row, c_val, DRIVERS_SCHEMA
                )
                if errors:
                    summary.errors.extend(errors)
                    result.errors.extend(errors)
                else:
                    summary.rows_validated += 1
            result.sheets_processed.append(summary)

        for b_sheet in booking_sheets:
            if b_sheet not in header_validations:
                continue
            b_val = header_validations[b_sheet]
            summary = SheetSummary(sheet_name=b_sheet, target_table="bookings")
            for row_num, raw_row in reader.stream_rows(b_sheet, min_row=2):
                summary.rows_read += 1
                if WorksheetValidator.is_blank_row(raw_row, b_val.col_index_to_db_col):
                    summary.rows_skipped += 1
                    continue
                norm_row, errors = WorksheetValidator.validate_and_normalize_row(
                    b_sheet, row_num, raw_row, b_val, BOOKINGS_SCHEMA
                )
                if errors:
                    summary.errors.extend(errors)
                    result.errors.extend(errors)
                else:
                    summary.rows_validated += 1
            result.sheets_processed.append(summary)

        for s in result.sheets_processed:
            result.rows_read += s.rows_read
            result.rows_validated += s.rows_validated
            result.rows_skipped += s.rows_skipped

        result.duration_seconds = time.perf_counter() - start_time
        result.message = f"Dry run completed successfully. Validated {result.rows_validated} row(s) with {len(result.errors)} error(s)."
        return result

    def _execute_live_import(
        self,
        reader: ExcelWorkbookReader,
        contacts_sheet: Optional[str],
        booking_sheets: List[str],
        header_validations: Dict[str, HeaderValidationResult],
        result: ImportResult,
        start_time: float,
        strict_validation: bool = True,
    ) -> ImportResult:
        """Perform transactional batch import into MySQL."""
        with self.db.transaction() as cursor:
            cursor.execute("SET FOREIGN_KEY_CHECKS = 0;")

            # 1. Process and insert Contacts sheet into 'drivers' first (ONLY source for drivers)
            known_drivers: Set[str] = set()
            if contacts_sheet and contacts_sheet in header_validations:
                c_val = header_validations[contacts_sheet]
                summary = SheetSummary(sheet_name=contacts_sheet, target_table="drivers")
                driver_batch: List[Dict[str, Any]] = []

                for row_num, raw_row in reader.stream_rows(contacts_sheet, min_row=2):
                    summary.rows_read += 1
                    if WorksheetValidator.is_blank_row(raw_row, c_val.col_index_to_db_col):
                        summary.rows_skipped += 1
                        continue

                    norm_row, errors = WorksheetValidator.validate_and_normalize_row(
                        contacts_sheet, row_num, raw_row, c_val, DRIVERS_SCHEMA
                    )
                    if errors:
                        summary.errors.extend(errors)
                        result.errors.extend(errors)
                        if strict_validation:
                            raise ValueError(f"Validation failed on {errors[0]}")
                    else:
                        summary.rows_validated += 1
                        if norm_row and norm_row.get("driver"):
                            known_drivers.add(norm_row["driver"])
                            driver_batch.append(norm_row)

                    if len(driver_batch) >= self.driver_repo.batch_size:
                        inserted = self.driver_repo.insert_batch(cursor, driver_batch)
                        summary.rows_inserted += inserted
                        driver_batch.clear()

                if driver_batch:
                    inserted = self.driver_repo.insert_batch(cursor, driver_batch)
                    summary.rows_inserted += inserted
                    driver_batch.clear()

                result.sheets_processed.append(summary)
                result.database_writes += summary.rows_inserted

            # 2. Process all Booking sheets into 'bookings' table
            for b_sheet in booking_sheets:
                if b_sheet not in header_validations:
                    continue
                b_val = header_validations[b_sheet]
                summary = SheetSummary(sheet_name=b_sheet, target_table="bookings")
                booking_batch: List[Dict[str, Any]] = []

                for row_num, raw_row in reader.stream_rows(b_sheet, min_row=2):
                    summary.rows_read += 1
                    if WorksheetValidator.is_blank_row(raw_row, b_val.col_index_to_db_col):
                        summary.rows_skipped += 1
                        continue

                    norm_row, errors = WorksheetValidator.validate_and_normalize_row(
                        b_sheet, row_num, raw_row, b_val, BOOKINGS_SCHEMA
                    )
                    if errors:
                        summary.errors.extend(errors)
                        result.errors.extend(errors)
                        if strict_validation:
                            raise ValueError(f"Validation failed on {errors[0]}")
                    else:
                        summary.rows_validated += 1
                        if norm_row:
                            booking_batch.append(norm_row)

                    if len(booking_batch) >= self.booking_repo.batch_size:
                        inserted = self.booking_repo.insert_batch(cursor, booking_batch)
                        summary.rows_inserted += inserted
                        booking_batch.clear()

                if booking_batch:
                    inserted = self.booking_repo.insert_batch(cursor, booking_batch)
                    summary.rows_inserted += inserted
                    booking_batch.clear()

                result.sheets_processed.append(summary)
                result.database_writes += summary.rows_inserted

            cursor.execute("SET FOREIGN_KEY_CHECKS = 1;")

        for s in result.sheets_processed:
            result.rows_read += s.rows_read
            result.rows_validated += s.rows_validated
            result.rows_inserted += s.rows_inserted
            result.rows_skipped += s.rows_skipped

        result.duration_seconds = time.perf_counter() - start_time
        result.status = "SUCCESS"
        result.message = f"Successfully imported {result.rows_inserted} row(s) into MySQL across {len(result.sheets_processed)} sheet(s)."
        return result
