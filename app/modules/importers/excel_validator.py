"""Header and row validation utilities for Excel worksheets."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from app.common.normalizer import normalize_cell
from app.common.schema_base import TableSchema
from app.modules.bookings.schema import BOOKINGS_SCHEMA
from app.modules.drivers.schema import DRIVERS_SCHEMA, is_contacts_sheet
from app.modules.importers.models import RowError


@dataclass
class HeaderValidationResult:
    """Outcome of validating a worksheet's header row."""

    sheet_name: str
    is_valid: bool
    headers: List[str] = field(default_factory=list)
    missing_columns: List[str] = field(default_factory=list)
    unexpected_columns: List[str] = field(default_factory=list)
    duplicate_headers: List[str] = field(default_factory=list)
    col_index_to_db_col: Dict[int, str] = field(default_factory=dict)
    error_message: str = ""


class WorksheetValidator:
    """Validates worksheet headers and row content against expected database schemas."""

    def __init__(self):
        self.first_booking_headers: Optional[List[str]] = None
        self.first_booking_sheet_name: Optional[str] = None

    def validate_headers(
        self, sheet_name: str, raw_headers: Sequence[Any]
    ) -> HeaderValidationResult:
        """Validate header row of a worksheet against the appropriate schema."""
        trimmed_headers: List[str] = []
        for h in raw_headers:
            if h is None:
                trimmed_headers.append("")
            else:
                trimmed_headers.append(str(h).strip())

        # Strip trailing empty headers (formula columns without headers)
        while trimmed_headers and trimmed_headers[-1] == "":
            trimmed_headers.pop()

        if not trimmed_headers:
            return HeaderValidationResult(
                sheet_name=sheet_name,
                is_valid=False,
                error_message=f"Sheet '{sheet_name}' has an empty header row.",
            )

        # Check for duplicate headers
        seen_headers: Set[str] = set()
        duplicates: List[str] = []
        for idx, h in enumerate(trimmed_headers):
            if not h:
                continue
            h_lower = h.lower()
            if h_lower in seen_headers and h_lower not in duplicates:
                duplicates.append(h)
            seen_headers.add(h_lower)

        if duplicates:
            return HeaderValidationResult(
                sheet_name=sheet_name,
                is_valid=False,
                headers=trimmed_headers,
                duplicate_headers=duplicates,
                error_message=f"Sheet '{sheet_name}' contains duplicate header(s): {', '.join(duplicates)}",
            )

        # Determine target schema
        if is_contacts_sheet(sheet_name):
            schema = DRIVERS_SCHEMA
        else:
            schema = BOOKINGS_SCHEMA

        col_index_to_db_col: Dict[int, str] = {}
        recognized_db_cols: Set[str] = set()
        unexpected_columns: List[str] = []

        for idx, h in enumerate(trimmed_headers):
            if not h:
                continue
            h_lower = h.lower()
            if h_lower in schema.header_aliases:
                db_col = schema.header_aliases[h_lower]
                if db_col in schema.columns:
                    col_index_to_db_col[idx] = db_col
                    recognized_db_cols.add(db_col)
            elif h_lower in schema.generated_columns:
                continue
            else:
                unexpected_columns.append(h)

        missing_columns: List[str] = []
        for req_col in schema.required_headers:
            if req_col not in recognized_db_cols:
                missing_columns.append(req_col)

        error_parts: List[str] = []
        if missing_columns:
            error_parts.append(f"missing required column(s): {', '.join(missing_columns)}")
        if unexpected_columns:
            error_parts.append(f"unexpected column(s): {', '.join(unexpected_columns)}")

        # Inter-sheet booking consistency validation
        if schema.table_name == "bookings":
            if self.first_booking_headers is None:
                self.first_booking_headers = trimmed_headers
                self.first_booking_sheet_name = sheet_name
            else:
                norm_current = [h.lower() for h in trimmed_headers if h]
                norm_first = [h.lower() for h in self.first_booking_headers if h]
                diff_missing = [h for h in norm_first if h not in norm_current]
                diff_extra = [h for h in norm_current if h not in norm_first]
                if diff_missing or diff_extra:
                    mismatch_parts = []
                    if diff_missing:
                        mismatch_parts.append(f"missing {diff_missing} compared to sheet '{self.first_booking_sheet_name}'")
                    if diff_extra:
                        mismatch_parts.append(f"extra {diff_extra} compared to sheet '{self.first_booking_sheet_name}'")
                    error_parts.append(f"Schema mismatch against first booking sheet '{self.first_booking_sheet_name}': {'; '.join(mismatch_parts)}")

        is_valid = len(error_parts) == 0
        error_msg = f"Sheet '{sheet_name}': {'; '.join(error_parts)}" if error_parts else ""

        return HeaderValidationResult(
            sheet_name=sheet_name,
            is_valid=is_valid,
            headers=trimmed_headers,
            missing_columns=missing_columns,
            unexpected_columns=unexpected_columns,
            duplicate_headers=duplicates,
            col_index_to_db_col=col_index_to_db_col,
            error_message=error_msg,
        )

    @staticmethod
    def is_blank_row(
        raw_row: Sequence[Any],
        col_index_to_db_col: Dict[int, str],
        is_contacts: bool = False,
    ) -> bool:
        """Determine if a row is completely blank or contains only default unedited flags."""
        if not raw_row or all(c is None for c in raw_row):
            return True

        substantive_values = 0
        for idx, val in enumerate(raw_row):
            if idx not in col_index_to_db_col:
                continue
            db_col = col_index_to_db_col[idx]
            if db_col == "received":
                continue
            if val is not None and str(val).strip() != "" and str(val).strip() != "0":
                substantive_values += 1

        return substantive_values == 0

    @staticmethod
    def validate_and_normalize_row(
        sheet_name: str,
        row_number: int,
        raw_row: Sequence[Any],
        header_val: HeaderValidationResult,
        schema: TableSchema,
    ) -> Tuple[Optional[Dict[str, Any]], List[RowError]]:
        """Validate and normalize cell values for a single row."""
        errors: List[RowError] = []
        normalized_row: Dict[str, Any] = {}

        for col_idx, db_col in header_val.col_index_to_db_col.items():
            raw_val = raw_row[col_idx] if col_idx < len(raw_row) else None
            col_def = schema.columns.get(db_col)
            if not col_def:
                continue

            try:
                norm_val = normalize_cell(raw_val, col_def)
            except ValueError as e:
                errors.append(
                    RowError(
                        sheet_name=sheet_name,
                        row_number=row_number,
                        column_name=db_col,
                        error_type="INVALID_TYPE",
                        message=str(e),
                        raw_value=raw_val,
                    )
                )
                continue

            if col_def.is_required and (norm_val is None or str(norm_val).strip() == ""):
                errors.append(
                    RowError(
                        sheet_name=sheet_name,
                        row_number=row_number,
                        column_name=db_col,
                        error_type="MISSING_REQUIRED",
                        message=f"Required field '{db_col}' cannot be null or empty",
                        raw_value=raw_val,
                    )
                )
                continue

            normalized_row[db_col] = norm_val

        if errors:
            return None, errors

        return normalized_row, []
