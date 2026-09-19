"""Data models for import results, sheet summaries, and row errors."""

from dataclasses import dataclass, field
from typing import Any, List, Optional


@dataclass
class RowError:
    """Represents a validation or database error on a specific row."""

    sheet_name: str
    row_number: int  # 1-based Excel row index
    column_name: Optional[str] = None
    error_type: str = "VALIDATION_ERROR"
    message: str = ""
    raw_value: Any = None

    def __str__(self) -> str:
        col_part = f" column '{self.column_name}'" if self.column_name else ""
        val_part = f" (raw value: {self.raw_value!r})" if self.raw_value is not None else ""
        return f"Sheet '{self.sheet_name}' Excel row {self.row_number}{col_part}: {self.message}{val_part}"


@dataclass
class SheetSummary:
    """Import summary for a single worksheet."""

    sheet_name: str
    target_table: str
    rows_read: int = 0
    rows_validated: int = 0
    rows_inserted: int = 0
    rows_skipped: int = 0
    errors: List[RowError] = field(default_factory=list)


@dataclass
class ImportResult:
    """Overall result of an Excel workbook import operation."""

    filename: str
    sheets_processed: List[SheetSummary] = field(default_factory=list)
    total_sheets: int = 0
    rows_read: int = 0
    rows_validated: int = 0
    rows_inserted: int = 0
    rows_skipped: int = 0
    database_writes: int = 0
    errors: List[RowError] = field(default_factory=list)
    duration_seconds: float = 0.0
    status: str = "SUCCESS"  # SUCCESS, FAILED, DRY_RUN
    message: str = ""

    def summary_text(self) -> str:
        """Format an informative text summary of the import result."""
        lines = [
            "=" * 60,
            "Excel-to-MySQL Import Result",
            "=" * 60,
            f"File:             {self.filename}",
            f"Status:           {self.status}",
            f"Sheets Processed: {len(self.sheets_processed)} (total discovered: {self.total_sheets})",
            f"Rows Read:        {self.rows_read:,}",
            f"Rows Validated:   {self.rows_validated:,}",
            f"Rows Inserted:    {self.rows_inserted:,}",
            f"Rows Skipped:     {self.rows_skipped:,}",
            f"Database Writes:  {self.database_writes:,}",
            f"Duration:         {self.duration_seconds:.2f}s",
        ]
        if self.message:
            lines.append(f"Message:          {self.message}")

        if self.sheets_processed:
            lines.append("-" * 60)
            lines.append(f"{'Sheet Name':<18} | {'Table':<10} | {'Read':<6} | {'Valid':<6} | {'Inserted':<8} | {'Skipped':<7}")
            lines.append("-" * 60)
            for s in self.sheets_processed:
                lines.append(
                    f"{s.sheet_name:<18} | {s.target_table:<10} | {s.rows_read:<6} | {s.rows_validated:<6} | {s.rows_inserted:<8} | {s.rows_skipped:<7}"
                )

        if self.errors:
            lines.append("-" * 60)
            lines.append(f"Errors Encountered ({len(self.errors)}):")
            for err in self.errors[:10]:
                lines.append(f"  * {err}")
            if len(self.errors) > 10:
                lines.append(f"  ... and {len(self.errors) - 10} more errors.")
        lines.append("=" * 60)
        return "\n".join(lines)
