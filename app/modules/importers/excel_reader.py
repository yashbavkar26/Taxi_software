"""Streaming Excel workbook reader using openpyxl in read_only mode."""

from pathlib import Path
from typing import Any, Generator, List, Optional, Tuple, Union
import openpyxl
from openpyxl.workbook import Workbook


class ExcelWorkbookReader:
    """Streaming Excel reader using openpyxl with read_only=True and data_only=True.

    Ensures minimal memory usage and safe resource cleanup.
    """

    def __init__(self, file_source: Union[str, Path, Any]):
        self.file_source = file_source
        self.workbook: Optional[Workbook] = None

    def __enter__(self) -> "ExcelWorkbookReader":
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def open(self) -> None:
        """Open the workbook in read-only and data-only mode."""
        try:
            self.workbook = openpyxl.load_workbook(
                self.file_source,
                read_only=True,
                data_only=True,
            )
        except Exception as e:
            raise ValueError(f"Failed to open Excel workbook '{self.file_source}': {e}") from e

        if not self.workbook.sheetnames:
            raise ValueError(f"Workbook '{self.file_source}' contains no worksheets.")

    def close(self) -> None:
        """Close the workbook to free file handles and resources."""
        if self.workbook is not None:
            try:
                self.workbook.close()
            except Exception:
                pass
            self.workbook = None

    @property
    def sheet_names(self) -> List[str]:
        """Return list of sheet names in the workbook."""
        if self.workbook is None:
            raise RuntimeError("Workbook is not opened.")
        return list(self.workbook.sheetnames)

    def read_headers(self, sheet_name: str) -> Tuple[Any, ...]:
        """Read the first row (headers) of a given worksheet."""
        if self.workbook is None:
            raise RuntimeError("Workbook is not opened.")
        ws = self.workbook[sheet_name]
        for row in ws.iter_rows(max_row=1, values_only=True):
            return row
        return ()

    def is_sheet_empty(self, sheet_name: str) -> bool:
        """Check if sheet has no rows or only completely empty cells."""
        if self.workbook is None:
            raise RuntimeError("Workbook is not opened.")
        ws = self.workbook[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                return False
        return True

    def stream_rows(
        self, sheet_name: str, min_row: int = 2
    ) -> Generator[Tuple[int, Tuple[Any, ...]], None, None]:
        """Yield (1-based row_number, row_values) starting from min_row.

        Memory overhead is O(1) per row.
        """
        if self.workbook is None:
            raise RuntimeError("Workbook is not opened.")
        ws = self.workbook[sheet_name]
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row_idx < min_row:
                continue
            yield row_idx, row
