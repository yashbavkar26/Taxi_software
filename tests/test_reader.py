"""Unit tests for streaming Excel reader."""

from pathlib import Path
import pytest

from app.modules.importers.excel_reader import ExcelWorkbookReader


class TestReader:
    def test_read_workbook_sheets_and_headers(self, make_workbook):
        file_path = make_workbook(
            {
                "Sheet1": [["ColA", "ColB"], [1, 2], [3, 4]],
                "Sheet2": [["X", "Y"], [10, 20]],
            }
        )

        with ExcelWorkbookReader(file_path) as reader:
            assert reader.sheet_names == ["Sheet1", "Sheet2"]
            assert list(reader.read_headers("Sheet1")) == ["ColA", "ColB"]
            rows = list(reader.stream_rows("Sheet1", min_row=2))
            assert len(rows) == 2
            assert rows[0] == (2, (1, 2))
            assert rows[1] == (3, (3, 4))

    def test_detect_empty_sheet(self, make_workbook):
        file_path = make_workbook(
            {
                "DataSheet": [["A", "B"], [1, 2]],
                "EmptySheet": [],
            }
        )

        with ExcelWorkbookReader(file_path) as reader:
            assert not reader.is_sheet_empty("DataSheet")
            assert reader.is_sheet_empty("EmptySheet")

    def test_nonexistent_file_raises_error(self, tmp_path: Path):
        with pytest.raises(ValueError, match="Failed to open Excel workbook"):
            with ExcelWorkbookReader(tmp_path / "non_existent.xlsx") as reader:
                pass
