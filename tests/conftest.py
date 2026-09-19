"""Pytest fixtures and helpers for creating test workbooks and database mocks."""

from pathlib import Path
from typing import Any, Dict, List, Optional
import openpyxl
import pytest

from app.core.config import Config
from app.core.database import DatabaseConnection


class MockCursor:
    """Mock database cursor tracking executed queries and parameters."""

    def __init__(self, should_fail_on_execute: bool = False):
        self.executed_queries: List[str] = []
        self.executed_params: List[List[Any]] = []
        self.statements: List[tuple] = []  # (query, params_list)
        self.should_fail_on_execute = should_fail_on_execute
        self.is_closed = False

    def execute(self, query: str, params: Optional[Any] = None) -> int:
        if self.should_fail_on_execute:
            raise RuntimeError("Simulated database failure on execute")
        self.executed_queries.append(query)
        param_list = [params] if params is not None else []
        self.executed_params.extend(param_list)
        self.statements.append((query, param_list))
        return 1

    def executemany(self, query: str, param_seq: List[List[Any]]) -> int:
        if self.should_fail_on_execute:
            raise RuntimeError("Simulated database execution failure in executemany")
        self.executed_queries.append(query)
        self.executed_params.extend(param_seq)
        self.statements.append((query, param_seq))
        return len(param_seq)

    def get_params_for_query(self, match_str: str) -> List[List[Any]]:
        """Return aggregated parameters executed for queries matching match_str."""
        params = []
        for query, p_seq in self.statements:
            if match_str in query:
                params.extend(p_seq)
        return params

    def fetchone(self):
        return (1,)

    def fetchall(self):
        return []

    def close(self):
        self.is_closed = True


class MockConnection:
    """Mock database connection tracking transactions (begin, commit, rollback)."""

    def __init__(self, should_fail_on_execute: bool = False):
        self.should_fail_on_execute = should_fail_on_execute
        self.committed = False
        self.rolled_back = False
        self.open = True
        self.cursors: List[MockCursor] = []

    def cursor(self) -> MockCursor:
        cur = MockCursor(should_fail_on_execute=self.should_fail_on_execute)
        self.cursors.append(cur)
        return cur

    def commit(self):
        self.committed = True

    def rollback(self):
        self.rolled_back = True

    def close(self):
        self.open = False


class MockDatabaseConnection(DatabaseConnection):
    """Subclass of DatabaseConnection injecting MockConnection."""

    def __init__(self, should_fail_on_execute: bool = False):
        super().__init__(Config())
        self.mock_conn = MockConnection(should_fail_on_execute=should_fail_on_execute)

    def connect(self):
        return self.mock_conn


@pytest.fixture
def make_workbook(tmp_path: Path):
    """Factory fixture to create temporary .xlsx files with customized sheets and rows."""

    def _create(sheets_data: Dict[str, List[List[Any]]], filename: str = "test.xlsx") -> Path:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for sheet_name, rows in sheets_data.items():
            ws = wb.create_sheet(title=sheet_name)
            for row in rows:
                ws.append(row)

        file_path = tmp_path / filename
        wb.save(file_path)
        wb.close()
        return file_path

    return _create


@pytest.fixture
def sample_booking_headers():
    return [
        "Received",
        "Sl.No.",
        "Name",
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
    ]


@pytest.fixture
def sample_contacts_headers():
    return ["Sr. no", "Driver", "Contact"]
