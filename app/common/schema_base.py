"""Base schema definition class for domain tables."""

from dataclasses import dataclass
from typing import Dict, Set
from app.common.normalizer import ColumnDef


@dataclass
class TableSchema:
    """Represents the expected schema for a destination MySQL table."""

    table_name: str
    primary_key: str
    columns: Dict[str, ColumnDef]
    # Header aliases: maps normalized lowercase header string to db_column
    header_aliases: Dict[str, str]
    # Minimum required headers that must be present in Excel
    required_headers: Set[str]
    # Stored generated columns in MySQL (excluded from INSERT statement)
    generated_columns: Set[str]
