"""Schema and column mappings for the drivers / contacts table."""

from app.common.normalizer import ColumnDef
from app.common.schema_base import TableSchema

DRIVERS_COLUMNS = {
    "driver": ColumnDef(db_column="driver", data_type="string", is_required=True, is_unique=True, max_length=255),
    "contact": ColumnDef(db_column="contact", data_type="phone", is_required=False, max_length=20),
}

DRIVERS_HEADER_ALIASES = {
    "driver": "driver",
    "driver name": "driver",
    "contact": "contact",
    "mobile": "contact",
    "phone": "contact",
    "sr. no": "sr_no",
    "sr no": "sr_no",
    "sr. no.": "sr_no",
    "sl. no.": "sr_no",
    "sl no": "sr_no",
}

DRIVERS_SCHEMA = TableSchema(
    table_name="drivers",
    primary_key="sr_no",
    columns=DRIVERS_COLUMNS,
    header_aliases=DRIVERS_HEADER_ALIASES,
    required_headers={"driver"},
    generated_columns=set(),
)


def is_contacts_sheet(sheet_name: str) -> bool:
    """Determine if a worksheet represents the drivers/contacts sheet."""
    return sheet_name.strip().lower() in ("contacts", "drivers", "driver contacts", "contact")
