"""Repository layer for bookings table operations."""

import logging
from typing import Any, Dict, List
from pymysql.cursors import Cursor

logger = logging.getLogger("app.modules.bookings.repository")


class BookingRepository:
    """Database operations for the 'bookings' table."""

    def __init__(self, duplicate_mode: str = "skip", batch_size: int = 1000):
        self.duplicate_mode = duplicate_mode.lower()
        self.batch_size = max(1, batch_size)

    def insert_batch(
        self, cursor: Cursor, booking_records: List[Dict[str, Any]]
    ) -> int:
        """Batch insert booking records into 'bookings' table using parameterized executemany."""
        if not booking_records:
            return 0

        columns = [
            "received",
            "sl_no",
            "coordinator",
            "guest_name",
            "guest_mobile_no",
            "travel_date",
            "pickup_location",
            "drop_location",
            "start_time",
            "return_time",
            "reason",
            "driver_name",
            "contact",
            "vehicle_type",
            "company",
        ]

        cols_str = ", ".join(f"`{c}`" for c in columns)
        placeholders = ", ".join(["%s"] * len(columns))

        if self.duplicate_mode == "skip":
            sql = f"INSERT IGNORE INTO `bookings` ({cols_str}) VALUES ({placeholders})"
        else:
            sql = f"INSERT INTO `bookings` ({cols_str}) VALUES ({placeholders})"

        def _get_val(record: Dict[str, Any], col: str) -> Any:
            if col == "company":
                v = record.get("company")
                if v is None:
                    v = record.get("Company")
                if v is None:
                    v = record.get("vendor")
                if v is None:
                    v = record.get("Vendor")
                return v if (v is not None and str(v).strip() != "") else "GIM"
            if col == "received":
                v = record.get("received")
                if v is None:
                    v = record.get("Received")
                return bool(v) if v is not None else False
            return record.get(col)

        total_inserted = 0
        for i in range(0, len(booking_records), self.batch_size):
            chunk = booking_records[i : i + self.batch_size]
            params = [[_get_val(record, c) for c in columns] for record in chunk]
            cursor.executemany(sql, params)
            total_inserted += len(chunk)


        logger.debug("Inserted %d records into 'bookings'", total_inserted)
        return total_inserted
