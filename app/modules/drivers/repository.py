"""Repository layer for drivers table operations."""

import logging
from typing import Any, Dict, List, Set
from pymysql.cursors import Cursor

logger = logging.getLogger("app.modules.drivers.repository")


class DriverRepository:
    """Database operations for the 'drivers' table."""

    def __init__(self, duplicate_mode: str = "skip", batch_size: int = 1000):
        self.duplicate_mode = duplicate_mode.lower()
        self.batch_size = max(1, batch_size)

    def insert_batch(
        self, cursor: Cursor, driver_records: List[Dict[str, Any]]
    ) -> int:
        """Batch insert driver records into 'drivers' table with duplicate handling."""
        if not driver_records:
            return 0

        columns = ["driver", "contact"]
        cols_str = ", ".join(f"`{c}`" for c in columns)
        placeholders = ", ".join(["%s"] * len(columns))

        if self.duplicate_mode == "skip":
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE `sr_no` = `sr_no`"
        elif self.duplicate_mode == "upsert":
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE `contact` = VALUES(`contact`)"
        else:  # 'fail'
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders})"

        total_inserted = 0
        for i in range(0, len(driver_records), self.batch_size):
            chunk = driver_records[i : i + self.batch_size]
            params = [[record.get(c) for c in columns] for record in chunk]
            cursor.executemany(sql, params)
            total_inserted += len(chunk)

        logger.debug("Inserted/processed %d driver records into 'drivers'", total_inserted)
        return total_inserted

    def ensure_drivers_exist(self, cursor: Cursor, driver_names: Set[str]) -> int:
        """Ensure driver names referenced in bookings exist in 'drivers' to satisfy FK."""
        valid_drivers = [d.strip() for d in driver_names if d and str(d).strip()]
        if not valid_drivers:
            return 0

        records = [{"driver": d, "contact": None} for d in valid_drivers]
        sql = "INSERT INTO `drivers` (`driver`, `contact`) VALUES (%s, %s) ON DUPLICATE KEY UPDATE `sr_no` = `sr_no`"

        for i in range(0, len(records), self.batch_size):
            chunk = records[i : i + self.batch_size]
            params = [[r["driver"], r["contact"]] for r in chunk]
            cursor.executemany(sql, params)

        return len(records)
