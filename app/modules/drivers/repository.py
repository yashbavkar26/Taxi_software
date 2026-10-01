"""Repository layer for drivers table operations."""

import logging
from typing import Any, Dict, Iterable, List, Optional, Set
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

    def get_existing_driver_names(
        self, cursor: Cursor, driver_names: Optional[Iterable[str]] = None
    ) -> Set[str]:
        """Fetch set of driver names that already exist in the 'drivers' table."""
        if driver_names is not None:
            names = [str(d).strip() for d in driver_names if d and str(d).strip()]
            if not names:
                return set()
            placeholders = ", ".join(["%s"] * len(names))
            sql = f"SELECT `driver` FROM `drivers` WHERE `driver` IN ({placeholders})"
            cursor.execute(sql, names)
        else:
            sql = "SELECT `driver` FROM `drivers`"
            cursor.execute(sql)

        rows = cursor.fetchall()
        existing: Set[str] = set()
        if rows:
            for row in rows:
                if isinstance(row, (list, tuple)) and len(row) > 0:
                    existing.add(str(row[0]).strip())
                elif isinstance(row, dict) and "driver" in row:
                    existing.add(str(row["driver"]).strip())
                elif isinstance(row, str):
                    existing.add(row.strip())
        return existing

    def insert_new_drivers_from_bookings(
        self,
        cursor: Cursor,
        booking_records: List[Dict[str, Any]],
        known_drivers: Optional[Set[str]] = None,
    ) -> int:
        """
        Check if drivers from bookings already exist in 'drivers' table;
        if not, insert the new drivers with their contact info from bookings.
        """
        if not booking_records:
            return 0

        # Collect unique driver names and their contact info from booking records
        candidates: Dict[str, Optional[str]] = {}
        for record in booking_records:
            driver_val = record.get("driver_name") or record.get("driver")
            if not driver_val:
                continue
            cleaned_name = str(driver_val).strip()
            if not cleaned_name:
                continue
            contact_val = record.get("contact")
            if cleaned_name not in candidates or (not candidates[cleaned_name] and contact_val):
                candidates[cleaned_name] = contact_val

        if not candidates:
            return 0

        # Filter against in-memory known_drivers cache first if provided (case-insensitive)
        known_lower = {d.lower() for d in known_drivers} if known_drivers is not None else set()
        unknown_candidates = {
            name: contact
            for name, contact in candidates.items()
            if name.lower() not in known_lower
        }

        if not unknown_candidates:
            return 0

        # First check if driver already exists in table
        existing_in_db = self.get_existing_driver_names(cursor, unknown_candidates.keys())
        existing_lower = {d.lower() for d in existing_in_db}
        if known_drivers is not None:
            known_drivers.update(existing_in_db)

        # If not, then only insert the driver from the bookings table
        to_insert = []
        seen_insert_lower = set()
        for name, contact in unknown_candidates.items():
            n_low = name.lower()
            if n_low not in existing_lower and n_low not in seen_insert_lower:
                seen_insert_lower.add(n_low)
                to_insert.append({"driver": name, "contact": contact})

        if not to_insert:
            return 0

        columns = ["driver", "contact"]
        cols_str = ", ".join(f"`{c}`" for c in columns)
        placeholders = ", ".join(["%s"] * len(columns))
        if self.duplicate_mode == "upsert":
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE `contact` = VALUES(`contact`)"
        elif self.duplicate_mode == "fail":
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders})"
        else:  # 'skip'
            sql = f"INSERT INTO `drivers` ({cols_str}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE `sr_no` = `sr_no`"

        total_inserted = 0
        for i in range(0, len(to_insert), self.batch_size):
            chunk = to_insert[i : i + self.batch_size]
            params = [[r["driver"], (str(r["contact"])[:20] if r["contact"] else None)] for r in chunk]
            cursor.executemany(sql, params)
            total_inserted += len(chunk)

        if known_drivers is not None:
            for r in to_insert:
                known_drivers.add(r["driver"])

        logger.info("Inserted %d new driver(s) from bookings into 'drivers'", total_inserted)
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

