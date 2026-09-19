"""Database connection and transaction manager for MySQL."""

import logging
from contextlib import contextmanager
from typing import Generator, Optional
import pymysql
from pymysql.cursors import Cursor

from app.core.config import Config

logger = logging.getLogger("app.core.database")


class DatabaseConnection:
    """Manages MySQL connection lifecycle and transactions using PyMySQL."""

    def __init__(self, config: Config):
        self.config = config
        self._connection: Optional[pymysql.Connection] = None

    def connect(self) -> pymysql.Connection:
        """Establish a connection to MySQL server."""
        if self._connection is None or not self._connection.open:
            try:
                logger.info("Connecting to MySQL at %s", self.config.safe_connection_info())
                self._connection = pymysql.connect(
                    host=self.config.db_host,
                    port=self.config.db_port,
                    user=self.config.db_user,
                    password=self.config.db_password,
                    database=self.config.db_name,
                    charset=self.config.db_charset,
                    autocommit=False,
                )
            except Exception as e:
                logger.error("Failed to connect to MySQL database: %s", e)
                raise ConnectionError(f"Database connection failed: {e}") from e
        return self._connection

    def close(self) -> None:
        """Close the connection."""
        if self._connection is not None:
            try:
                if self._connection.open:
                    self._connection.close()
            except Exception:
                pass
            finally:
                self._connection = None

    @contextmanager
    def transaction(self) -> Generator[Cursor, None, None]:
        """Context manager providing an atomic database transaction.

        Commits on normal exit, rolls back on any error.
        """
        conn = self.connect()
        cursor = conn.cursor()
        try:
            yield cursor
            conn.commit()
            logger.debug("Transaction successfully committed.")
        except Exception as e:
            conn.rollback()
            logger.error("Transaction rolled back due to error: %s", e)
            raise
        finally:
            cursor.close()

    def test_connection(self) -> bool:
        """Test whether database is reachable."""
        try:
            conn = self.connect()
            with conn.cursor() as cur:
                cur.execute("SELECT 1;")
                return cur.fetchone()[0] == 1
        except Exception as e:
            logger.warning("Database connection test failed: %s", e)
            return False
