"""Configuration management for Taxi Software."""

import os
from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv

env_path = Path(".env")
if env_path.exists():
    load_dotenv(dotenv_path=env_path)
else:
    load_dotenv()


@dataclass
class Config:
    """Database and application runtime configuration."""

    db_host: str = "localhost"
    db_port: int = 3306
    db_user: str = "root"
    db_password: str = ""
    db_name: str = "taxi_database"
    db_charset: str = "utf8mb4"
    batch_size: int = 1000
    duplicate_mode: str = "skip"  # 'fail', 'skip', 'upsert'
    reverse_sheets: bool = True
    auto_register_drivers: bool = False
    log_level: str = "INFO"

    @classmethod
    def from_env(cls) -> "Config":
        """Build Config instance from environment variables with defaults."""
        return cls(
            db_host=os.getenv("DB_HOST", "localhost"),
            db_port=int(os.getenv("DB_PORT", "3306")),
            db_user=os.getenv("DB_USER", "root"),
            db_password=os.getenv("DB_PASSWORD", ""),
            db_name=os.getenv("DB_NAME", "taxi_database"),
            db_charset=os.getenv("DB_CHARSET", "utf8mb4"),
            batch_size=int(os.getenv("BATCH_SIZE", "1000")),
            duplicate_mode=os.getenv("DUPLICATE_MODE", "skip").lower(),
            reverse_sheets=os.getenv("REVERSE_SHEETS", "true").lower() in ("true", "1", "yes"),
            auto_register_drivers=os.getenv("AUTO_REGISTER_DRIVERS", "false").lower() in ("true", "1", "yes"),
            log_level=os.getenv("LOG_LEVEL", "INFO").upper(),
        )

    def safe_connection_info(self) -> str:
        """Return safe string representing connection without exposing password."""
        return f"{self.db_user}@{self.db_host}:{self.db_port}/{self.db_name}"
