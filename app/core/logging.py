"""Centralized logging configuration for Taxi Software."""

import logging


def setup_logging(level_name: str = "INFO") -> None:
    """Configure console logging format and log level."""
    numeric_level = getattr(logging, level_name.upper(), logging.INFO)
    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
