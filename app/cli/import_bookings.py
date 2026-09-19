"""CLI command to import bookings and driver contacts from Excel into MySQL."""

import argparse
import sys
from pathlib import Path

from app.core.config import Config
from app.core.logging import setup_logging
from app.modules.importers.service import ExcelImportService


def main(argv=None) -> int:
    """Parse CLI arguments and run ExcelImportService."""
    parser = argparse.ArgumentParser(
        description="Robust Excel-to-MySQL import pipeline for Taxi Software."
    )
    parser.add_argument(
        "file_path",
        type=str,
        help="Path to the Excel .xlsx file to import.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate workbook and rows without making any database writes.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=None,
        help="Number of rows per batch insert (default: 1000 or from config).",
    )
    parser.add_argument(
        "--duplicate-mode",
        choices=["skip", "upsert", "fail"],
        default=None,
        help="Duplicate handling strategy for unique constraints (default: skip).",
    )
    parser.add_argument(
        "--no-auto-register-drivers",
        action="store_true",
        help="Do not auto-register unknown driver names referenced in bookings.",
    )
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default=None,
        help="Override logging level (default: from config or INFO).",
    )

    args = parser.parse_args(argv)

    config = Config.from_env()
    if args.log_level:
        config.log_level = args.log_level
    setup_logging(config.log_level)

    if args.batch_size:
        config.batch_size = args.batch_size
    if args.duplicate_mode:
        config.duplicate_mode = args.duplicate_mode
    if args.no_auto_register_drivers:
        config.auto_register_drivers = False

    file_path = Path(args.file_path)
    if not file_path.exists():
        print(f"Error: File not found: '{args.file_path}'", file=sys.stderr)
        return 1

    service = ExcelImportService(config=config)
    result = service.import_workbook(
        file_path=file_path,
        dry_run=args.dry_run,
        batch_size=config.batch_size,
        duplicate_mode=config.duplicate_mode,
    )

    print()
    print(result.summary_text())
    print()

    if result.status in ("SUCCESS", "DRY_RUN") and not (args.dry_run and result.errors):
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
