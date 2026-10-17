"""CLI command to synchronize masterfile entries into vendor files."""

import argparse
import sys
from pathlib import Path

from app.modules.vendors.sync_service import sync_all_vendors, watch_and_sync


def main(argv=None) -> int:
    # Ensure stdout/stderr are redirected when running under pythonw.exe
    if sys.stdout is None:
        try:
            sys.stdout = open("vendor_sync.log", "a", encoding="utf-8")
        except Exception:
            pass
    if sys.stderr is None:
        try:
            sys.stderr = open("vendor_sync.log", "a", encoding="utf-8")
        except Exception:
            pass

    parser = argparse.ArgumentParser(
        description="Synchronize masterfile trips into individual vendor Excel files."
    )
    parser.add_argument(
        "--masterfile",
        type=str,
        default=None,
        help="Path to the masterfile Excel workbook (default: auto-detected dynamically).",
    )
    parser.add_argument(
        "--out-dir",
        type=str,
        default=None,
        help="Directory to save vendor workbooks (default: same directory as masterfile).",
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="Run continuously in background, auto-syncing vendor files whenever masterfile is saved in Excel.",
    )

    args = parser.parse_args(argv)

    if args.watch:
        watch_and_sync(
            masterfile_path=args.masterfile,
            output_dir=args.out_dir,
        )
        return 0

    summary = sync_all_vendors(
        masterfile_path=args.masterfile,
        output_dir=args.out_dir,
    )
    print("Sync complete! Summary of synchronized vendor files:")
    for vendor, count in sorted(summary.items()):
        print(f"  • {vendor}.xlsx: {count} trips")

    return 0


if __name__ == "__main__":
    sys.exit(main())
