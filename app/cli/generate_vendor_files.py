"""CLI tool to generate live-syncing vendor Excel workbooks linked to the masterfile."""

import argparse
import sys
from pathlib import Path

from app.modules.vendors.vendor_generator import (
    ACTIVE_VENDORS,
    create_vendor_workbook,
    generate_all_vendor_workbooks,
)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate live-syncing vendor Excel workbooks linked to masterfile_template.xlsx."
    )
    parser.add_argument(
        "--vendor",
        type=str,
        default=None,
        help="Generate a workbook for a specific vendor (e.g., 'Arvind'). If omitted, generates all.",
    )
    parser.add_argument(
        "--masterfile",
        type=str,
        default="masterfile_template.xlsx",
        help="Filename of the masterfile workbook to link against (default: 'masterfile_template.xlsx').",
    )
    parser.add_argument(
        "--out-dir",
        type=str,
        default=".",
        help="Directory where workbooks should be saved (default: current directory / root).",
    )
    parser.add_argument(
        "--template-only",
        action="store_true",
        help="Generate only vendor_template.xlsx without individual vendor files.",
    )

    args = parser.parse_args(argv)
    out_path = Path(args.out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    if args.template_only:
        target = out_path / "vendor_template.xlsx"
        create_vendor_workbook(
            vendor_name=None,
            masterfile_filename=args.masterfile,
            output_path=target,
        )
        print(f"Created template: {target}")
        return 0

    if args.vendor:
        target = out_path / f"{args.vendor}.xlsx"
        create_vendor_workbook(
            vendor_name=args.vendor,
            masterfile_filename=args.masterfile,
            output_path=target,
        )
        print(f"Created vendor workbook: {target}")
        return 0

    # Default: Generate all active vendor files and vendor_template.xlsx
    created = generate_all_vendor_workbooks(
        output_dir=out_path,
        masterfile_filename=args.masterfile,
    )
    print(f"Successfully generated {len(created)} vendor workbooks in '{out_path}':")
    for p in created:
        print(f" - {p.name}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
