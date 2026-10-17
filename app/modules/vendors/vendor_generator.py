"""Generator for live-linked vendor Excel workbooks and vendor template.

Each vendor file uses native Excel dynamic array FILTER formulas to pull
live data directly from masterfile_template.xlsx without requiring any external
scripts or macros during operation.
"""

from pathlib import Path
from typing import List, Optional, Sequence, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


# Active vendors registered in Taxi Software system
ACTIVE_VENDORS: List[str] = [
    "Arvind",
    "Bharat",
    "Dilip",
    "Gajanan",
    "Harshal",
    "Kunal",
    "Lalu",
    "Manoj",
    "Maulla",
    "Pritesh",
    "Satesh",
    "Shubham",
    "Subodh",
    "Yogesh",
]

# Masterfile table headers
MASTERFILE_HEADERS: List[str] = [
    "Sl. No",           # A (1)
    "Vendor",           # B (2)
    "Date",             # C (3)
    "Vech.no",          # D (4)
    "Vech Type",        # E (5)
    "Starting Km",      # F (6)
    "Ending Km",        # G (7)
    "Total Km",         # H (8)
    "Night",            # I (9)
    "Rate",             # J (10)
    "Toll",             # K (11)
    "Amt",              # L (12)
    "Guest",            # M (13)
    "Total Amt",        # N (14)
    "Payment Details",  # O (15)
    "In Time",          # P (16)
    "Out Time",         # Q (17)
    "Driver Name",      # R (18)
    "Places",           # S (19)
    "Payment Date",     # T (20)
]

COL_WIDTHS = {
    "A": 9,   # Sl. No
    "B": 16,  # Vendor
    "C": 14,  # Date
    "D": 18,  # Vech.no
    "E": 16,  # Vech Type
    "F": 15,  # Starting Km
    "G": 15,  # Ending Km
    "H": 13,  # Total Km
    "I": 12,  # Night
    "J": 10,  # Rate
    "K": 12,  # Toll
    "L": 14,  # Amt
    "M": 22,  # Guest
    "N": 16,  # Total Amt
    "O": 16,  # Payment Details
    "P": 13,  # In Time
    "Q": 13,  # Out Time
    "R": 20,  # Driver Name
    "S": 22,  # Places
    "T": 15,  # Payment Date
}


def create_vendor_workbook(
    vendor_name: Optional[str] = None,
    masterfile_filename: str = "masterfile_template.xlsx",
    output_path: Optional[Union[str, Path]] = None,
) -> Path:
    """Create a standardized Excel workbook dynamically linked to masterfile_template.xlsx.

    Args:
        vendor_name: Name of the vendor (e.g. 'Arvind'). If None, creates a blank template.
        masterfile_filename: Name of the masterfile workbook to link against.
        output_path: File path to save to. Defaults to '{vendor_name}.xlsx' or 'vendor_template.xlsx'.
    """
    if output_path is None:
        filename = f"{vendor_name}.xlsx" if vendor_name else "vendor_template.xlsx"
        output_path = Path(filename)
    else:
        output_path = Path(output_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendor_Trips"
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A7"

    font_family = "Trebuchet MS"

    # Styling definitions
    title_font = Font(name=font_family, size=13, bold=True, color="FF1A73E8")
    label_font = Font(name=font_family, size=10, bold=True, color="FF3C4043")
    vendor_font = Font(name=font_family, size=11, bold=True, color="FF137333")
    kpi_title_font = Font(name=font_family, size=9, bold=True, color="FF5F6368")
    kpi_num_font = Font(name=font_family, size=13, bold=True, color="FF202124")
    tbl_header_font = Font(name=font_family, size=10.5, bold=True, color="FFFFFFFF")

    title_fill = PatternFill(start_color="FFE8F0FE", end_color="FFE8F0FE", fill_type="solid")
    kpi_hdr_fill = PatternFill(start_color="FFF1F3F4", end_color="FFF1F3F4", fill_type="solid")
    kpi_card_fill = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
    vendor_box_fill = PatternFill(start_color="FFE6F4EA", end_color="FFE6F4EA", fill_type="solid")
    table_hdr_fill = PatternFill(start_color="FF1A73E8", end_color="FF1A73E8", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="FFD3D3D3")
    medium_border_side = Side(border_style="medium", color="FF1A73E8")
    box_border = Border(
        left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side
    )
    tbl_header_border = Border(
        left=thin_border_side, right=thin_border_side, top=medium_border_side, bottom=medium_border_side
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Row 1: Title Banner
    ws.merge_cells("A1:T1")
    title_cell = ws.cell(1, 1, "VENDOR TRIP LOG & BILLING SUMMARY")
    title_cell.font = title_font
    title_cell.alignment = align_center
    title_cell.fill = title_fill
    ws.row_dimensions[1].height = 32

    # Row 2: Vendor Selector and Source Info
    ws.row_dimensions[2].height = 26
    c_vl = ws.cell(2, 1, "Active Vendor:")
    c_vl.font = label_font
    c_vl.alignment = align_right

    c_vval = ws.cell(2, 2, vendor_name if vendor_name else "")
    c_vval.font = vendor_font
    c_vval.alignment = align_center
    c_vval.fill = vendor_box_fill
    c_vval.border = box_border

    c_sl = ws.cell(2, 3, "Master Source:")
    c_sl.font = label_font
    c_sl.alignment = align_right

    c_sval = ws.cell(2, 4, masterfile_filename)
    c_sval.font = Font(name=font_family, size=10, italic=True, color="FF5F6368")
    c_sval.alignment = align_left

    c_note = ws.cell(2, 6, "Note: Real-time dynamic mirror of masterfile data via Excel array formula.")
    c_note.font = Font(name=font_family, size=9.5, italic=True, color="FF70757A")
    c_note.alignment = align_left

    # Row 3 & 4: Executive Summary KPI Cards
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 26

    kpi_configs = [
        (
            2,
            "TOTAL TRIPS",
            f'=IF(OR(ISBLANK($B$2),$B$2=""),0,COUNTIF(\'[{masterfile_filename}]Masterfile\'!$B$2:$B$5000, $B$2))',
            "#,##0",
        ),
        (
            4,
            "TOTAL DISTANCE (KM)",
            f'=IF(OR(ISBLANK($B$2),$B$2=""),0,SUMIFS(\'[{masterfile_filename}]Masterfile\'!$H$2:$H$5000, \'[{masterfile_filename}]Masterfile\'!$B$2:$B$5000, $B$2))',
            "#,##0.00",
        ),
        (
            6,
            "TOTAL AMOUNT (INR)",
            f'=IF(OR(ISBLANK($B$2),$B$2=""),0,SUMIFS(\'[{masterfile_filename}]Masterfile\'!$N$2:$N$5000, \'[{masterfile_filename}]Masterfile\'!$B$2:$B$5000, $B$2))',
            "₹#,##0.00",
        ),
        (
            8,
            "PAYMENT PENDING",
            f'=IF(OR(ISBLANK($B$2),$B$2=""),0,COUNTIFS(\'[{masterfile_filename}]Masterfile\'!$B$2:$B$5000, $B$2, \'[{masterfile_filename}]Masterfile\'!$O$2:$O$5000, FALSE))',
            "#,##0",
        ),
    ]

    for col_idx, label, formula, num_fmt in kpi_configs:
        c_title = ws.cell(3, col_idx, label)
        c_title.font = kpi_title_font
        c_title.fill = kpi_hdr_fill
        c_title.alignment = align_center
        c_title.border = box_border

        c_val = ws.cell(4, col_idx, formula)
        c_val.font = kpi_num_font
        c_val.fill = kpi_card_fill
        c_val.alignment = align_center
        c_val.number_format = num_fmt
        c_val.border = box_border

    # Row 5: Spacer row
    ws.row_dimensions[5].height = 10

    # Row 6: Masterfile Table Headers
    ws.row_dimensions[6].height = 28
    for c_idx, header_text in enumerate(MASTERFILE_HEADERS, 1):
        cell = ws.cell(6, c_idx, header_text)
        cell.font = tbl_header_font
        cell.fill = table_hdr_fill
        cell.alignment = align_center
        cell.border = tbl_header_border

    # Set Column Widths
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Row 7: Spilled Dynamic Array FILTER Formula
    ws.row_dimensions[7].height = 22
    spill_formula = (
        f'=IF(OR(ISBLANK($B$2),$B$2=""),"Please select a vendor in cell B2",'
        f"_xlfn.FILTER('[{masterfile_filename}]Masterfile'!$A$2:$T$5000,"
        f"'[{masterfile_filename}]Masterfile'!$B$2:$B$5000=$B$2,"
        f'"No trips recorded for "&$B$2))'
    )
    c_spill = ws.cell(7, 1, spill_formula)
    c_spill.font = Font(name=font_family, size=10, color="FF202124")
    c_spill.alignment = align_center

    # Pre-configure number formatting on data columns down to row 200
    for r in range(7, 201):
        ws.cell(r, 1).alignment = align_center  # Sl. No
        ws.cell(r, 2).alignment = align_center  # Vendor
        ws.cell(r, 3).alignment = align_center  # Date
        ws.cell(r, 3).number_format = "DD/MM/YYYY"
        ws.cell(r, 4).alignment = align_center  # Vech.no
        ws.cell(r, 5).alignment = align_center  # Vech Type
        ws.cell(r, 6).alignment = align_right   # Starting Km
        ws.cell(r, 6).number_format = "#,##0.00"
        ws.cell(r, 7).alignment = align_right   # Ending Km
        ws.cell(r, 7).number_format = "#,##0.00"
        ws.cell(r, 8).alignment = align_right   # Total Km
        ws.cell(r, 8).number_format = "#,##0.00"
        ws.cell(r, 9).alignment = align_right   # Night
        ws.cell(r, 9).number_format = "#,##0.00"
        ws.cell(r, 10).alignment = align_right  # Rate
        ws.cell(r, 10).number_format = "#,##0.00"
        ws.cell(r, 11).alignment = align_right  # Toll
        ws.cell(r, 11).number_format = "#,##0.00"
        ws.cell(r, 12).alignment = align_right  # Amt
        ws.cell(r, 12).number_format = "#,##0.00"
        ws.cell(r, 13).alignment = align_left   # Guest
        ws.cell(r, 14).alignment = align_right  # Total Amt
        ws.cell(r, 14).number_format = "#,##0.00"
        ws.cell(r, 15).alignment = align_center # Payment Details (Bool)
        ws.cell(r, 16).alignment = align_center # In Time
        ws.cell(r, 17).alignment = align_center # Out Time
        ws.cell(r, 18).alignment = align_left   # Driver Name
        ws.cell(r, 19).alignment = align_left   # Places
        ws.cell(r, 20).alignment = align_center # Payment Date
        ws.cell(r, 20).number_format = "DD/MM/YYYY"

    # Add Dropdown Validation for Cell B2 (Active Vendor)
    vendor_list_str = f'"{", ".join(ACTIVE_VENDORS)}"'
    dv = DataValidation(
        type="list",
        formula1=vendor_list_str,
        allow_blank=True,
    )
    dv.error = "Select a valid vendor from the registered active vendors list."
    dv.errorTitle = "Invalid Vendor"
    dv.prompt = "Choose a vendor to display all live trips and billing totals."
    dv.promptTitle = "Select Vendor"
    ws.add_data_validation(dv)
    dv.add("B2")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


def generate_all_vendor_workbooks(
    output_dir: Union[str, Path] = ".",
    masterfile_filename: str = "masterfile_template.xlsx",
) -> List[Path]:
    """Generate vendor_template.xlsx and linked workbooks for all registered active vendors.

    Args:
        output_dir: Directory where workbooks should be saved.
        masterfile_filename: Masterfile workbook filename used for cross-workbook links.

    Returns:
        List of created Path objects.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    created: List[Path] = []

    # 1. Standardized Reusable Template
    template_path = out / "vendor_template.xlsx"
    create_vendor_workbook(
        vendor_name=None,
        masterfile_filename=masterfile_filename,
        output_path=template_path,
    )
    created.append(template_path)

    # 2. Individual Linked Workbooks for Active Vendors
    for vendor in ACTIVE_VENDORS:
        v_path = out / f"{vendor}.xlsx"
        create_vendor_workbook(
            vendor_name=vendor,
            masterfile_filename=masterfile_filename,
            output_path=v_path,
        )
        created.append(v_path)

    return created


if __name__ == "__main__":
    generated = generate_all_vendor_workbooks()
    print(f"Generated {len(generated)} vendor workbooks successfully.")
