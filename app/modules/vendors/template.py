"""Excel Masterfile template generator matching database schema and Masterfile conventions."""

import os
from pathlib import Path
from typing import List, Sequence, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


def generate_masterfile_template(
    output_path: Union[str, Path] = "masterfile_template.xlsx",
    additional_paths: Sequence[Union[str, Path]] = (),
) -> Path:
    """Create a standardized Excel workbook template for masterfile / vendor entries.

    Features:
    - Matches MySQL masterfile schema in taxi_schema.sql (sl_no, vendor, travel_date,
      vehicle_no, vehicle_type, starting_km, ending_km, total_km, night, rate, toll,
      amt, guest_name, total_amt, payment_details, in_time, out_time, driver_name,
      places, payment_date).
    - Faithful to Requirements/Input/Masterfile.xlsx structure while enhancing layout,
      validations, lookups, and formula automation.
    - Automated formulas for Total Km (ending_km - starting_km), 80-km package Amt calculation,
      and Total Amt (amt + toll + night).
    - Multi-sheet workbook:
      1. Masterfile: Active sheet pre-populated with realistic sample rows and ready rows.
      2. Sample: Clean blank template sheet ready to duplicate for new billing periods.
      3. Vendors: Complete directory of active vendors with contact and vehicle info.
      4. Drivers: Directory of active drivers and contacts.
      5. Vehicles: Fleet vehicle registry (vehicle number, type, assigned vendor).
      6. Rate_Card: Standard 80-km package pricing rules and excess km rates.
      7. Annexure_Sample: Vendor monthly annexure view matching view_annexure_arvind.
    - Data validation dropdowns for Vendor, Vehicle Type, Rate, Payment Details, and Driver.
    - Professional executive styling (Segoe UI, Soft Slate Blue palette, distinct formula accents).
    """
    wb = openpyxl.Workbook()

    # Font & Color Palette
    font_family = "Segoe UI"
    primary_header_fill = PatternFill(start_color="FFEAEFF5", end_color="FFEAEFF5", fill_type="solid")  # Executive Slate Blue
    formula_header_fill = PatternFill(start_color="FFDCE7F5", end_color="FFDCE7F5", fill_type="solid")  # Soft Ice Blue for formulas
    reference_header_fill = PatternFill(start_color="FFF0F4F8", end_color="FFF0F4F8", fill_type="solid")

    header_font = Font(name=font_family, size=11, bold=True, color="FF1E293B")
    formula_header_font = Font(name=font_family, size=11, bold=True, color="FF1D4ED8")  # Highlighted blue
    data_font = Font(name=font_family, size=10, bold=False, color="FF0F172A")
    formula_data_font = Font(name=font_family, size=10, bold=False, color="FF1E40AF")
    bold_data_font = Font(name=font_family, size=10, bold=True, color="FF0F172A")

    thin_border_side = Side(border_style="thin", color="FFCBD5E1")
    header_border_side = Side(border_style="thin", color="FF94A3B8")

    header_border = Border(left=header_border_side, right=header_border_side, top=header_border_side, bottom=header_border_side)
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # Masterfile Headers matching taxi_schema.sql and Masterfile.xlsx
    headers = [
        "Sl. No",           # A (1) - sl_no
        "Vendor",           # B (2) - vendor
        "Date",             # C (3) - travel_date
        "Vech.no",          # D (4) - vehicle_no
        "Vech Type",        # E (5) - vehicle_type
        "Starting Km",      # F (6) - starting_km
        "Ending Km",        # G (7) - ending_km
        "Total Km",         # H (8) - total_km (STORED GENERATED)
        "Night",            # I (9) - night
        "Rate",             # J (10) - rate
        "Toll",             # K (11) - toll
        "Amt",              # L (12) - amt
        "Guest",            # M (13) - guest_name
        "Total Amt",        # N (14) - total_amt (STORED GENERATED)
        "Payment Details",  # O (15) - payment_details
        "In Time",          # P (16) - in_time
        "Out Time",         # Q (17) - out_time
        "Driver Name",      # R (18) - driver_name
        "Places",           # S (19) - places
        "Payment Date",     # T (20) - payment_date
    ]

    # Formula columns: H (8, Total Km), L (12, Amt), N (14, Total Amt)
    formula_col_indices = {8, 12, 14}

    col_widths = {
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
        "R": 18,  # Driver Name
        "S": 28,  # Places
        "T": 15,  # Payment Date
    }

    # Realistic sample records directly synthesized from Masterfile.xlsx
    sample_records = [
        {
            "sl_no": 1,
            "vendor": "Arvind",
            "date": "01/12/2025",
            "vech_no": "GA 07 T 2608",
            "vech_type": "Baleno",
            "start_km": 59989,
            "end_km": 60058,
            "night": None,
            "rate": 22,
            "toll": None,
            "amt": None,  # Calculated via formula
            "guest": "BDA PL",
            "payment_details": False,
            "in_time": "8:15 AM",
            "out_time": "10:30 AM",
            "driver": "Abdul",
            "places": "Local BDA assignment",
            "payment_date": None,
        },
        {
            "sl_no": 2,
            "vendor": "Arvind",
            "date": "01/12/2025",
            "vech_no": "GA 07 T 2608",
            "vech_type": "Baleno",
            "start_km": 60030,
            "end_km": 60140,
            "night": 100,
            "rate": 22,
            "toll": None,
            "amt": None,
            "guest": "GIM",
            "payment_details": False,
            "in_time": "4:30 PM",
            "out_time": "7:30 PM",
            "driver": "Abdul",
            "places": "Campus to Panjim & return",
            "payment_date": None,
        },
        {
            "sl_no": 3,
            "vendor": "Arvind",
            "date": "01/12/2025",
            "vech_no": "GA 07  8151",
            "vech_type": "Swift Dzire",
            "start_km": 250355,
            "end_km": 250582,
            "night": 400,
            "rate": 22,
            "toll": None,
            "amt": None,
            "guest": "GIM",
            "payment_details": False,
            "in_time": "1:30 PM",
            "out_time": "11:00 PM",
            "driver": "Suyog",
            "places": "South Goa official duty",
            "payment_date": None,
        },
        {
            "sl_no": 4,
            "vendor": "Arvind",
            "date": "01/12/2025",
            "vech_no": "GA 07 F 8056",
            "vech_type": "Swift Dzire",
            "start_km": 230280,
            "end_km": 230390,
            "night": 100,
            "rate": 22,
            "toll": 50,
            "amt": None,
            "guest": "GIM",
            "payment_details": False,
            "in_time": "5:00 AM",
            "out_time": "9:00 AM",
            "driver": "Rahul",
            "places": "Mopa Airport pickup",
            "payment_date": None,
        },
        {
            "sl_no": 5,
            "vendor": "Gajanan",
            "date": "01/12/2025",
            "vech_no": "GA 06 T 4317",
            "vech_type": "Swift Dzire",
            "start_km": 333064,
            "end_km": 333174,
            "night": 200,
            "rate": None,
            "toll": None,
            "amt": 1600,  # Fixed package rate
            "guest": "GIM",
            "payment_details": False,
            "in_time": "8:00 PM",
            "out_time": "11:30 PM",
            "driver": "Gajanan",
            "places": "GIM to Dabolim Airport",
            "payment_date": None,
        },
        {
            "sl_no": 6,
            "vendor": "Harshal",
            "date": "01/12/2025",
            "vech_no": "GA 11 T 4914",
            "vech_type": "Swift Dzire",
            "start_km": 117630,
            "end_km": 117716,
            "night": 200,
            "rate": None,
            "toll": None,
            "amt": 1400,  # Fixed airport package
            "guest": "GIM",
            "payment_details": True,
            "in_time": "8:40 PM",
            "out_time": "11:30 PM",
            "driver": "Harshal",
            "places": "Mopa Airport to GIM",
            "payment_date": "25/02/2026",
        },
        {
            "sl_no": 7,
            "vendor": "Lalu",
            "date": "01/12/2025",
            "vech_no": "GA 03 V 6153",
            "vech_type": "Swift Dzire",
            "start_km": 39926,
            "end_km": 39997,
            "night": None,
            "rate": 18,
            "toll": None,
            "amt": None,
            "guest": "GIM",
            "payment_details": True,
            "in_time": "7:00 PM",
            "out_time": "10:00 PM",
            "driver": "Malik",
            "places": "Faculty local disposal",
            "payment_date": "15/01/2026",
        },
        {
            "sl_no": 8,
            "vendor": "Bharat",
            "date": "02/12/2025",
            "vech_no": "GA 07 F 4403",
            "vech_type": "Ertiga",
            "start_km": 84210,
            "end_km": 84325,
            "night": 150,
            "rate": 24,
            "toll": 100,
            "amt": None,
            "guest": "Dr. Roy",
            "payment_details": False,
            "in_time": "9:30 AM",
            "out_time": "6:00 PM",
            "driver": "Bharat",
            "places": "North Goa site visit",
            "payment_date": None,
        },
        {
            "sl_no": 9,
            "vendor": "Dilip",
            "date": "03/12/2025",
            "vech_no": "GA 06 T 4967",
            "vech_type": "Swift Dzire",
            "start_km": 142100,
            "end_km": 142215,
            "night": None,
            "rate": 20,
            "toll": None,
            "amt": None,
            "guest": "Dr. Divya Singhal",
            "payment_details": False,
            "in_time": "10:00 AM",
            "out_time": "2:30 PM",
            "driver": "Lokesh",
            "places": "GIM to Old Goa",
            "payment_date": None,
        },
        {
            "sl_no": 10,
            "vendor": "Yogesh",
            "date": "04/12/2025",
            "vech_no": "GA 05 T 8121",
            "vech_type": "Ertiga",
            "start_km": 95120,
            "end_km": 95260,
            "night": 300,
            "rate": 25,
            "toll": 50,
            "amt": None,
            "guest": "MBA Executive Batch",
            "payment_details": False,
            "in_time": "8:00 AM",
            "out_time": "8:30 PM",
            "driver": "Yogesh",
            "places": "Corporate field project",
            "payment_date": None,
        },
    ]

    def build_amt_formula(row_idx: int) -> str:
        """Construct the 80km package formula or keep manual entry."""
        h = f"H{row_idx}"
        j = f"J{row_idx}"
        return (
            f'=IF(OR(ISBLANK({h}),ISBLANK({j}),{j}=""),"",'
            f'IF({j}=35,IF({h}<=80,3500,3500+(({h}-80)*25)),'
            f'IF({h}<=80,{j}*100,{j}*100+(({h}-80)*{j}))))'
        )

    # ----------------------------------------------------
    # Sheet 1: Masterfile (Active Entry Sheet with Samples)
    # ----------------------------------------------------
    ws_master = wb.active
    ws_master.title = "Masterfile"
    ws_master.views.sheetView[0].showGridLines = True
    ws_master.freeze_panes = "A2"
    ws_master.row_dimensions[1].height = 28

    for col_idx, h in enumerate(headers, 1):
        cell = ws_master.cell(1, col_idx, h)
        cell.font = formula_header_font if col_idx in formula_col_indices else header_font
        cell.fill = formula_header_fill if col_idx in formula_col_indices else primary_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for col_letter, width in col_widths.items():
        ws_master.column_dimensions[col_letter].width = width

    total_master_rows = 50
    for r in range(2, total_master_rows + 2):
        ws_master.row_dimensions[r].height = 22
        is_sample = (r - 2) < len(sample_records)
        rec = sample_records[r - 2] if is_sample else None

        # A: Sl. No
        c_a = ws_master.cell(r, 1, rec["sl_no"] if is_sample else (r - 1))
        c_a.alignment = align_center
        c_a.font = data_font
        c_a.border = cell_border

        # B: Vendor
        c_b = ws_master.cell(r, 2, rec["vendor"] if is_sample else None)
        c_b.alignment = align_left
        c_b.font = bold_data_font if is_sample else data_font
        c_b.border = cell_border

        # C: Date
        c_c = ws_master.cell(r, 3, rec["date"] if is_sample else None)
        c_c.alignment = align_center
        c_c.font = data_font
        c_c.number_format = "DD/MM/YYYY"
        c_c.border = cell_border

        # D: Vech.no
        c_d = ws_master.cell(r, 4, rec["vech_no"] if is_sample else None)
        c_d.alignment = align_center
        c_d.font = data_font
        c_d.border = cell_border

        # E: Vech Type
        c_e = ws_master.cell(r, 5, rec["vech_type"] if is_sample else None)
        c_e.alignment = align_center
        c_e.font = data_font
        c_e.border = cell_border

        # F: Starting Km
        c_f = ws_master.cell(r, 6, rec["start_km"] if is_sample else None)
        c_f.alignment = align_right
        c_f.font = data_font
        c_f.number_format = "#,##0.00"
        c_f.border = cell_border

        # G: Ending Km
        c_g = ws_master.cell(r, 7, rec["end_km"] if is_sample else None)
        c_g.alignment = align_right
        c_g.font = data_font
        c_g.number_format = "#,##0.00"
        c_g.border = cell_border

        # H: Total Km (Formula)
        total_km_formula = f'=IF(OR(ISBLANK(G{r}),ISBLANK(F{r})),"",G{r}-F{r})'
        c_h = ws_master.cell(r, 8, total_km_formula)
        c_h.alignment = align_right
        c_h.font = formula_data_font
        c_h.number_format = "#,##0.00"
        c_h.border = cell_border

        # I: Night
        c_i = ws_master.cell(r, 9, rec["night"] if is_sample else None)
        c_i.alignment = align_right
        c_i.font = data_font
        c_i.number_format = "#,##0.00"
        c_i.border = cell_border

        # J: Rate
        c_j = ws_master.cell(r, 10, rec["rate"] if is_sample else None)
        c_j.alignment = align_center
        c_j.font = data_font
        c_j.number_format = "0"
        c_j.border = cell_border

        # K: Toll
        c_k = ws_master.cell(r, 11, rec["toll"] if is_sample else None)
        c_k.alignment = align_right
        c_k.font = data_font
        c_k.number_format = "#,##0.00"
        c_k.border = cell_border

        # L: Amt (Formula for package or fixed amount)
        if is_sample and rec["amt"] is not None:
            amt_val = rec["amt"]
        else:
            amt_val = build_amt_formula(r)
        c_l = ws_master.cell(r, 12, amt_val)
        c_l.alignment = align_right
        c_l.font = formula_data_font if str(amt_val).startswith("=") else data_font
        c_l.number_format = "#,##0.00"
        c_l.border = cell_border

        # M: Guest
        c_m = ws_master.cell(r, 13, rec["guest"] if is_sample else None)
        c_m.alignment = align_left
        c_m.font = data_font
        c_m.border = cell_border

        # N: Total Amt (Formula: Amt + Toll + Night)
        total_amt_formula = f'=IF(AND(ISBLANK(L{r}),ISBLANK(K{r}),ISBLANK(I{r})),"",SUM(L{r},K{r},I{r}))'
        c_n = ws_master.cell(r, 14, total_amt_formula)
        c_n.alignment = align_right
        c_n.font = formula_data_font
        c_n.number_format = "#,##0.00"
        c_n.border = cell_border

        # O: Payment Details (Boolean)
        c_o = ws_master.cell(r, 15, rec["payment_details"] if is_sample else False)
        c_o.alignment = align_center
        c_o.font = data_font
        c_o.border = cell_border

        # P: In Time
        c_p = ws_master.cell(r, 16, rec["in_time"] if is_sample else None)
        c_p.alignment = align_center
        c_p.font = data_font
        c_p.border = cell_border

        # Q: Out Time
        c_q = ws_master.cell(r, 17, rec["out_time"] if is_sample else None)
        c_q.alignment = align_center
        c_q.font = data_font
        c_q.border = cell_border

        # R: Driver Name
        c_r = ws_master.cell(r, 18, rec["driver"] if is_sample else None)
        c_r.alignment = align_left
        c_r.font = data_font
        c_r.border = cell_border

        # S: Places
        c_s = ws_master.cell(r, 19, rec["places"] if is_sample else None)
        c_s.alignment = align_left
        c_s.font = data_font
        c_s.border = cell_border

        # T: Payment Date
        c_t = ws_master.cell(r, 20, rec["payment_date"] if is_sample else None)
        c_t.alignment = align_center
        c_t.font = data_font
        c_t.number_format = "DD/MM/YYYY"
        c_t.border = cell_border

    # Data Validations on Masterfile
    dv_vendor = DataValidation(type="list", formula1="=Vendors!$B$2:$B$50", allow_blank=True)
    ws_master.add_data_validation(dv_vendor)
    dv_vendor.add("B2:B200")

    dv_vech = DataValidation(type="list", formula1='"Baleno,Swift Dzire,Ertiga,Innova,Crysta,Wagon R,Kia,Sedan,Tempo"', allow_blank=True)
    ws_master.add_data_validation(dv_vech)
    dv_vech.add("E2:E200")

    dv_rate = DataValidation(type="list", formula1='"18,20,22,23,24,25,28,35"', allow_blank=True)
    ws_master.add_data_validation(dv_rate)
    dv_rate.add("J2:J200")

    dv_pay = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_master.add_data_validation(dv_pay)
    dv_pay.add("O2:O200")

    dv_driver = DataValidation(type="list", formula1="=Drivers!$B$2:$B$50", allow_blank=True)
    ws_master.add_data_validation(dv_driver)
    dv_driver.add("R2:R200")

    # ----------------------------------------------------
    # Sheet 2: Sample (Clean Blank Template to Duplicate)
    # ----------------------------------------------------
    ws_sample = wb.create_sheet("Sample")
    ws_sample.views.sheetView[0].showGridLines = True
    ws_sample.freeze_panes = "A2"
    ws_sample.row_dimensions[1].height = 28

    for col_idx, h in enumerate(headers, 1):
        cell = ws_sample.cell(1, col_idx, h)
        cell.font = formula_header_font if col_idx in formula_col_indices else header_font
        cell.fill = formula_header_fill if col_idx in formula_col_indices else primary_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for col_letter, width in col_widths.items():
        ws_sample.column_dimensions[col_letter].width = width

    for r in range(2, 52):
        ws_sample.row_dimensions[r].height = 22

        # A: Sl. No
        c_a = ws_sample.cell(r, 1, r - 1)
        c_a.alignment = align_center
        c_a.font = data_font
        c_a.border = cell_border

        # B: Vendor
        c_b = ws_sample.cell(r, 2, None)
        c_b.alignment = align_left
        c_b.font = data_font
        c_b.border = cell_border

        # C: Date
        c_c = ws_sample.cell(r, 3, None)
        c_c.alignment = align_center
        c_c.font = data_font
        c_c.number_format = "DD/MM/YYYY"
        c_c.border = cell_border

        # D: Vech.no
        c_d = ws_sample.cell(r, 4, None)
        c_d.alignment = align_center
        c_d.font = data_font
        c_d.border = cell_border

        # E: Vech Type
        c_e = ws_sample.cell(r, 5, None)
        c_e.alignment = align_center
        c_e.font = data_font
        c_e.border = cell_border

        # F: Starting Km
        c_f = ws_sample.cell(r, 6, None)
        c_f.alignment = align_right
        c_f.font = data_font
        c_f.number_format = "#,##0.00"
        c_f.border = cell_border

        # G: Ending Km
        c_g = ws_sample.cell(r, 7, None)
        c_g.alignment = align_right
        c_g.font = data_font
        c_g.number_format = "#,##0.00"
        c_g.border = cell_border

        # H: Total Km (Formula)
        total_km_formula = f'=IF(OR(ISBLANK(G{r}),ISBLANK(F{r})),"",G{r}-F{r})'
        c_h = ws_sample.cell(r, 8, total_km_formula)
        c_h.alignment = align_right
        c_h.font = formula_data_font
        c_h.number_format = "#,##0.00"
        c_h.border = cell_border

        # I: Night
        c_i = ws_sample.cell(r, 9, None)
        c_i.alignment = align_right
        c_i.font = data_font
        c_i.number_format = "#,##0.00"
        c_i.border = cell_border

        # J: Rate
        c_j = ws_sample.cell(r, 10, None)
        c_j.alignment = align_center
        c_j.font = data_font
        c_j.number_format = "0"
        c_j.border = cell_border

        # K: Toll
        c_k = ws_sample.cell(r, 11, None)
        c_k.alignment = align_right
        c_k.font = data_font
        c_k.number_format = "#,##0.00"
        c_k.border = cell_border

        # L: Amt (Formula)
        c_l = ws_sample.cell(r, 12, build_amt_formula(r))
        c_l.alignment = align_right
        c_l.font = formula_data_font
        c_l.number_format = "#,##0.00"
        c_l.border = cell_border

        # M: Guest
        c_m = ws_sample.cell(r, 13, None)
        c_m.alignment = align_left
        c_m.font = data_font
        c_m.border = cell_border

        # N: Total Amt (Formula)
        total_amt_formula = f'=IF(AND(ISBLANK(L{r}),ISBLANK(K{r}),ISBLANK(I{r})),"",SUM(L{r},K{r},I{r}))'
        c_n = ws_sample.cell(r, 14, total_amt_formula)
        c_n.alignment = align_right
        c_n.font = formula_data_font
        c_n.number_format = "#,##0.00"
        c_n.border = cell_border

        # O: Payment Details
        c_o = ws_sample.cell(r, 15, False)
        c_o.alignment = align_center
        c_o.font = data_font
        c_o.border = cell_border

        # P: In Time
        c_p = ws_sample.cell(r, 16, None)
        c_p.alignment = align_center
        c_p.font = data_font
        c_p.border = cell_border

        # Q: Out Time
        c_q = ws_sample.cell(r, 17, None)
        c_q.alignment = align_center
        c_q.font = data_font
        c_q.border = cell_border

        # R: Driver Name
        c_r = ws_sample.cell(r, 18, None)
        c_r.alignment = align_left
        c_r.font = data_font
        c_r.border = cell_border

        # S: Places
        c_s = ws_sample.cell(r, 19, None)
        c_s.alignment = align_left
        c_s.font = data_font
        c_s.border = cell_border

        # T: Payment Date
        c_t = ws_sample.cell(r, 20, None)
        c_t.alignment = align_center
        c_t.font = data_font
        c_t.number_format = "DD/MM/YYYY"
        c_t.border = cell_border

    dv_s_vendor = DataValidation(type="list", formula1="=Vendors!$B$2:$B$50", allow_blank=True)
    ws_sample.add_data_validation(dv_s_vendor)
    dv_s_vendor.add("B2:B200")

    dv_s_vech = DataValidation(type="list", formula1='"Baleno,Swift Dzire,Ertiga,Innova,Crysta,Wagon R,Kia,Sedan,Tempo"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_vech)
    dv_s_vech.add("E2:E200")

    dv_s_rate = DataValidation(type="list", formula1='"18,20,22,23,24,25,28,35"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_rate)
    dv_s_rate.add("J2:J200")

    dv_s_pay = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_pay)
    dv_s_pay.add("O2:O200")

    dv_s_driver = DataValidation(type="list", formula1="=Drivers!$B$2:$B$50", allow_blank=True)
    ws_sample.add_data_validation(dv_s_driver)
    dv_s_driver.add("R2:R200")

    # ----------------------------------------------------
    # Sheet 3: Vendors (Directory of All 14 Active Vendors)
    # ----------------------------------------------------
    ws_vendors = wb.create_sheet("Vendors")
    ws_vendors.views.sheetView[0].showGridLines = True
    ws_vendors.freeze_panes = "A2"
    ws_vendors.row_dimensions[1].height = 28

    v_headers = ["Vendor ID", "Vendor Name", "Contact Person", "Phone Number", "Primary Vehicle Models", "Status"]
    v_widths = {"A": 12, "B": 20, "C": 20, "D": 18, "E": 28, "F": 12}

    for c_idx, vh in enumerate(v_headers, 1):
        cell = ws_vendors.cell(1, c_idx, vh)
        cell.font = header_font
        cell.fill = reference_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in v_widths.items():
        ws_vendors.column_dimensions[c_let].width = w

    vendors_data = [
        (1, "Arvind", "Arvind", "9822150660", "Swift Dzire, Baleno, Kia, Ertiga", "Active"),
        (2, "Bharat", "Bharat", "9890509032", "Ertiga", "Active"),
        (3, "Dilip", "Dilip", "9860051693", "Swift Dzire, Ertiga", "Active"),
        (4, "Gajanan", "Gajanan", "9923921751", "Swift Dzire", "Active"),
        (5, "Harshal", "Harshal", "9021859362", "Swift Dzire", "Active"),
        (6, "Kunal", "Kunal", "7030179838", "Swift Dzire", "Active"),
        (7, "Lalu", "Lalu", "8888943041", "Swift Dzire, Baleno, Ertiga, Crysta", "Active"),
        (8, "Manoj", "Manoj", "7972136310", "Swift Dzire", "Active"),
        (9, "Maulla", "Maulla", "9822153644", "Swift Dzire", "Active"),
        (10, "Pritesh", "Pritesh", "8788173178", "Swift Dzire, Ertiga", "Active"),
        (11, "Satesh", "Satesh", "7249649936", "Swift Dzire, Ertiga", "Active"),
        (12, "Shubham", "Shubham", "7798189510", "Wagon R", "Active"),
        (13, "Subodh", "Subodh", "9765354361", "Swift Dzire", "Active"),
        (14, "Yogesh", "Yogesh", "7498434925", "Ertiga", "Active"),
    ]

    for r_idx, (vid, vname, vcontact, vphone, vmodels, vstatus) in enumerate(vendors_data, 2):
        ws_vendors.row_dimensions[r_idx].height = 22
        ws_vendors.cell(r_idx, 1, vid).alignment = align_center
        ws_vendors.cell(r_idx, 2, vname).alignment = align_left
        ws_vendors.cell(r_idx, 3, vcontact).alignment = align_left
        c_ph = ws_vendors.cell(r_idx, 4, vphone)
        c_ph.alignment = align_center
        c_ph.number_format = "@"
        ws_vendors.cell(r_idx, 5, vmodels).alignment = align_left
        c_st = ws_vendors.cell(r_idx, 6, vstatus)
        c_st.alignment = align_center
        for c in range(1, 7):
            ws_vendors.cell(r_idx, c).font = data_font
            ws_vendors.cell(r_idx, c).border = cell_border

    # ----------------------------------------------------
    # Sheet 4: Drivers (Full Driver Directory with Phones)
    # ----------------------------------------------------
    ws_drivers = wb.create_sheet("Drivers")
    ws_drivers.views.sheetView[0].showGridLines = True
    ws_drivers.freeze_panes = "A2"
    ws_drivers.row_dimensions[1].height = 28

    dr_headers = ["Sr. no", "Driver", "Contact"]
    dr_widths = {"A": 10, "B": 22, "C": 18}

    for c_idx, dh in enumerate(dr_headers, 1):
        cell = ws_drivers.cell(1, c_idx, dh)
        cell.font = header_font
        cell.fill = reference_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in dr_widths.items():
        ws_drivers.column_dimensions[c_let].width = w

    drivers_data = [
        (1, "Abdul", "7276752343"),
        (2, "Abhishek", "8830087698"),
        (3, "Arvind", "9822150660"),
        (4, "Ashok", "9822150661"),
        (5, "Basayya", "9309702950"),
        (6, "Bharat", "9890509032"),
        (7, "Deepak", "8788173179"),
        (8, "Dilip", "9860051693"),
        (9, "Francis", "8999200950"),
        (10, "Gajanan", "9923921751"),
        (11, "Harshal", "9021859362"),
        (12, "Hussain", "8551885657"),
        (13, "Irfan", "9172653036"),
        (14, "Isaik", "8888943042"),
        (15, "Jaffer", "9673552131"),
        (16, "Kunal", "7030179838"),
        (17, "Lalu", "8888943041"),
        (18, "Lokesh", "7507672046"),
        (19, "Malik", "8888943043"),
        (20, "Manjunath", "8668727157"),
        (21, "Manoj", "7972136310"),
        (22, "Maulla", "9822153644"),
        (23, "Naresh", "7972003791"),
        (24, "Pratap", "9765170368"),
        (25, "Pritesh", "8788173178"),
        (26, "Rafiq", "9916957397"),
        (27, "Rahul", "8007029326"),
        (28, "Rajdeep", "8805926693"),
        (29, "Raju", "9579713789"),
        (30, "Sairaj", "9168793714"),
        (31, "Satesh", "7249649936"),
        (32, "Shankar", "9673233774"),
        (33, "Shiva", "9284180938"),
        (34, "Shubham", "7798189510"),
        (35, "Subhash", "9765369632"),
        (36, "Subodh", "9765354361"),
        (37, "Suyog", "7020035307"),
        (38, "Vijay", "9763363821"),
        (39, "Vishal", "7249649937"),
        (40, "Yogesh", "7498434925"),
    ]

    for r_idx, (sr, name, phone) in enumerate(drivers_data, 2):
        ws_drivers.row_dimensions[r_idx].height = 22
        ws_drivers.cell(r_idx, 1, sr).alignment = align_center
        ws_drivers.cell(r_idx, 2, name).alignment = align_left
        c_dp = ws_drivers.cell(r_idx, 3, phone)
        c_dp.alignment = align_center
        c_dp.number_format = "@"
        for c in range(1, 4):
            ws_drivers.cell(r_idx, c).font = data_font
            ws_drivers.cell(r_idx, c).border = cell_border

    # ----------------------------------------------------
    # Sheet 5: Vehicles (Fleet Registry from Masterfile)
    # ----------------------------------------------------
    ws_vech = wb.create_sheet("Vehicles")
    ws_vech.views.sheetView[0].showGridLines = True
    ws_vech.freeze_panes = "A2"
    ws_vech.row_dimensions[1].height = 28

    vh_headers = ["Sr. no", "Vehicle No", "Vehicle Type", "Owner / Vendor"]
    vh_widths = {"A": 10, "B": 20, "C": 18, "D": 20}

    for c_idx, vch_h in enumerate(vh_headers, 1):
        cell = ws_vech.cell(1, c_idx, vch_h)
        cell.font = header_font
        cell.fill = reference_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in vh_widths.items():
        ws_vech.column_dimensions[c_let].width = w

    vehicles_data = [
        (1, "GA 07 T 2608", "Baleno", "Arvind"),
        (2, "GA 07 F 8056", "Swift Dzire", "Arvind"),
        (3, "GA 07 F 8151", "Swift Dzire", "Arvind"),
        (4, "GA 07 F 5145", "Swift Dzire", "Arvind"),
        (5, "GA 06 T 4317", "Swift Dzire", "Gajanan"),
        (6, "GA 11 T 4914", "Swift Dzire", "Harshal"),
        (7, "GA 03 V 6153", "Swift Dzire", "Lalu"),
        (8, "GA 07 T 3003", "Swift Dzire", "Lalu"),
        (9, "GA 07 T 1725", "Baleno", "Lalu"),
        (10, "GA 03 V 6158", "Swift Dzire", "Lalu"),
        (11, "GA 07 F 4403", "Ertiga", "Bharat"),
        (12, "GA 06 T 4967", "Swift Dzire", "Dilip"),
        (13, "GA 06 T 5992", "Swift Dzire", "Dilip"),
        (14, "GA 06 T 5030", "Swift Dzire", "Dilip"),
        (15, "GA 03 AH 9710", "Swift Dzire", "Kunal"),
        (16, "GA 06 T 5004", "Swift Dzire", "Manoj"),
        (17, "GA 07 F 5292", "Swift Dzire", "Maulla"),
        (18, "GA 11 T 5181", "Swift Dzire", "Pritesh"),
        (19, "GA 11 T 5419", "Swift Dzire", "Pritesh"),
        (20, "GA 11 T 4807", "Swift Dzire", "Pritesh"),
        (21, "GA 11 T 6060", "Swift Dzire", "Satesh"),
        (22, "GA 11 T 2332", "Swift Dzire", "Satesh"),
        (23, "GA 11 T 3960", "Ertiga", "Satesh"),
        (24, "GA 05 T 4997", "Wagon R", "Shubham"),
        (25, "GA 11 T 4815", "Swift Dzire", "Subodh"),
        (26, "GA 05 T 8121", "Ertiga", "Yogesh"),
    ]

    for r_idx, (sr, vno, vtype, vowner) in enumerate(vehicles_data, 2):
        ws_vech.row_dimensions[r_idx].height = 22
        ws_vech.cell(r_idx, 1, sr).alignment = align_center
        ws_vech.cell(r_idx, 2, vno).alignment = align_center
        ws_vech.cell(r_idx, 3, vtype).alignment = align_center
        ws_vech.cell(r_idx, 4, vowner).alignment = align_left
        for c in range(1, 5):
            ws_vech.cell(r_idx, c).font = data_font
            ws_vech.cell(r_idx, c).border = cell_border

    # ----------------------------------------------------
    # Sheet 6: Rate_Card (Package Rules & Pricing Reference)
    # ----------------------------------------------------
    ws_rates = wb.create_sheet("Rate_Card")
    ws_rates.views.sheetView[0].showGridLines = True
    ws_rates.freeze_panes = "A2"
    ws_rates.row_dimensions[1].height = 28

    rc_headers = ["Rate (Rs/Km)", "80 Km Pkg Base (Rs)", "Extra Km Rate (Rs/Km)", "Primary Vehicle Type", "Remarks & Applicability"]
    rc_widths = {"A": 15, "B": 22, "C": 22, "D": 25, "E": 36}

    for c_idx, rch in enumerate(rc_headers, 1):
        cell = ws_rates.cell(1, c_idx, rch)
        cell.font = header_font
        cell.fill = reference_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in rc_widths.items():
        ws_rates.column_dimensions[c_let].width = w

    rates_data = [
        (18, 1800, 18, "Baleno / Hatchback", "Standard local 80-km package (Rs 18/km above 80)"),
        (20, 2000, 20, "Swift Dzire / Sedan", "Standard 80-km package (Rs 20/km above 80)"),
        (22, 2200, 22, "Swift Dzire / Baleno", "Primary executive 80-km package (Arvind vendor rate)"),
        (23, 2300, 23, "Sedan / Compact MUV", "Custom route package"),
        (24, 2400, 23, "Ertiga / MUV", "80-km package base Rs 2400, Rs 23/km above 80"),
        (25, 2500, 23, "Ertiga / Crysta", "80-km package base Rs 2500, Rs 23/km above 80"),
        (28, 2800, 28, "Innova / Crysta", "Premium executive 80-km package"),
        (35, 3500, 25, "Crysta / Luxury", "80-km package base Rs 3500, Rs 25/km above 80"),
    ]

    for r_idx, (rt, base_pkg, extra_km, vmod, rem) in enumerate(rates_data, 2):
        ws_rates.row_dimensions[r_idx].height = 22
        ws_rates.cell(r_idx, 1, rt).alignment = align_center
        c_base = ws_rates.cell(r_idx, 2, base_pkg)
        c_base.alignment = align_right
        c_base.number_format = "₹#,##0.00"
        c_extra = ws_rates.cell(r_idx, 3, extra_km)
        c_extra.alignment = align_right
        c_extra.number_format = "₹#,##0.00"
        ws_rates.cell(r_idx, 4, vmod).alignment = align_left
        ws_rates.cell(r_idx, 5, rem).alignment = align_left
        for c in range(1, 6):
            ws_rates.cell(r_idx, c).font = data_font
            ws_rates.cell(r_idx, c).border = cell_border

    # ----------------------------------------------------
    # Sheet 7: Annexure_Sample (Mirroring view_annexure_arvind)
    # ----------------------------------------------------
    ws_annex = wb.create_sheet("Annexure_Sample")
    ws_annex.views.sheetView[0].showGridLines = True
    ws_annex.freeze_panes = "A2"
    ws_annex.row_dimensions[1].height = 28

    annex_headers = [
        "Sl. No",
        "Travel Date",
        "Guest Name",
        "Vehicle No",
        "Vehicle Type",
        "Starting Reading",
        "Ending Reading",
        "Total Kms",
        "Rate Per KM Above 80",
        "Total KM Rs",
        "Toll Charges",
        "Night Charges",
        "Waiting Charges",
        "Total Rs",
        "Timing",
        "Driver",
    ]
    annex_widths = {
        "A": 9, "B": 14, "C": 20, "D": 18, "E": 16,
        "F": 16, "G": 16, "H": 13, "I": 20, "J": 14,
        "K": 14, "L": 14, "M": 15, "N": 16, "O": 22, "P": 18,
    }

    for c_idx, ah in enumerate(annex_headers, 1):
        cell = ws_annex.cell(1, c_idx, ah)
        cell.font = header_font
        cell.fill = primary_header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in annex_widths.items():
        ws_annex.column_dimensions[c_let].width = w

    # Arvind vendor sample entries mirroring view_annexure_arvind SQL view
    annex_rows = [
        (1, "01/12/2025", "BDA PL", "GA 07 T 2608", "Baleno", 59989, 60058, 69, 22, 2200, 0, 0, 0, 2200, "8:15 AM - 10:30 AM", "Abdul"),
        (2, "01/12/2025", "GIM", "GA 07 T 2608", "Baleno", 60030, 60140, 110, 22, 2860, 0, 100, 0, 2960, "4:30 PM - 7:30 PM", "Abdul"),
        (3, "01/12/2025", "GIM", "GA 07  8151", "Swift Dzire", 250355, 250582, 227, 22, 5434, 0, 400, 0, 5834, "1:30 PM - 11:00 PM", "Suyog"),
        (4, "01/12/2025", "GIM", "GA 07 F 8056", "Swift Dzire", 230280, 230390, 110, 22, 2860, 50, 100, 0, 3010, "5:00 AM - 9:00 AM", "Rahul"),
    ]

    for r_idx, arow in enumerate(annex_rows, 2):
        ws_annex.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(arow, 1):
            cell = ws_annex.cell(r_idx, c_idx, val)
            cell.font = data_font
            cell.border = cell_border
            if c_idx in (1, 2, 4, 5, 9, 15):
                cell.alignment = align_center
            elif c_idx in (3, 16):
                cell.alignment = align_left
            else:
                cell.alignment = align_right
                if c_idx in (6, 7, 8):
                    cell.number_format = "#,##0.00"
                elif c_idx in (10, 11, 12, 13, 14):
                    cell.number_format = "#,##0.00"

    # Save to primary and additional paths
    primary_path = Path(output_path)
    primary_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(primary_path)

    for add_path in additional_paths:
        p = Path(add_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        wb.save(p)

    wb.close()
    return primary_path


if __name__ == "__main__":
    generate_masterfile_template(
        output_path="masterfile_template.xlsx",
        additional_paths=["Requirements/Input/masterfile_template.xlsx"],
    )
    print("Masterfile template generation complete.")
