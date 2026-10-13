"""Excel bookings template generator matching database schema and Jan 2026 conventions."""

import os
from pathlib import Path
from typing import List, Sequence, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


def generate_bookings_template(
    output_path: Union[str, Path] = "bookings_template.xlsx",
    additional_paths: Sequence[Union[str, Path]] = (),
) -> Path:
    """Create a standardized Excel workbook template for bookings and driver contacts.

    Features:
    - Matches MySQL bookings schema columns (received, sl_no, coordinator/name, guest,
      guest_mobile_no, travel_date, from, to, start_time, return_time, reason,
      driver, contact, vehicle_type/vech_no, company, driver_message, guest_message).
    - Auto-lookup formula for driver contact from the Contacts worksheet.
    - Automated formulas for Driver Message and Guest Message matching the stored generated columns.
    - Full list of active drivers and contacts pre-populated on the Contacts worksheet.
    - Data validation dropdowns for Received (TRUE/FALSE), Driver (from Contacts),
      Vehicle Type (Ertiga, Sedan, Innova, etc.), and Company (GIM, Vendor, etc.).
    - Professional styling faithful to the Jan 2026 bookings workbook (Trebuchet MS, Soft Rose palette).
    """
    wb = openpyxl.Workbook()

    # Font & Color Palette
    font_family = "Trebuchet MS"
    header_fill = PatternFill(start_color="FFFFD5E6", end_color="FFFFD5E6", fill_type="solid")  # Jan 2026 Soft Rose
    accent_fill = PatternFill(start_color="FFE8F0FE", end_color="FFE8F0FE", fill_type="solid")  # Soft Blue for messages

    header_font = Font(name=font_family, size=11, bold=True, color="FF000000")
    header_font_accent = Font(name=font_family, size=11, bold=True, color="FF1A73E8")
    data_font = Font(name=font_family, size=10.5, bold=False, color="FF000000")
    formula_font = Font(name=font_family, size=10, bold=False, color="FF202124")

    thin_side = Side(border_style="thin", color="FFD3D3D3")
    header_side = Side(border_style="thin", color="FFBFBFBF")

    header_border = Border(left=header_side, right=header_side, top=header_side, bottom=header_side)
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_msg = Alignment(horizontal="left", vertical="top", wrap_text=True)

    headers = [
        "Received",
        "Sl.No.",
        "Name",
        "Guest",
        "Guest Mobile No.",
        "Travel Date",
        "From",
        "To",
        "Start Time",
        "Return Time",
        "Reason",
        "Driver",
        "Contact",
        "Vech No",
        "Company",
        "Driver Message",
        "Guest Message",
    ]

    col_widths = {
        "A": 10,  # Received
        "B": 8,   # Sl.No.
        "C": 18,  # Name (Coordinator)
        "D": 22,  # Guest
        "E": 18,  # Guest Mobile No.
        "F": 15,  # Travel Date
        "G": 24,  # From
        "H": 24,  # To
        "I": 14,  # Start Time
        "J": 14,  # Return Time
        "K": 32,  # Reason
        "L": 18,  # Driver
        "M": 16,  # Contact
        "N": 16,  # Vech No
        "O": 14,  # Company
        "P": 38,  # Driver Message
        "Q": 46,  # Guest Message
    }

    # ----------------------------------------------------
    # Sheet 1: Bookings (Sample & Active Entry Sheet)
    # ----------------------------------------------------
    ws_bookings = wb.active
    ws_bookings.title = "Bookings"
    ws_bookings.views.sheetView[0].showGridLines = True
    ws_bookings.freeze_panes = "A2"
    ws_bookings.row_dimensions[1].height = 28

    for col_idx, h in enumerate(headers, 1):
        cell = ws_bookings.cell(1, col_idx, h)
        cell.font = header_font_accent if col_idx in (16, 17) else header_font
        cell.fill = accent_fill if col_idx in (16, 17) else header_fill
        cell.alignment = align_center
        cell.border = header_border

    for col_letter, width in col_widths.items():
        ws_bookings.column_dimensions[col_letter].width = width

    # Sample rows demonstrating real-world usage
    sample_records = [
        {
            "received": True,
            "sl_no": 1,
            "coordinator": "Gayathri",
            "guest": "Dr. Divya Singhal",
            "mobile": "9922910939",
            "date": "29-01-2026",
            "from": "GIM Campus",
            "to": "Old Goa",
            "start": "10:45 AM",
            "return": "12:40 PM",
            "reason": "Keynote speaker at Business Conclave",
            "driver": "Rahul",
            "vech": "Ertiga",
            "company": "GIM",
        },
        {
            "received": True,
            "sl_no": 2,
            "coordinator": "Kingshuk Sarkar",
            "guest": "Kingshuk Sarkar",
            "mobile": "9007704524",
            "date": "29-01-2026",
            "from": "Dabolim Airport",
            "to": "GIM Campus",
            "start": "03:00 PM",
            "return": "04:30 PM",
            "reason": "Return to campus after interview round",
            "driver": "Suyog",
            "vech": "Sedan",
            "company": "GIM",
        },
        {
            "received": False,
            "sl_no": 3,
            "coordinator": "Palak Pangam",
            "guest": "Mr. Brendon Mascarenhas",
            "mobile": "9421346240",
            "date": "30-01-2026",
            "from": "Porvorim",
            "to": "GIM Campus and back",
            "start": "05:30 PM",
            "return": "10:00 PM",
            "reason": "Visiting Faculty for PGP II course",
            "driver": "Rafiq",
            "vech": "Ertiga",
            "company": "GIM",
        },
    ]

    total_rows = 35
    for r in range(2, total_rows + 2):
        ws_bookings.row_dimensions[r].height = 42 if r <= 4 else 26
        is_sample = (r - 2) < len(sample_records)
        rec = sample_records[r - 2] if is_sample else None

        # A: Received
        c_a = ws_bookings.cell(r, 1, rec["received"] if is_sample else False)
        c_a.alignment = align_center
        c_a.font = data_font
        c_a.border = cell_border

        # B: Sl.No.
        c_b = ws_bookings.cell(r, 2, rec["sl_no"] if is_sample else None)
        c_b.alignment = align_center
        c_b.font = data_font
        c_b.border = cell_border

        # C: Name
        c_c = ws_bookings.cell(r, 3, rec["coordinator"] if is_sample else None)
        c_c.alignment = align_left
        c_c.font = data_font
        c_c.border = cell_border

        # D: Guest
        c_d = ws_bookings.cell(r, 4, rec["guest"] if is_sample else None)
        c_d.alignment = align_left
        c_d.font = data_font
        c_d.border = cell_border

        # E: Guest Mobile No.
        c_e = ws_bookings.cell(r, 5, rec["mobile"] if is_sample else None)
        c_e.alignment = align_center
        c_e.font = data_font
        c_e.number_format = "@"
        c_e.border = cell_border

        # F: Travel Date
        c_f = ws_bookings.cell(r, 6, rec["date"] if is_sample else None)
        c_f.alignment = align_center
        c_f.font = data_font
        c_f.number_format = "DD-MM-YYYY"
        c_f.border = cell_border

        # G: From
        c_g = ws_bookings.cell(r, 7, rec["from"] if is_sample else None)
        c_g.alignment = align_left
        c_g.font = data_font
        c_g.border = cell_border

        # H: To
        c_h = ws_bookings.cell(r, 8, rec["to"] if is_sample else None)
        c_h.alignment = align_left
        c_h.font = data_font
        c_h.border = cell_border

        # I: Start Time
        c_i = ws_bookings.cell(r, 9, rec["start"] if is_sample else None)
        c_i.alignment = align_center
        c_i.font = data_font
        c_i.border = cell_border

        # J: Return Time
        c_j = ws_bookings.cell(r, 10, rec["return"] if is_sample else None)
        c_j.alignment = align_center
        c_j.font = data_font
        c_j.border = cell_border

        # K: Reason
        c_k = ws_bookings.cell(r, 11, rec["reason"] if is_sample else None)
        c_k.alignment = align_left
        c_k.font = data_font
        c_k.border = cell_border

        # L: Driver
        c_l = ws_bookings.cell(r, 12, rec["driver"] if is_sample else None)
        c_l.alignment = align_left
        c_l.font = data_font
        c_l.border = cell_border

        # M: Contact (Formula)
        contact_formula = f'=IF(L{r}="","",_xlfn.XLOOKUP(L{r}, Contacts!$B$2:$B$100, Contacts!$C$2:$C$100, ""))'
        c_m = ws_bookings.cell(r, 13, contact_formula)
        c_m.alignment = align_center
        c_m.font = data_font
        c_m.number_format = "@"
        c_m.border = cell_border

        # N: Vech No
        c_n = ws_bookings.cell(r, 14, rec["vech"] if is_sample else None)
        c_n.alignment = align_center
        c_n.font = data_font
        c_n.border = cell_border

        # O: Company
        c_o = ws_bookings.cell(r, 15, rec["company"] if is_sample else None)
        c_o.alignment = align_center
        c_o.font = data_font
        c_o.border = cell_border

        # P: Driver Message Formula
        driver_msg_formula = (
            f'=IF(D{r}="","","Tomorrow" & CHAR(10) & D{r} & '
            f'IF(ISBLANK(E{r}),""," " & E{r}) & CHAR(10) & '
            f'G{r} & " to " & H{r} & CHAR(10) & '
            f'IF(ISBLANK(I{r}),"",TEXT(I{r}, "h:mm AM/PM")))'
        )
        c_p = ws_bookings.cell(r, 16, driver_msg_formula)
        c_p.alignment = align_msg
        c_p.font = formula_font
        c_p.border = cell_border

        # Q: Guest Message Formula
        guest_msg_formula = (
            f'=IF(L{r}="","","From " & IF(OR(ISBLANK(O{r}),O{r}=""),"GIM",O{r}) & CHAR(10) & '
            f'"Driver " & L{r} & IF(ISBLANK(M{r}),""," " & M{r}) & '
            f'" will come to pick you tomorrow at " & '
            f'IF(ISBLANK(I{r}),"",TEXT(I{r}, "h:mm AM/PM")) & " at " & G{r})'
        )
        c_q = ws_bookings.cell(r, 17, guest_msg_formula)
        c_q.alignment = align_msg
        c_q.font = formula_font
        c_q.border = cell_border

    # Data Validations on Bookings
    dv_received = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_bookings.add_data_validation(dv_received)
    dv_received.add("A2:A100")

    dv_driver = DataValidation(type="list", formula1="=Contacts!$B$2:$B$100", allow_blank=True)
    ws_bookings.add_data_validation(dv_driver)
    dv_driver.add("L2:L100")

    dv_vech = DataValidation(type="list", formula1='"Ertiga,Sedan,Innova,Swift Dzire,Wagon R,Crysta,Tempo"', allow_blank=True)
    ws_bookings.add_data_validation(dv_vech)
    dv_vech.add("N2:N100")

    dv_company = DataValidation(type="list", formula1='"GIM,Vendor,Private,Other"', allow_blank=True)
    ws_bookings.add_data_validation(dv_company)
    dv_company.add("O2:O100")

    # ----------------------------------------------------
    # Sheet 2: Sample (Clean Daily Template to Duplicate)
    # ----------------------------------------------------
    ws_sample = wb.create_sheet("Sample")
    ws_sample.views.sheetView[0].showGridLines = True
    ws_sample.freeze_panes = "A2"
    ws_sample.row_dimensions[1].height = 28

    for col_idx, h in enumerate(headers, 1):
        cell = ws_sample.cell(1, col_idx, h)
        cell.font = header_font_accent if col_idx in (16, 17) else header_font
        cell.fill = accent_fill if col_idx in (16, 17) else header_fill
        cell.alignment = align_center
        cell.border = header_border

    for col_letter, width in col_widths.items():
        ws_sample.column_dimensions[col_letter].width = width

    for r in range(2, 27):
        ws_sample.row_dimensions[r].height = 26

        c_a = ws_sample.cell(r, 1, False)
        c_a.alignment = align_center
        c_a.font = data_font
        c_a.border = cell_border

        c_b = ws_sample.cell(r, 2, None)
        c_b.alignment = align_center
        c_b.font = data_font
        c_b.border = cell_border

        for c in range(3, 12):
            cell = ws_sample.cell(r, c)
            cell.alignment = align_center if c in (5, 6, 9, 10) else align_left
            cell.font = data_font
            cell.border = cell_border
            if c == 5:
                cell.number_format = "@"
            elif c == 6:
                cell.number_format = "DD-MM-YYYY"

        c_l = ws_sample.cell(r, 12)
        c_l.alignment = align_left
        c_l.font = data_font
        c_l.border = cell_border

        contact_formula = f'=IF(L{r}="","",_xlfn.XLOOKUP(L{r}, Contacts!$B$2:$B$100, Contacts!$C$2:$C$100, ""))'
        c_m = ws_sample.cell(r, 13, contact_formula)
        c_m.alignment = align_center
        c_m.font = data_font
        c_m.number_format = "@"
        c_m.border = cell_border

        c_n = ws_sample.cell(r, 14)
        c_n.alignment = align_center
        c_n.font = data_font
        c_n.border = cell_border

        c_o = ws_sample.cell(r, 15)
        c_o.alignment = align_center
        c_o.font = data_font
        c_o.border = cell_border

        driver_msg_formula = (
            f'=IF(D{r}="","","Tomorrow" & CHAR(10) & D{r} & '
            f'IF(ISBLANK(E{r}),""," " & E{r}) & CHAR(10) & '
            f'G{r} & " to " & H{r} & CHAR(10) & '
            f'IF(ISBLANK(I{r}),"",TEXT(I{r}, "h:mm AM/PM")))'
        )
        c_p = ws_sample.cell(r, 16, driver_msg_formula)
        c_p.alignment = align_msg
        c_p.font = formula_font
        c_p.border = cell_border

        guest_msg_formula = (
            f'=IF(L{r}="","","From " & IF(OR(ISBLANK(O{r}),O{r}=""),"GIM",O{r}) & CHAR(10) & '
            f'"Driver " & L{r} & IF(ISBLANK(M{r}),""," " & M{r}) & '
            f'" will come to pick you tomorrow at " & '
            f'IF(ISBLANK(I{r}),"",TEXT(I{r}, "h:mm AM/PM")) & " at " & G{r})'
        )
        c_q = ws_sample.cell(r, 17, guest_msg_formula)
        c_q.alignment = align_msg
        c_q.font = formula_font
        c_q.border = cell_border

    dv_s_rec = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_rec)
    dv_s_rec.add("A2:A100")

    dv_s_driver = DataValidation(type="list", formula1="=Contacts!$B$2:$B$100", allow_blank=True)
    ws_sample.add_data_validation(dv_s_driver)
    dv_s_driver.add("L2:L100")

    dv_s_vech = DataValidation(type="list", formula1='"Ertiga,Sedan,Innova,Swift Dzire,Wagon R,Crysta,Tempo"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_vech)
    dv_s_vech.add("N2:N100")

    dv_s_comp = DataValidation(type="list", formula1='"GIM,Vendor,Private,Other"', allow_blank=True)
    ws_sample.add_data_validation(dv_s_comp)
    dv_s_comp.add("O2:O100")

    # ----------------------------------------------------
    # Sheet 3: Contacts (Full 35-Driver Directory from Jan 2026)
    # ----------------------------------------------------
    ws_contacts = wb.create_sheet("Contacts")
    ws_contacts.views.sheetView[0].showGridLines = True
    ws_contacts.freeze_panes = "A2"
    ws_contacts.row_dimensions[1].height = 28

    c_headers = ["Sr. no", "Driver", "Contact"]
    c_widths = {"A": 10, "B": 22, "C": 18}

    for c_idx, ch in enumerate(c_headers, 1):
        cell = ws_contacts.cell(1, c_idx, ch)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = header_border

    for c_let, w in c_widths.items():
        ws_contacts.column_dimensions[c_let].width = w

    drivers_data = [
        (1, "Abhishek", "8830087698"),
        (2, "Arvind", "9822150660"),
        (3, "Abdul", "7276752343"),
        (4, "Rahul", "8007029326"),
        (5, "Suyog", "7020035307"),
        (6, "Yogesh", "7498434925"),
        (7, "Shubham", "7798189510"),
        (8, "Lalu", "8888943041"),
        (9, "Dilip", "9860051693"),
        (10, "Lokesh", "7507672046"),
        (11, "Subhash", "9765369632"),
        (12, "Sairaj", "9168793714"),
        (13, "Kunal", "7030179838"),
        (14, "Harshal", "9021859362"),
        (15, "Pritesh", "8788173178"),
        (16, "Pratap", "9765170368"),
        (17, "Maulla", "9822153644"),
        (18, "Shiva", "9284180938"),
        (19, "Naresh", "7972003791"),
        (20, "Basayya", "9309702950"),
        (21, "Raju", "9579713789"),
        (22, "Gajanan", "9923921751"),
        (23, "Hussain", "8551885657"),
        (24, "Irfan", "9172653036"),
        (25, "Manoj", "7972136310"),
        (26, "Satesh", "7249649936"),
        (27, "Bharat", "9890509032"),
        (28, "Jaffer", "9673552131"),
        (29, "Vijay", "9763363821"),
        (30, "Francis", "8999200950"),
        (31, "Subodh", "9765354361"),
        (32, "Manjunath", "8668727157"),
        (33, "Shankar", "9673233774"),
        (34, "Rajdeep", "8805926693"),
        (35, "Rafiq", "9916957397"),
    ]

    for r_idx, (sr, name, phone) in enumerate(drivers_data, 2):
        ws_contacts.row_dimensions[r_idx].height = 22

        c1 = ws_contacts.cell(r_idx, 1, sr)
        c1.alignment = align_center
        c1.font = data_font
        c1.border = cell_border

        c2 = ws_contacts.cell(r_idx, 2, name)
        c2.alignment = align_left
        c2.font = data_font
        c2.border = cell_border

        c3 = ws_contacts.cell(r_idx, 3, phone)
        c3.alignment = align_center
        c3.font = data_font
        c3.number_format = "@"
        c3.border = cell_border

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
    generate_bookings_template(
        output_path="bookings_template.xlsx",
        additional_paths=["Requirements/Input/bookings_template.xlsx"],
    )
    print("Template generation complete.")
