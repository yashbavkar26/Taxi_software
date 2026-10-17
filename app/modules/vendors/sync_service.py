"""Vendor Synchronization Service.

Automatically partitions and synchronizes entries from masterfile_template.xlsx
into individual vendor workbooks (e.g. Arvind.xlsx, Bharat.xlsx) matching the
exact formatting of the user's reference file.
"""

import os
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


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
    "A": 9.0,   # Sl. No
    "B": 16.0,  # Vendor
    "C": 14.0,  # Date
    "D": 18.0,  # Vech.no
    "E": 16.0,  # Vech Type
    "F": 15.0,  # Starting Km
    "G": 13.0,  # Ending Km
    "H": 13.0,  # Total Km
    "I": 12.0,  # Night
    "J": 10.0,  # Rate
    "K": 12.0,  # Toll
    "L": 14.0,  # Amt
    "M": 22.0,  # Guest
    "N": 16.0,  # Total Amt
    "O": 13.0,  # Payment Details
    "P": 13.0,  # In Time
    "Q": 13.0,  # Out Time
    "R": 20.0,  # Driver Name
    "S": 22.0,  # Places
    "T": 15.0,  # Payment Date
}


def read_masterfile_trips(masterfile_path: Union[str, Path] = "masterfile_template.xlsx") -> Dict[str, List[Dict[str, Any]]]:
    """Read all trips from the masterfile workbook and group them by vendor.

    Args:
        masterfile_path: Path to masterfile_template.xlsx or equivalent.

    Returns:
        Dict mapping vendor name to list of trip record dictionaries.
    """
    path = Path(masterfile_path)
    if not path.exists():
        raise FileNotFoundError(f"Masterfile not found at '{path}'")

    wb = openpyxl.load_workbook(path, data_only=False)
    # Find Masterfile sheet
    sheet_name = "Masterfile" if "Masterfile" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    vendor_trips: Dict[str, List[Dict[str, Any]]] = {}

    for r in range(2, ws.max_row + 1):
        vendor_val = ws.cell(r, 2).value
        if not vendor_val or str(vendor_val).strip() == "":
            continue

        vendor_name = str(vendor_val).strip()
        trip = {
            "vendor": vendor_name,
            "date": ws.cell(r, 3).value,
            "vech_no": ws.cell(r, 4).value,
            "vech_type": ws.cell(r, 5).value,
            "start_km": ws.cell(r, 6).value,
            "end_km": ws.cell(r, 7).value,
            "total_km": ws.cell(r, 8).value,
            "night": ws.cell(r, 9).value,
            "rate": ws.cell(r, 10).value,
            "toll": ws.cell(r, 11).value,
            "amt": ws.cell(r, 12).value,
            "guest": ws.cell(r, 13).value,
            "total_amt": ws.cell(r, 14).value,
            "payment_details": ws.cell(r, 15).value if ws.cell(r, 15).value is not None else False,
            "in_time": ws.cell(r, 16).value,
            "out_time": ws.cell(r, 17).value,
            "driver": ws.cell(r, 18).value,
            "places": ws.cell(r, 19).value,
            "payment_date": ws.cell(r, 20).value,
        }

        if vendor_name not in vendor_trips:
            vendor_trips[vendor_name] = []
        vendor_trips[vendor_name].append(trip)

    wb.close()
    return vendor_trips


def write_vendor_workbook(
    vendor_name: str,
    trips: List[Dict[str, Any]],
    output_path: Union[str, Path],
) -> Path:
    """Generate or update a vendor workbook matching the exact Arvind.xlsx formatting.

    Args:
        vendor_name: Name of the vendor.
        trips: List of trip records for this vendor.
        output_path: Destination file path.
    """
    out_file = Path(output_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendor_Trips"
    ws.freeze_panes = "A2"
    ws.views.sheetView[0].showGridLines = True

    font_family = "Trebuchet MS"

    # Styling definitions matching Arvind.xlsx
    header_font = Font(name=font_family, size=10.5, bold=True, color="FFFFFFFF")
    header_fill = PatternFill(start_color="FF1A73E8", end_color="FF1A73E8", fill_type="solid")
    data_font = Font(name=font_family, size=10, bold=False, color="FF0F172A")
    formula_data_font = Font(name=font_family, size=10, bold=False, color="FF0F172A")

    thin_border_side = Side(border_style="thin", color="FFCBD5E1")
    header_border_side = Side(border_style="medium", color="FF1A73E8")

    header_border = Border(
        left=thin_border_side, right=thin_border_side, top=header_border_side, bottom=header_border_side
    )
    cell_border = Border(
        left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Row 1: Header
    ws.row_dimensions[1].height = 28.05
    for c_idx, h_text in enumerate(MASTERFILE_HEADERS, 1):
        cell = ws.cell(1, c_idx, h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = header_border

    # Set Column Widths
    for col_letter, w in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = w

    # Rows 2..N: Trips Data
    for idx, trip in enumerate(trips, 1):
        r = idx + 1
        ws.row_dimensions[r].height = 22.05

        # Col A (1): Sl. No
        c_a = ws.cell(r, 1, idx)
        c_a.alignment = align_center

        # Col B (2): Vendor
        c_b = ws.cell(r, 2, vendor_name)
        c_b.alignment = align_left

        # Col C (3): Date
        c_c = ws.cell(r, 3, trip.get("date"))
        c_c.alignment = align_center
        c_c.number_format = "dd/mm/yyyy"

        # Col D (4): Vech.no
        c_d = ws.cell(r, 4, trip.get("vech_no"))
        c_d.alignment = align_center

        # Col E (5): Vech Type
        c_e = ws.cell(r, 5, trip.get("vech_type"))
        c_e.alignment = align_center

        # Col F (6): Starting Km
        c_f = ws.cell(r, 6, trip.get("start_km"))
        c_f.alignment = align_right
        c_f.number_format = "#,##0.00"

        # Col G (7): Ending Km
        c_g = ws.cell(r, 7, trip.get("end_km"))
        c_g.alignment = align_right
        c_g.number_format = "#,##0.00"

        # Col H (8): Total Km (Formula)
        c_h = ws.cell(r, 8, f'=IF(OR(ISBLANK(G{r}),ISBLANK(F{r})),"",G{r}-F{r})')
        c_h.alignment = align_right
        c_h.font = formula_data_font
        c_h.number_format = "#,##0.00"

        # Col I (9): Night
        c_i = ws.cell(r, 9, trip.get("night"))
        c_i.alignment = align_right
        c_i.number_format = "#,##0.00"

        # Col J (10): Rate
        c_j = ws.cell(r, 10, trip.get("rate"))
        c_j.alignment = align_center
        c_j.number_format = "0"

        # Col K (11): Toll
        c_k = ws.cell(r, 11, trip.get("toll"))
        c_k.alignment = align_right
        c_k.number_format = "#,##0.00"

        # Col L (12): Amt (Formula or fixed numeric package)
        orig_amt = trip.get("amt")
        if isinstance(orig_amt, (int, float)) and not str(orig_amt).startswith("="):
            amt_val = orig_amt
        else:
            amt_val = f'=IF(OR(ISBLANK(H{r}),ISBLANK(J{r}),J{r}=""),"",IF(J{r}=35,IF(H{r}<=80,3500,3500+((H{r}-80)*25)),IF(H{r}<=80,J{r}*100,J{r}*100+((H{r}-80)*J{r}))))'
        c_l = ws.cell(r, 12, amt_val)
        c_l.alignment = align_right
        c_l.font = formula_data_font
        c_l.number_format = "#,##0.00"

        # Col M (13): Guest
        c_m = ws.cell(r, 13, trip.get("guest"))
        c_m.alignment = align_left

        # Col N (14): Total Amt (Formula)
        c_n = ws.cell(r, 14, f'=IF(AND(ISBLANK(L{r}),ISBLANK(K{r}),ISBLANK(I{r})),"",SUM(L{r},K{r},I{r}))')
        c_n.alignment = align_right
        c_n.font = formula_data_font
        c_n.number_format = "#,##0.00"

        # Col O (15): Payment Details
        c_o = ws.cell(r, 15, trip.get("payment_details", False))
        c_o.alignment = align_center

        # Col P (16): In Time
        c_p = ws.cell(r, 16, trip.get("in_time"))
        c_p.alignment = align_center

        # Col Q (17): Out Time
        c_q = ws.cell(r, 17, trip.get("out_time"))
        c_q.alignment = align_center

        # Col R (18): Driver Name
        c_r = ws.cell(r, 18, trip.get("driver"))
        c_r.alignment = align_left

        # Col S (19): Places
        c_s = ws.cell(r, 19, trip.get("places"))
        c_s.alignment = align_left

        # Col T (20): Payment Date
        c_t = ws.cell(r, 20, trip.get("payment_date"))
        c_t.alignment = align_center
        c_t.number_format = "dd/mm/yyyy"

        # Apply fonts and borders across all columns in row
        for c_idx in range(1, 21):
            cell = ws.cell(r, c_idx)
            cell.font = data_font
            cell.border = cell_border

    out_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_file)
    wb.close()
    return out_file


def sync_all_vendors(
    masterfile_path: Union[str, Path] = "masterfile_template.xlsx",
    output_dir: Union[str, Path] = ".",
) -> Dict[str, int]:
    """Sync all vendor entries from masterfile into individual vendor workbooks.

    Ensures all registered active vendors plus any newly encountered vendors in the
    masterfile get an up-to-date, formatted workbook.

    Args:
        masterfile_path: Path to masterfile workbook.
        output_dir: Directory where vendor workbooks are stored.

    Returns:
        Dict mapping vendor name to count of trips synchronized.
    """
    out_dir = Path(output_dir)
    vendor_trips = read_masterfile_trips(masterfile_path)

    # Union of registered active vendors + all vendors present in masterfile
    all_vendors: Set[str] = set(ACTIVE_VENDORS).union(vendor_trips.keys())

    sync_summary: Dict[str, int] = {}
    for vendor in sorted(all_vendors):
        trips = vendor_trips.get(vendor, [])
        v_file = out_dir / f"{vendor}.xlsx"
        write_vendor_workbook(vendor, trips, v_file)
        sync_summary[vendor] = len(trips)

    return sync_summary


def _safe_print(*args, **kwargs):
    import sys
    if sys.stdout is not None:
        try:
            print(*args, **kwargs)
        except Exception:
            pass


def watch_and_sync(
    masterfile_path: Union[str, Path] = "masterfile_template.xlsx",
    output_dir: Union[str, Path] = ".",
    poll_interval: float = 1.0,
):
    """Continuously monitor masterfile_path and trigger auto-sync on save.

    Args:
        masterfile_path: Path to masterfile workbook to watch.
        output_dir: Directory where vendor workbooks are stored.
        poll_interval: Seconds between file modification checks.
    """
    m_path = Path(masterfile_path)
    _safe_print(f"[*] Starting Vendor Auto-Sync Watcher...")
    _safe_print(f"[*] Monitoring: {m_path.resolve()}")
    _safe_print(f"[*] Output Dir: {Path(output_dir).resolve()}")
    _safe_print("[*] Press Ctrl+C to stop.")

    last_mtime = m_path.stat().st_mtime if m_path.exists() else 0
    last_size = m_path.stat().st_size if m_path.exists() else 0

    # Initial sync on startup
    try:
        summary = sync_all_vendors(m_path, output_dir)
        _safe_print(f"[*] Initial sync complete: {len(summary)} vendor workbooks up to date.")
    except Exception as e:
        _safe_print(f"[!] Warning on initial sync: {e}")

    while True:
        try:
            time.sleep(poll_interval)
            if not m_path.exists():
                continue

            current_mtime = m_path.stat().st_mtime
            current_size = m_path.stat().st_size

            if current_mtime != last_mtime or current_size != last_size:
                # Wait briefly for Excel to finish writing
                time.sleep(0.4)
                try:
                    summary = sync_all_vendors(m_path, output_dir)
                    last_mtime = m_path.stat().st_mtime
                    last_size = m_path.stat().st_size
                    timestamp = time.strftime("%H:%M:%S")
                    active_count = sum(1 for c in summary.values() if c > 0)
                    total_trips = sum(summary.values())
                    _safe_print(
                        f"[{timestamp}] Auto-Synced {len(summary)} vendor files "
                        f"({active_count} active with {total_trips} trips)."
                    )
                except (PermissionError, IOError):
                    # Excel is still actively writing the file; wait and retry on next tick
                    time.sleep(0.5)

        except KeyboardInterrupt:
            _safe_print("\n[*] Auto-sync watcher stopped.")
            break
        except Exception as err:
            time.sleep(0.5)


if __name__ == "__main__":
    results = sync_all_vendors()
    print("Manual sync complete:")
    for v, count in results.items():
        if count > 0:
            print(f" - {v}: {count} trips")
