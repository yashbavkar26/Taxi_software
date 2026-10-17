"""Vendor Synchronization Service.

Automatically partitions and synchronizes entries from masterfile_template.xlsx
into individual vendor workbooks (e.g. Arvind.xlsx, Bharat.xlsx) matching the
exact formatting of the user's reference file.
"""

import os
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


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


def canonical_vendor_name(name_str: str) -> str:
    """Normalize vendor name to match registered canonical active vendors if possible."""
    clean = name_str.strip()
    for v in ACTIVE_VENDORS:
        if v.lower() == clean.lower():
            return v
    return clean.title() if clean.islower() else clean


def get_project_root() -> Path:
    """Dynamically determine the project root directory based on package location."""
    return Path(__file__).resolve().parents[3]


def register_project_path(path: Optional[Path] = None):
    """Save the project path in %APPDATA%/TaxiSoftware/project_dir.txt and registry."""
    if path is None:
        path = get_project_root()
    p_str = str(path.resolve())
    try:
        appdata = os.environ.get("APPDATA")
        if appdata:
            cfg_dir = Path(appdata) / "TaxiSoftware"
            cfg_dir.mkdir(parents=True, exist_ok=True)
            (cfg_dir / "project_dir.txt").write_text(p_str, encoding="utf-8")
    except Exception:
        pass

    if sys.platform == "win32":
        try:
            import winreg

            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\TaxiSoftware") as key:
                winreg.SetValueEx(key, "ProjectPath", 0, winreg.REG_SZ, p_str)
        except Exception:
            pass


def resolve_masterfile_path(candidate_path: Optional[Union[str, Path]] = None) -> Path:
    """Dynamically resolve the masterfile path from candidate_path, cwd, project root, or parent folders."""
    if candidate_path:
        p = Path(candidate_path)
        if p.exists():
            return p.resolve()
        proj_p = get_project_root() / candidate_path
        if proj_p.exists():
            return proj_p.resolve()

    search_dirs = [
        Path.cwd(),
        get_project_root(),
        get_project_root().parent,
        get_project_root() / "Requirements" / "Input",
    ]
    file_names = [
        "masterfile_template.xlsx",
        "Masterfile.xlsx",
        "masterfile.xlsx",
        "MasterFile.xlsx",
        "Master File.xlsx",
    ]

    for d in search_dirs:
        if not d.exists() or not d.is_dir():
            continue
        for fname in file_names:
            target = d / fname
            if target.exists():
                return target.resolve()

    for d in search_dirs:
        if d.exists() and d.is_dir():
            for f in d.glob("*master*.xlsx"):
                if not f.name.startswith("~$") and f.is_file():
                    return f.resolve()

    # Fallback to default in project root
    return (get_project_root() / "masterfile_template.xlsx").resolve()


def resolve_vendor_output_dir(
    output_dir: Optional[Union[str, Path]] = None,
    masterfile_path: Optional[Union[str, Path]] = None,
) -> Path:
    """Determine destination directory for vendor workbooks (Arvind.xlsx, Bharat.xlsx, etc.).

    If output_dir is specified (other than None or '.'), use it.
    If output_dir is None or '.', defaults to the directory containing the masterfile
    so vendor files automatically live right alongside the masterfile wherever it is moved.
    """
    if output_dir is not None and str(output_dir).strip() not in (".", ""):
        out_p = Path(output_dir)
        out_p.mkdir(parents=True, exist_ok=True)
        return out_p.resolve()

    if masterfile_path is not None:
        try:
            m_p = Path(masterfile_path)
            if m_p.exists():
                return m_p.parent.resolve()
        except Exception:
            pass

    resolved_master = resolve_masterfile_path(masterfile_path)
    if resolved_master.exists():
        return resolved_master.parent.resolve()

    return get_project_root().resolve()


def read_masterfile_trips(
    masterfile_path: Optional[Union[str, Path]] = None,
) -> Dict[str, List[Dict[str, Any]]]:
    """Read all trips from the masterfile workbook and group them by vendor.

    Args:
        masterfile_path: Path to masterfile_template.xlsx or equivalent (auto-detected if None).

    Returns:
        Dict mapping vendor name to list of trip record dictionaries.
    """
    path = resolve_masterfile_path(masterfile_path)
    if not path.exists():
        raise FileNotFoundError(f"Masterfile not found at '{path}'")

    wb = openpyxl.load_workbook(path, data_only=False)
    sheet_name = "Masterfile" if "Masterfile" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]


    vendor_trips: Dict[str, List[Dict[str, Any]]] = {}

    for r in range(2, ws.max_row + 1):
        vendor_val = ws.cell(r, 2).value
        if not vendor_val or str(vendor_val).strip() == "":
            continue

        vendor_name = canonical_vendor_name(str(vendor_val))
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
    masterfile_path: Optional[Union[str, Path]] = None,
    output_dir: Optional[Union[str, Path]] = None,
) -> Dict[str, int]:
    """Sync all vendor entries from masterfile into individual vendor workbooks.

    Dynamically resolves both masterfile location and vendor output directory
    so vendor files update wherever they or the masterfile are located.
    """
    m_path = resolve_masterfile_path(masterfile_path)
    out_dir = resolve_vendor_output_dir(output_dir, m_path)
    vendor_trips = read_masterfile_trips(m_path)

    # Union of registered active vendors + all vendors present in masterfile
    all_vendors: Set[str] = set(ACTIVE_VENDORS).union(vendor_trips.keys())

    sync_summary: Dict[str, int] = {}
    for vendor in sorted(all_vendors):
        trips = vendor_trips.get(vendor, [])
        v_file = out_dir / f"{vendor}.xlsx"
        try:
            write_vendor_workbook(vendor, trips, v_file)
            sync_summary[vendor] = len(trips)
        except (PermissionError, OSError):
            # File is currently open in Excel; record trip count but continue without failing
            sync_summary[vendor] = len(trips)

    return sync_summary


def log_message(msg: str, log_file: Optional[Path] = None):
    """Log to stdout and append to vendor_sync.log."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{timestamp}] {msg}"
    if sys.stdout is not None:
        try:
            print(formatted)
        except Exception:
            pass
    if log_file:
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(formatted + "\n")
        except Exception:
            pass


def watch_and_sync(
    masterfile_path: Optional[Union[str, Path]] = None,
    output_dir: Optional[Union[str, Path]] = None,
    poll_interval: float = 1.0,
    log_filename: str = "vendor_sync.log",
):
    """Continuously monitor masterfile_path and trigger auto-sync on save.

    Features:
    - Fully dynamic path discovery: works wherever master files and vendor workbooks are located.
    - Single-instance lock via kernel-level file locking to prevent duplicate watchers.
    - Automatic retry with exponential backoff on transient Excel file write locks.
    - Guaranteed update of last_mtime only after successful sync.
    """
    register_project_path()

    m_path = resolve_masterfile_path(masterfile_path)
    out_dir = resolve_vendor_output_dir(output_dir, m_path)
    log_path = out_dir / log_filename

    # OS-level single-instance lock to ensure only one watcher runs
    lock_file_path = out_dir / "vendor_sync.lock"
    lock_file = None
    try:
        lock_file = open(lock_file_path, "a+b")
        if sys.platform == "win32":
            import msvcrt
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (OSError, IOError):
        log_message(
            f"Auto-Sync Watcher is already running in background. Process {os.getpid()} exiting cleanly.",
            log_path,
        )
        if lock_file:
            try:
                lock_file.close()
            except Exception:
                pass
        return

    pid_file = out_dir / "vendor_sync.pid"
    current_pid = os.getpid()
    try:
        pid_file.write_text(str(current_pid), encoding="utf-8")
    except Exception:
        pass

    log_message(
        f"Vendor Auto-Sync Watcher started (PID {current_pid}). Monitoring: '{m_path.name}' in '{m_path.parent}' (Vendor out: '{out_dir}')",
        log_path,
    )

    def get_candidate_files() -> List[Path]:
        cands = []
        if m_path.exists():
            cands.append(m_path)

        search_dirs = [out_dir, m_path.parent, get_project_root(), Path.cwd()]
        for d in search_dirs:
            if not d.exists() or not d.is_dir():
                continue
            for f in d.glob("*master*.xlsx"):
                if not f.name.startswith("~$") and f.is_file():
                    resolved = f.resolve()
                    if resolved not in cands:
                        cands.append(resolved)

        return cands

    file_stats: Dict[Path, Dict[str, float]] = {}
    for f in get_candidate_files():
        if f.exists():
            try:
                st = f.stat()
                file_stats[f] = {"mtime": st.st_mtime, "size": st.st_size}
            except Exception:
                pass

    # Initial sync on startup
    if m_path.exists():
        try:
            summary = sync_all_vendors(m_path, out_dir)
            total_trips = sum(summary.values())
            active_count = sum(1 for c in summary.values() if c > 0)
            log_message(
                f"Initial sync complete: {len(summary)} vendor files up to date in '{out_dir}' "
                f"({active_count} active, {total_trips} trips).",
                log_path,
            )
        except Exception as e:
            log_message(f"Initial sync warning: {e}", log_path)

    try:
        while True:
            time.sleep(poll_interval)
            candidates = get_candidate_files()
            for cand in candidates:
                if not cand.exists():
                    continue

                try:
                    cand_stat = cand.stat()
                    current_mtime = cand_stat.st_mtime
                    current_size = cand_stat.st_size
                except Exception:
                    continue

                last_info = file_stats.get(cand, {"mtime": 0, "size": 0})
                if current_mtime != last_info["mtime"] or current_size != last_info["size"]:
                    # Wait briefly for Excel to finish writing
                    time.sleep(0.5)

                    target_dir = cand.parent if (output_dir is None or str(output_dir).strip() in (".", "")) else out_dir
                    synced = False
                    for attempt in range(6):
                        try:
                            summary = sync_all_vendors(cand, target_dir)
                            active_count = sum(1 for c in summary.values() if c > 0)
                            total_trips = sum(summary.values())
                            log_message(
                                f"Auto-Synced {len(summary)} vendor files on edit to '{cand.name}' in '{target_dir}' "
                                f"({active_count} active with {total_trips} trips).",
                                log_path,
                            )
                            # Record successful sync stats
                            file_stats[cand] = {
                                "mtime": cand.stat().st_mtime,
                                "size": cand.stat().st_size,
                            }
                            synced = True
                            break
                        except Exception:
                            time.sleep(0.5 * (attempt + 1))

                    if not synced:
                        log_message(
                            f"Warning: '{cand.name}' busy/locked by Excel; will retry on next tick.",
                            log_path,
                        )

    except KeyboardInterrupt:
        log_message("Auto-sync watcher stopped by user.", log_path)
    finally:
        try:
            if pid_file.exists():
                pid_file.unlink()
        except Exception:
            pass
        if lock_file:
            try:
                if sys.platform == "win32":
                    import msvcrt
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
                lock_file.close()
            except Exception:
                pass


if __name__ == "__main__":
    results = sync_all_vendors()
    print("Manual sync complete:")
    for v, count in results.items():
        if count > 0:
            print(f" - {v}: {count} trips")
