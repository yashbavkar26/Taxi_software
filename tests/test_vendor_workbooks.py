"""Automated tests for vendor synchronization and vendor workbooks."""

from pathlib import Path
import openpyxl
import pytest

from app.modules.vendors.sync_service import (
    ACTIVE_VENDORS,
    MASTERFILE_HEADERS,
    read_masterfile_trips,
    sync_all_vendors,
    write_vendor_workbook,
)
from app.modules.vendors.vendor_generator import (
    create_vendor_workbook,
    generate_all_vendor_workbooks,
)
from app.cli.sync_vendors import main as cli_sync_main
from app.cli.generate_vendor_files import main as cli_generate_main


class TestVendorSyncAndWorkbooks:
    @pytest.fixture(autouse=True)
    def ensure_workbooks_synced(self):
        """Ensure all vendor workbooks are synchronized before running tests."""
        generate_all_vendor_workbooks(output_dir=".", masterfile_filename="masterfile_template.xlsx")
        sync_all_vendors(masterfile_path="masterfile_template.xlsx", output_dir=".")

    def test_vendor_template_exists_and_valid(self):
        """Verify vendor_template.xlsx exists at root with valid formulas."""
        tpl_path = Path("vendor_template.xlsx")
        assert tpl_path.exists(), "vendor_template.xlsx must exist at root"
        wb = openpyxl.load_workbook(tpl_path, data_only=False)
        ws = wb["Vendor_Trips"]

        assert ws.cell(1, 1).value == "VENDOR TRIP LOG & BILLING SUMMARY"
        assert "FILTER(" in str(ws.cell(7, 1).value)
        wb.close()

    def test_all_14_vendor_workbooks_exist(self):
        """Verify all 14 active vendor files exist in project root."""
        for vendor in ACTIVE_VENDORS:
            v_path = Path(f"{vendor}.xlsx")
            assert v_path.exists(), f"Workbook for vendor '{vendor}' must exist at root"
            assert v_path.stat().st_size > 0

    def test_arvind_workbook_formatting_and_data(self):
        """Verify Arvind.xlsx matches user's exact required formatting and synced rows."""
        wb = openpyxl.load_workbook("Arvind.xlsx", data_only=False)
        ws = wb["Vendor_Trips"]

        # Header in row 1
        headers = [ws.cell(1, c).value for c in range(1, 21)]
        assert headers == MASTERFILE_HEADERS

        # Header styling
        assert ws.row_dimensions[1].height >= 28
        assert ws.freeze_panes == "A2"

        # Check Arvind has matching trips synced from masterfile_template.xlsx
        arvind_trips = read_masterfile_trips("masterfile_template.xlsx").get("Arvind", [])
        assert ws.max_row == len(arvind_trips) + 1

        # Verify Row 2 values and formulas
        assert ws.cell(2, 1).value == 1
        assert ws.cell(2, 2).value == "Arvind"
        assert ws.cell(2, 3).value == "01/12/2025"
        assert ws.cell(2, 4).value == "GA 07 T 2608"

        # Formula in Col H (Total Km)
        assert ws.cell(2, 8).value == '=IF(OR(ISBLANK(G2),ISBLANK(F2)),"",G2-F2)'
        # Formula in Col L (Amt)
        assert "IF(" in str(ws.cell(2, 12).value)
        # Formula in Col N (Total Amt)
        assert ws.cell(2, 14).value == '=IF(AND(ISBLANK(L2),ISBLANK(K2),ISBLANK(I2)),"",SUM(L2,K2,I2))'

        # Row 6: Trip 5
        assert ws.cell(6, 1).value == 5
        assert ws.cell(6, 2).value == "Arvind"
        assert ws.cell(6, 3).value == "17/10/2026"
        assert ws.cell(6, 4).value == "GA 05 T 8122"

        wb.close()

    @pytest.mark.parametrize("vendor_name", ["Bharat", "Gajanan", "Harshal", "Lalu", "Dilip", "Yogesh"])
    def test_active_vendors_synced_trips(self, vendor_name: str):
        """Verify other active vendors have their matching trips synced."""
        wb = openpyxl.load_workbook(f"{vendor_name}.xlsx", data_only=False)
        ws = wb["Vendor_Trips"]

        headers = [ws.cell(1, c).value for c in range(1, 21)]
        assert headers == MASTERFILE_HEADERS
        assert ws.max_row >= 2  # At least 1 trip synced for these active vendors

        for r in range(2, ws.max_row + 1):
            assert ws.cell(r, 2).value == vendor_name
            # Formulas preserved
            assert "G" in str(ws.cell(r, 8).value)  # Total Km formula
            assert "SUM(" in str(ws.cell(r, 14).value)  # Total Amt formula

        wb.close()

    def test_sync_auto_creates_new_vendor(self, tmp_path: Path):
        """Verify synchronization automatically generates a new file when a new vendor appears."""
        test_master = tmp_path / "test_master.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Masterfile"
        ws.append(MASTERFILE_HEADERS)
        ws.append([1, "Zaid Logistics", "18/10/2026", "GA 07 T 9999", "Sedan", 1000, 1100, "", 0, 20, 0, "", "Guest", "", False, "10:00 AM", "2:00 PM", "Driver", "Panjim", None])
        wb.save(test_master)
        wb.close()

        summary = sync_all_vendors(masterfile_path=test_master, output_dir=tmp_path)
        assert "Zaid Logistics" in summary
        assert summary["Zaid Logistics"] == 1

        new_v_file = tmp_path / "Zaid Logistics.xlsx"
        assert new_v_file.exists()
        nwb = openpyxl.load_workbook(new_v_file, data_only=False)
        nws = nwb["Vendor_Trips"]
        assert nws.cell(2, 2).value == "Zaid Logistics"
        assert nws.cell(2, 4).value == "GA 07 T 9999"
        nwb.close()

    def test_cli_sync_vendors(self, tmp_path: Path):
        """Verify sync_vendors CLI executes cleanly."""
        ret = cli_sync_main(["--masterfile", "masterfile_template.xlsx", "--out-dir", str(tmp_path)])
        assert ret == 0
        assert (tmp_path / "Arvind.xlsx").exists()
        assert (tmp_path / "Bharat.xlsx").exists()
