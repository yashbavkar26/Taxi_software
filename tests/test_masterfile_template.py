"""Tests for masterfile schema and Excel template generator."""

from pathlib import Path
import openpyxl
import pytest

from app.modules.vendors.schema import (
    MASTERFILE_COLUMNS,
    MASTERFILE_GENERATED_COLUMNS,
    MASTERFILE_HEADER_ALIASES,
    MASTERFILE_REQUIRED_HEADERS,
    MASTERFILE_SCHEMA,
    is_masterfile_sheet,
)
from app.modules.vendors.template import generate_masterfile_template


class TestMasterfileTemplate:
    @pytest.fixture
    def template_wb(self, tmp_path):
        file_path = tmp_path / "test_masterfile_template.xlsx"
        generate_masterfile_template(output_path=file_path)
        wb = openpyxl.load_workbook(file_path, data_only=False)
        return wb

    def test_template_files_exist_in_project(self):
        """Verify the template exists at root and in Requirements/Input."""
        root_path = Path("masterfile_template.xlsx")
        req_path = Path("Requirements/Input/masterfile_template.xlsx")
        if not root_path.exists() or not req_path.exists():
            generate_masterfile_template(output_path=root_path, additional_paths=[req_path])
        assert root_path.exists(), "masterfile_template.xlsx should exist at root"
        assert req_path.exists(), "masterfile_template.xlsx should exist in Requirements/Input"
        assert root_path.stat().st_size > 0
        assert req_path.stat().st_size > 0

    def test_required_sheets_exist(self, template_wb):
        """Verify all essential and reference sheets exist in the workbook."""
        expected_sheets = [
            "Masterfile",
            "Sample",
            "Vendors",
            "Drivers",
            "Vehicles",
            "Rate_Card",
            "Annexure_Sample",
        ]
        for sheet in expected_sheets:
            assert sheet in template_wb.sheetnames, f"Sheet '{sheet}' must be present"

    def test_headers_and_db_column_alignment(self, template_wb):
        """Verify that Masterfile headers match taxi_schema.sql and MASTERFILE_SCHEMA."""
        ws = template_wb["Masterfile"]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert len(headers) == 20

        # Verify mapping through MASTERFILE_HEADER_ALIASES
        mapped_db_cols = []
        for h in headers:
            clean = str(h).strip().lower()
            assert clean in MASTERFILE_HEADER_ALIASES, f"Header '{h}' must be defined in aliases"
            db_col = MASTERFILE_HEADER_ALIASES[clean]
            assert db_col in MASTERFILE_COLUMNS, f"DB Column '{db_col}' must exist in MASTERFILE_COLUMNS"
            mapped_db_cols.append(db_col)

        # Expected DB columns from taxi_schema.sql masterfile table
        expected_db_cols = [
            "sl_no",
            "vendor",
            "travel_date",
            "vehicle_no",
            "vehicle_type",
            "starting_km",
            "ending_km",
            "total_km",
            "night",
            "rate",
            "toll",
            "amt",
            "guest_name",
            "total_amt",
            "payment_details",
            "in_time",
            "out_time",
            "driver_name",
            "places",
            "payment_date",
        ]
        assert mapped_db_cols == expected_db_cols

    def test_sample_sheet_identical_headers(self, template_wb):
        """Verify Sample sheet has identical headers and layout as Masterfile."""
        ws_master = template_wb["Masterfile"]
        ws_sample = template_wb["Sample"]
        master_headers = [ws_master.cell(1, c).value for c in range(1, ws_master.max_column + 1)]
        sample_headers = [ws_sample.cell(1, c).value for c in range(1, ws_sample.max_column + 1)]
        assert master_headers == sample_headers

    def test_formulas_in_sample_and_masterfile(self, template_wb):
        """Verify Total Km, Amt, and Total Amt formulas are present."""
        ws = template_wb["Sample"]
        for r in range(2, 6):
            # Total Km in column H (8)
            f_total_km = ws.cell(r, 8).value
            assert str(f_total_km).startswith("="), f"Row {r} Total Km must be formula"
            assert f"G{r}-F{r}" in str(f_total_km)

            # Amt in column L (12)
            f_amt = ws.cell(r, 12).value
            assert str(f_amt).startswith("="), f"Row {r} Amt must be formula"
            assert f"H{r}" in str(f_amt) and f"J{r}" in str(f_amt)

            # Total Amt in column N (14)
            f_total_amt = ws.cell(r, 14).value
            assert str(f_total_amt).startswith("="), f"Row {r} Total Amt must be formula"
            assert f"SUM(L{r},K{r},I{r})" in str(f_total_amt)

    def test_data_validations_configured(self, template_wb):
        """Verify Excel data validation dropdowns are applied."""
        ws = template_wb["Masterfile"]
        dv_list = ws.data_validations.dataValidation
        assert len(dv_list) >= 4, "Data validations for Vendor, Vehicle Type, Rate, Payment, Driver required"

        formulas = [dv.formula1 for dv in dv_list]
        assert any("Vendors" in str(f) for f in formulas)
        assert any("Drivers" in str(f) for f in formulas)
        assert any("TRUE,FALSE" in str(f) for f in formulas)
        assert any("18,20,22" in str(f) for f in formulas)

    def test_reference_sheets_data(self, template_wb):
        """Verify Vendors, Drivers, Vehicles, and Rate_Card contain populated records."""
        # Vendors sheet
        ws_v = template_wb["Vendors"]
        assert ws_v.max_row >= 15
        vendors = [ws_v.cell(r, 2).value for r in range(2, ws_v.max_row + 1) if ws_v.cell(r, 2).value]
        assert "Arvind" in vendors
        assert "Gajanan" in vendors
        assert "Harshal" in vendors
        assert "Lalu" in vendors

        # Drivers sheet
        ws_d = template_wb["Drivers"]
        assert ws_d.max_row >= 26
        drivers = [ws_d.cell(r, 2).value for r in range(2, ws_d.max_row + 1) if ws_d.cell(r, 2).value]
        assert "Abdul" in drivers
        assert "Rahul" in drivers
        assert "Suyog" in drivers

        # Vehicles sheet
        ws_veh = template_wb["Vehicles"]
        assert ws_veh.max_row >= 20
        vech_nos = [ws_veh.cell(r, 2).value for r in range(2, ws_veh.max_row + 1) if ws_veh.cell(r, 2).value]
        assert "GA 07 T 2608" in vech_nos
        assert "GA 06 T 4317" in vech_nos

        # Rate Card sheet
        ws_rc = template_wb["Rate_Card"]
        assert ws_rc.max_row >= 8
        rates = [ws_rc.cell(r, 1).value for r in range(2, ws_rc.max_row + 1) if ws_rc.cell(r, 1).value]
        assert 18 in rates
        assert 20 in rates
        assert 22 in rates
        assert 35 in rates

    def test_annexure_sample_matches_sql_view(self, template_wb):
        """Verify Annexure_Sample matches the 16 columns of view_annexure_arvind in taxi_schema.sql."""
        ws_annex = template_wb["Annexure_Sample"]
        headers = [ws_annex.cell(1, c).value for c in range(1, ws_annex.max_column + 1)]
        expected = [
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
        assert headers == expected
        assert ws_annex.max_row >= 5  # Has sample rows

    def test_reference_masterfile_headers_alias_compatibility(self):
        """Verify all original headers in Requirements/Input/Masterfile.xlsx map correctly."""
        ref_path = Path("Requirements/Input/Masterfile.xlsx")
        assert ref_path.exists()
        wb_ref = openpyxl.load_workbook(ref_path, read_only=True)
        ws_ref = wb_ref.active
        raw_headers = [cell.value for cell in next(ws_ref.iter_rows(max_row=1))]
        wb_ref.close()

        recognized = []
        for h in raw_headers:
            if h is None:
                continue
            clean = str(h).strip().lower()
            assert clean in MASTERFILE_HEADER_ALIASES, f"Original header '{h}' should be in aliases"
            recognized.append(MASTERFILE_HEADER_ALIASES[clean])

        # Essential columns must all be recognized
        assert "vendor" in recognized
        assert "travel_date" in recognized
        assert "vehicle_no" in recognized
        assert "vehicle_type" in recognized
        assert "starting_km" in recognized
        assert "ending_km" in recognized
        assert "total_km" in recognized
        assert "amt" in recognized
        assert "total_amt" in recognized
        assert "places" in recognized

    def test_is_masterfile_sheet_helper(self):
        """Test sheet name detection helper."""
        assert is_masterfile_sheet("Masterfile") is True
        assert is_masterfile_sheet("master file") is True
        assert is_masterfile_sheet("Vendor Book") is True
        assert is_masterfile_sheet("Sheet1") is True
        assert is_masterfile_sheet("Contacts") is False
        assert is_masterfile_sheet("Feb 01") is False
