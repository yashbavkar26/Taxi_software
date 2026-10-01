"""Unit tests for the CLI entry point."""

from pathlib import Path
from app.cli.import_bookings import main


class TestCLI:
    def test_cli_dry_run_sample_file(self):
        sample_file = Path("Requirements/Input/Bookings Sample.xlsx")
        exit_code = main(["--dry-run", str(sample_file)])
        assert exit_code == 0

    def test_cli_file_not_found(self):
        exit_code = main(["this_file_does_not_exist.xlsx"])
        assert exit_code == 1

    def test_cli_help(self, capsys):
        try:
            main(["--help"])
        except SystemExit as e:
            assert e.code == 0
        captured = capsys.readouterr()
        assert "Excel-to-MySQL" in captured.out

    def test_cli_dry_run_with_no_auto_register_drivers(self):
        sample_file = Path("Requirements/Input/Bookings Sample.xlsx")
        exit_code = main(["--dry-run", "--no-auto-register-drivers", str(sample_file)])
        assert exit_code == 0

