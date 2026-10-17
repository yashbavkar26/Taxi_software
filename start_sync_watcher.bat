@echo off
title Taxi Software - Vendor Masterfile Live Sync
color 0A
echo ===============================================================
echo     TAXI SOFTWARE - LIVE VENDOR MASTERFILE SYNCHRONIZER
echo ===============================================================
echo.
echo Monitoring: masterfile_template.xlsx
echo When you make entries and save the masterfile, all vendor files
echo (Arvind.xlsx, Bharat.xlsx, etc.) will update automatically!
echo.
echo Press Ctrl+C in this window anytime to stop the synchronizer.
echo ===============================================================
echo.

python -m app.cli.sync_vendors --watch
pause
