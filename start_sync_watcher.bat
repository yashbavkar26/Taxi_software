@echo off
cd /d "%~dp0"
title Taxi Software - Vendor Masterfile Live Sync
color 0A
echo ===============================================================
echo     TAXI SOFTWARE - LIVE VENDOR MASTERFILE SYNCHRONIZER
echo ===============================================================
echo.
echo Directory: %CD%
echo Monitoring masterfile entries and auto-syncing all vendor files...
echo.
echo Press Ctrl+C in this window anytime to stop the synchronizer.
echo ===============================================================
echo.

python -m app.cli.sync_vendors --watch
pause

