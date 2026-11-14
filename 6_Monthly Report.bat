@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Monthly Report
echo ==============================================
python tracker.py report
echo.
pause
