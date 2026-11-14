@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Scheduler (checks inbox every 15 min)
echo ==============================================
echo Keep this window open. Press Ctrl+C to stop.
echo.
python scheduler.py
echo.
pause
