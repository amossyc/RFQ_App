@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Process a Single RFQ
echo ==============================================
echo   (processes ONLY this RFQ, sends the reply email,
echo    and marks it complete so it is not re-run later)
echo.
set /p RFQ="Enter TRFQ number (e.g. TRFQ26417322): "
echo.
python main.py --force %RFQ%
echo.
pause
