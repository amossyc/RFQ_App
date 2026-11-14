@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Redo an RFQ (after pricelist update)
echo ==============================================
set /p RFQ="Enter TRFQ number(s), comma-separated (e.g. TRFQ26415581 or TRFQ26415581,TRFQ26415582): "
echo.
python main.py --force %RFQ% --no-reply
echo.
pause
