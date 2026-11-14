@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Process a Link (email had no weblink)
echo ==============================================
set /p LINK="Paste the RFQ link (e.g. https://sinwa.epbrowser.com/portal-ext/rfq/...): "
echo.
python main.py --link "%LINK%"
echo.
pause
