@echo off
cd /d "%~dp0"
echo ==============================================
echo   RFQ App - Re-run an email-list (table) RFQ
echo ==============================================
set /p TQ="Enter TQ/TRFQ number (e.g. TQ26175171): "
set /p TO="Reply-to address (leave blank for Reply-All): "
echo.
if "%TO%"=="" goto replyall
python retry_email_rfq.py %TQ% %TO%
goto done
:replyall
python retry_email_rfq.py %TQ%
:done
echo.
pause
