@echo off
REM =====================================================
REM  PICK Equity Derivatives Tracker — Task Scheduler Setup
REM  Run this as Administrator to create scheduled tasks
REM =====================================================

SET EXCEL_PATH="C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
SET WORKBOOK_PATH="C:\Users\%USERNAME%\Documents\EquityDerivTracker.xlsm"

REM -- Create task: Auto-scan every weekday at market open (9 AM) --
schtasks /create /tn "EquityDerivTracker_AutoScan" ^
  /tr "%EXCEL_PATH% %WORKBOOK_PATH% /autoscan" ^
  /sc weekly /d MON,TUE,WED,THU,FRI ^
  /st 08:55 ^
  /f

REM -- Create task: Close workbook at market close (5:30 PM) --
schtasks /create /tn "EquityDerivTracker_Close" ^
  /tr "taskkill /im excel.exe /f" ^
  /sc weekly /d MON,TUE,WED,THU,FRI ^
  /st 17:30 ^
  /f

echo.
echo Tasks created:
echo   - EquityDerivTracker_AutoScan (Mon-Fri 08:55)
echo   - EquityDerivTracker_Close (Mon-Fri 17:30)
echo.
echo IMPORTANT: Update EXCEL_PATH and WORKBOOK_PATH in this script.
echo.
pause
