@echo off
setlocal
cd /d C:\LOKA
set PY=C:\Progra~1\Python312\python.exe

:menu
cls
echo ============================================
echo    LOKA BOOKKEEPING
echo ============================================
echo.
echo   1.  Enter today's numbers  (opens today.txt)
echo   2.  PREVIEW  what will be recorded
echo   3.  RECORD IT  (writes to the books)
echo.
echo   4.  Open the dashboard
echo   5.  Open the workbook (Excel)
echo   6.  Health check
echo   7.  Refresh dashboard only
echo.
echo   8.  Teach a category  (fix anything it didn't recognise)
echo   9.  Show what I've taught it
echo.
echo  10.  Report - week / month / all time
echo  11.  Record propinas (tips to staff)
echo  12.  Cash count - re-anchor to what you counted
echo  13.  Money paid back to Capital
echo  19.  Expense PAID FROM CAPITAL (rent, electricity...)
echo  14.  Settle the owner ledger with me
echo  15.  UNDO - restore an earlier backup
echo.
echo  16.  SCAN RECEIPT PHOTOS  (reads them on this PC)
echo  17.  Open the receipts drop folder
echo  18.  Add Spanish to the receipt reader (one time, needs admin)
echo.
echo   0.  Exit
echo.
set /p c="Choose: "

if "%c%"=="1" goto edit
if "%c%"=="2" goto preview
if "%c%"=="3" goto apply
if "%c%"=="4" start "" "C:\LOKA\dashboard.html" & goto menu
if "%c%"=="5" start "" "C:\LOKA\LOKA_Restaurant_Manager.xlsx" & goto menu
if "%c%"=="6" goto doctor
if "%c%"=="7" %PY% loka.py refresh-all & pause & goto menu
if "%c%"=="8" goto teach
if "%c%"=="9" %PY% teach.py --list & pause & goto menu
if "%c%"=="10" %PY% tools.py report & pause & goto menu
if "%c%"=="11" %PY% tools.py tips & pause & goto menu
if "%c%"=="12" %PY% tools.py recount & pause & goto menu
if "%c%"=="13" %PY% tools.py repay & pause & goto menu
if "%c%"=="14" %PY% tools.py settle & pause & goto menu
if "%c%"=="15" %PY% tools.py restore & pause & goto menu
if "%c%"=="16" %PY% scan.py & pause & goto menu
if "%c%"=="17" start "" "C:\LOKA\Bills\_DROP_BILLS_HERE" & goto menu
if "%c%"=="18" powershell -NoProfile -ExecutionPolicy Bypass -File "C:\LOKA\add_spanish_ocr.ps1" & goto menu
if "%c%"=="19" %PY% tools.py capex & pause & goto menu
if "%c%"=="0" exit
goto menu

:teach
%PY% teach.py
echo.
echo   (choose 'A' below to add another word by hand, or press Enter to go back)
set /p more="   A = add another, Enter = back: "
if /i "%more%"=="A" %PY% teach.py --add & pause
goto menu

:edit
if not exist today.txt copy template_day.txt today.txt >nul
notepad today.txt
goto menu

:preview
if not exist today.txt echo No today.txt yet - choose 1 first. & pause & goto menu
%PY% eod.py today.txt
echo.
pause
goto menu

:apply
if not exist today.txt echo No today.txt yet - choose 1 first. & pause & goto menu
%PY% eod.py today.txt
echo.
echo ============================================
set /p y="Record this to the books? (y/n): "
if /i not "%y%"=="y" goto menu
%PY% eod.py today.txt --apply
echo.
%PY% doctor.py
echo.
echo Archiving today.txt...
for /f "tokens=2 delims=:" %%a in ('findstr /i "^DATE" today.txt') do set d=%%a
if not exist days mkdir days
copy today.txt "days\day_%date:~-4%%date:~4,2%%date:~7,2%.txt" >nul
del today.txt
echo Done.  Now open VS Code and PUSH.
pause
goto menu

:doctor
%PY% doctor.py
echo.
pause
goto menu
