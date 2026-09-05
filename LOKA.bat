@echo off
setlocal
cd /d C:\LOKA
set PY=C:\Progra~1\Python312\python.exe

:menu
cls
echo ================================================================
echo    LOKA BOOKKEEPING
echo ================================================================
echo.
echo   --- EVERY DAY ------------------------------------------------
echo    1.  Enter today's numbers  (opens today.txt)
echo    2.  PREVIEW  what will be recorded
echo    3.  RECORD IT  (writes to the books)
echo    4.  ADD MORE to a day already entered
echo.
echo   --- MONEY FROM / BACK TO CAPITAL -----------------------------
echo    5.  Expense PAID WITH CAPITAL MONEY  (rent, electricity...)
echo    6.  Money PAID BACK to Capital
echo    7.  Settle the owner ledger with me
echo.
echo   --- LOOK AT THINGS -------------------------------------------
echo    8.  Open the dashboard
echo    9.  Open the workbook (Excel)
echo   10.  Report - week / month / all time
echo   11.  Health check
echo.
echo   --- NOW AND THEN ---------------------------------------------
echo   12.  Record propinas (tips to staff)
echo   13.  Cash count - re-anchor to what you counted
echo   14.  Scan receipt photos  (reads them on this PC)
echo   15.  Open the receipts drop folder
echo.
echo   --- FIXING / SETUP -------------------------------------------
echo   16.  Teach a category  (fix anything it didn't recognise)
echo   17.  Show what I've taught it
echo   18.  Refresh dashboard only
echo   19.  UNDO - restore an earlier backup
echo   20.  Add Spanish to the receipt reader (one time, needs admin)
echo.
echo    0.  Exit
echo.
set /p c="Choose: "

if "%c%"=="1"  goto edit
if "%c%"=="2"  goto preview
if "%c%"=="3"  goto apply
if "%c%"=="4"  goto addmore
if "%c%"=="5"  %PY% tools.py capex   & pause & goto menu
if "%c%"=="6"  %PY% tools.py repay   & pause & goto menu
if "%c%"=="7"  %PY% tools.py settle  & pause & goto menu
if "%c%"=="8"  start "" "C:\LOKA\dashboard.html" & goto menu
if "%c%"=="9"  start "" "C:\LOKA\LOKA_Restaurant_Manager.xlsx" & goto menu
if "%c%"=="10" %PY% tools.py report  & pause & goto menu
if "%c%"=="11" goto doctor
if "%c%"=="12" %PY% tools.py tips    & pause & goto menu
if "%c%"=="13" %PY% tools.py recount & pause & goto menu
if "%c%"=="14" %PY% scan.py          & pause & goto menu
if "%c%"=="15" start "" "C:\LOKA\Bills\_DROP_BILLS_HERE" & goto menu
if "%c%"=="16" goto teach
if "%c%"=="17" %PY% teach.py --list  & pause & goto menu
if "%c%"=="18" %PY% loka.py refresh-all & pause & goto menu
if "%c%"=="19" %PY% tools.py restore & pause & goto menu
if "%c%"=="20" powershell -NoProfile -ExecutionPolicy Bypass -File "C:\LOKA\add_spanish_ocr.ps1" & goto menu
if "%c%"=="0"  exit
goto menu

:teach
%PY% teach.py
echo.
echo   (choose 'A' below to add another word by hand, or press Enter to go back)
set /p more="   A = add another, Enter = back: "
if /i "%more%"=="A" %PY% teach.py --add & pause
goto menu

:addmore
cls
echo ================================================================
echo    ADD MORE TO A DAY ALREADY ENTERED
echo ================================================================
echo.
echo   Use this when you already recorded a day and then took
echo   MORE money or spent MORE later the same day.
echo.
echo   Put ONLY THE EXTRA amounts in the file - not the whole
echo   day again. They get ADDED to what is already there.
echo.
echo   Make sure the DATE line is the day you are topping up.
echo.
pause
if not exist extra.txt copy template_day.txt extra.txt >nul
notepad extra.txt
echo.
%PY% eod.py extra.txt --add
echo.
echo ================================================================
set /p y="Add this to the day? (y/n): "
if /i not "%y%"=="y" goto menu
%PY% eod.py extra.txt --add --apply
echo.
%PY% doctor.py
if not exist days mkdir days
copy extra.txt "days\extra_%date:~-4%%date:~4,2%%date:~7,2%_%random%.txt" >nul
del extra.txt
echo.
echo Done.  Now open VS Code and PUSH.
pause
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
echo ================================================================
set /p y="Record this to the books? (y/n): "
if /i not "%y%"=="y" goto menu
%PY% eod.py today.txt --apply
echo.
%PY% doctor.py
echo.
echo Archiving today.txt...
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
