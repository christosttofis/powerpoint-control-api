@echo off
setlocal
cd /d "%~dp0"
title PowerPoint Control API - Setup

echo ============================================================
echo   PowerPoint Control API - one-time setup
echo ============================================================
echo.

REM ---- 1. Find Python 3.10+ ------------------------------------------------
set "PY="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if defined PY goto have_python
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=python"
if defined PY goto have_python

echo [!] Python 3.10 or newer was not found.
where winget >nul 2>&1
if errorlevel 1 goto no_winget
echo     Installing Python 3.12 with winget...
echo.
winget install --id Python.Python.3.12 -e --scope user --accept-package-agreements --accept-source-agreements
echo.
echo ============================================================
echo   Python installed. CLOSE this window and double-click
echo   setup.bat AGAIN so Windows picks up the new Python.
echo ============================================================
pause
exit /b 0

:no_winget
echo     Install Python from https://www.python.org/downloads/
echo     IMPORTANT: tick "Add python.exe to PATH" in the installer.
echo     Then run setup.bat again.
start "" https://www.python.org/downloads/
pause
exit /b 1

:have_python
for /f "delims=" %%v in ('%PY% --version') do echo [OK] Found %%v

REM ---- 2. Virtual environment + dependencies --------------------------------
if exist ".venv\Scripts\python.exe" goto have_venv
echo [..] Creating virtual environment in .venv
%PY% -m venv .venv
if errorlevel 1 goto fail
:have_venv

echo [..] Installing dependencies (flask, pywin32)...
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 goto fail

".venv\Scripts\python.exe" -c "import flask, win32com.client, pythoncom" >nul 2>&1
if errorlevel 1 goto fail
echo [OK] Dependencies installed

REM ---- 3. Check that desktop PowerPoint is installed -----------------------
reg query "HKCR\PowerPoint.Application" >nul 2>&1
if errorlevel 1 (
    echo [!] Desktop PowerPoint was NOT detected on this PC.
    echo     You need Microsoft PowerPoint - Office 2016+ or Microsoft 365 desktop app.
    echo     PowerPoint Online / web version will not work.
) else (
    echo [OK] PowerPoint detected
)

REM ---- 4. Firewall --------------------------------------------------------
echo.
echo Other devices must be allowed to reach this laptop on port 8000.
choice /c YN /m "Add a Windows Firewall rule now (needs admin, recommended)"
if errorlevel 2 goto skip_fw
call "%~dp0allow_firewall.bat"
goto done
:skip_fw
echo     Skipped. When start.bat runs and Windows Firewall asks, click "Allow access"
echo     and tick BOTH Private and Public networks.

:done
echo.
echo ============================================================
echo   Setup complete!  Now double-click start.bat
echo ============================================================
pause
exit /b 0

:fail
echo.
echo [X] Setup failed. Scroll up to see the error, or see Troubleshooting in README.md
pause
exit /b 1
