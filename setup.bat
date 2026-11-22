@echo off
setlocal
cd /d "%~dp0"
title PowerPoint Control API - Setup

echo ============================================================
echo   PowerPoint Control API - one-time setup
echo ============================================================
echo.

REM ---- 1. Find or install uv --------------------------------------------------
set "UV="
where uv >nul 2>&1
if not errorlevel 1 set "UV=uv"
if defined UV goto have_uv
if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV=%USERPROFILE%\.local\bin\uv.exe"
if defined UV goto have_uv

echo [..] Installing uv - fast Python package manager...
powershell -NoProfile -ExecutionPolicy ByPass -Command "irm https://astral.sh/uv/install.ps1 | iex"
if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV goto fail

:have_uv
for /f "delims=" %%v in ('"%UV%" --version') do echo [OK] Found %%v

REM ---- 2. Python + dependencies (uv downloads Python itself if needed) ------
echo [..] Installing Python and dependencies - flask, pywin32...
"%UV%" sync
if errorlevel 1 goto fail

".venv\Scripts\python.exe" -c "import flask, win32com.client, pythoncom" >nul 2>&1
if errorlevel 1 goto fail
echo [OK] Python and dependencies ready

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
