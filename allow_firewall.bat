@echo off
REM Adds an inbound Windows Firewall rule for TCP port 8000. Asks for admin rights.
net session >nul 2>&1
if errorlevel 1 (
    echo Requesting administrator rights for the firewall rule...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs -Wait"
    exit /b
)
netsh advfirewall firewall delete rule name="PowerPoint Control API" >nul 2>&1
netsh advfirewall firewall add rule name="PowerPoint Control API" dir=in action=allow protocol=TCP localport=8000 profile=any
if errorlevel 1 (
    echo [X] Could not add the firewall rule.
) else (
    echo [OK] Firewall rule "PowerPoint Control API" added - port 8000 is open.
)
timeout /t 3 >nul
