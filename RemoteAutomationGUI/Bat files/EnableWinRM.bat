@echo off
setlocal EnableDelayedExpansion

:: Elevation check
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin privileges...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

:: Set log path
set LOGFILE=%windir%\Temp\WinRM_Deploy.log
echo === WinRM Deployment Started: %DATE% %TIME% === >> "%LOGFILE%"

:: Enable WinRM silently
echo Enabling WinRM... >> "%LOGFILE%"
winrm quickconfig -quiet >> "%LOGFILE%" 2>&1

:: Enable PowerShell remoting
echo Enabling PS Remoting... >> "%LOGFILE%"
powershell -Command "Enable-PSRemoting -Force" >> "%LOGFILE%" 2>&1

:: Set WinRM service to auto-start
echo Configuring WinRM service... >> "%LOGFILE%"
sc config WinRM start= auto >> "%LOGFILE%" 2>&1

:: Start the service
echo Starting WinRM service... >> "%LOGFILE%"
net start WinRM >> "%LOGFILE%" 2>&1

:: Verify success
echo Verifying WinRM... >> "%LOGFILE%"
powershell -Command "Test-WSMan -ComputerName localhost" >> "%LOGFILE%" 2>&1

echo === WinRM Deployment Completed: %DATE% %TIME% === >> "%LOGFILE%"
echo Done. Log saved to %LOGFILE%

