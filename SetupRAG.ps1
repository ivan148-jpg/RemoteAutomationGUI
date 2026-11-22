Write-Host "`n=== Remote Automation Setup Starting ===`n" -ForegroundColor Cyan

# Step 1: Enable WinRM
Write-Host "[1/4] Configuring WinRM..." -ForegroundColor Yellow
try {
    winrm quickconfig -quiet | Out-Null
    Write-Host "      Done: WinRM configured successfully.`n" -ForegroundColor Green
} catch {
    Write-Host "      ERROR: WinRM setup failed: $_`n" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Step 2: Set execution policy
Write-Host "[2/4] Setting execution policy to RemoteSigned..." -ForegroundColor Yellow
try {
    Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
    Write-Host "      Done: Execution policy set.`n" -ForegroundColor Green
} catch {
    Write-Host "      WARNING: Failed to set execution policy: $_`n" -ForegroundColor DarkYellow
}

# Step 3: Install SQL Server module
Write-Host "[3/4] Checking and installing SqlServer module (this may take a moment)..." -ForegroundColor Yellow
try {
    if (-not (Get-Module -ListAvailable -Name SqlServer)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
        Install-Module -Name SqlServer -Scope AllUsers -Force -AllowClobber
        Write-Host "      Done: SqlServer module installed.`n" -ForegroundColor Green
    } else {
        Write-Host "      Done: SqlServer module already present.`n" -ForegroundColor Green
    }
} catch {
    Write-Host "      WARNING: Failed to install SqlServer module: $_`n" -ForegroundColor DarkYellow
}


# Step 4: Reminder for TrustedHosts
Write-Host "[4/4] TrustedHosts Notice:" -ForegroundColor Yellow
Write-Host "      Reminder: When you first open the application, use 'Trusted_Host_expand_list_from GUI.ps1' to add your target servers to TrustedHosts.`n" -ForegroundColor Cyan

Write-Host "=== Setup Complete ===" -ForegroundColor Green
Read-Host "Press Enter to continue"
