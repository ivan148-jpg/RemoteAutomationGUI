# Auto-elevate if not running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell -Verb runAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    exit
}

Write-Host "`n=== Remote Automation Setup Starting ===`n"

# Get the folder where the script or .exe is running
$scriptDir = [System.AppDomain]::CurrentDomain.BaseDirectory
$ipFile = Join-Path $scriptDir "TrustedHosts_list.txt"

# Step 1: Enable WinRM
Write-Host "Configuring WinRM..."
try {
    winrm quickconfig -quiet
    Write-Host "WinRM configured successfully.`n"
} catch {
    Write-Warning "WinRM setup failed: $_"
    Read-Host "Press Enter to exit"
    exit
}

# Step 2: Set execution policy
Write-Host "Setting execution policy to RemoteSigned..."
try {
    Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
    Write-Host "Execution policy set.`n"
} catch {
    Write-Warning "Failed to set execution policy: $_"
}

# Step 3: Load plain IPs from file
if (-not (Test-Path $ipFile)) {
    Write-Warning "IP list file not found: $ipFile"
    Read-Host "Press Enter to exit"
    exit
}

$ips = Get-Content $ipFile | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }

if (-not $ips) {
    Write-Warning "No valid IPs found in file."
    Read-Host "Press Enter to exit"
    exit
}



# Step 4: Install SQL Server module
Write-Host "Installing SqlServer module..."
try {
    if (-not (Get-Module -ListAvailable -Name SqlServer)) {
        Install-Module -Name SqlServer -Scope AllUsers -Force -AllowClobber
        Write-Host "SqlServer module installed.`n"
    } else {
        Write-Host "SqlServer module already present.`n"
    }
} catch {
    Write-Warning "Failed to install SqlServer module: $_"
}

Write-Host "`n=== Setup Complete ==="
Read-Host "Press Enter to close this window"


# Step 5: Update TrustedHosts
Write-Host "Updating TrustedHosts..."
try {
    $combined = ($ips | Sort-Object -Unique) -join ','
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $combined
    Write-Host "TrustedHosts updated successfully.`n"
} catch {
    Write-Warning "Failed to update TrustedHosts: $_"
    Read-Host "Press Enter to exit"
    exit
}