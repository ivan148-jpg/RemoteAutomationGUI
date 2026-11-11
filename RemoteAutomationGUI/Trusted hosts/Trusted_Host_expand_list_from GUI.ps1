param(
    [string[]]$Server,   # servers selected in your GUI
    [string]$LogPath = "C:\RemoteAutomationGUI\Logs\TrustedHostsUpdate.log"
)

# === Elevation Auto-Relaunch ===
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    $arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    if ($Server) {
        $arguments += " -Server " + ($Server -join " ")
    }
    if ($LogPath) {
        $arguments += " -LogPath `"$LogPath`""
    }

    # Relaunch elevated and keep window open
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    return
}

try {
    # === GET CURRENT TRUSTEDHOSTS ===
    $current = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value -split ',' | ForEach-Object { $_.Trim() }

    # === MERGE AND DEDUPLICATE ===
    $merged = $current + $Server
    $unique = $merged | Sort-Object -Unique
    $updatedValue = $unique -join ','

    # === UPDATE TRUSTEDHOSTS (LOCAL MACHINE) ===
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value $updatedValue -Force

    # === LOG THE UPDATE ===
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "`n[$timestamp] TrustedHosts updated with:"
    $Server | ForEach-Object { Add-Content -Path $LogPath -Value " - $_" }

    # === Confirmation ===
    Write-Host "✅ TrustedHosts updated. New entries added:" -ForegroundColor Green
    $Server | ForEach-Object { Write-Host " - $_" -ForegroundColor Cyan }
}
catch {
    Write-Host "❌ Failed to update TrustedHosts: $($_.Exception.Message)" -ForegroundColor Red
}
