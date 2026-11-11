param(
    [string[]]$Server,
    [switch]$DebugMode
)

foreach ($srv in $Server) {
    try {
        $result = Invoke-Command -ComputerName $srv -ScriptBlock {
            $log = @()

            # Get all enabled network adapters with IP addresses
            $adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object {
                $_.IPEnabled -eq $true -and $_.IPAddress
            }

            if (-not $adapters) {
                return "⚠ No active IP-enabled adapters found"
            }

        $index = 0
foreach ($adapter in $adapters) {
    foreach ($ip in $adapter.IPAddress) {
        if ($ip -match '^\d{1,3}(\.\d{1,3}){3}$') {  # IPv4 only
            $label = if ($index -eq 0) { "🌐 Primary IP" } else { "🔄 Secondary IP $index" }
           $log += ("{0}: {1}" -f $label, $ip)
            $index++
        }
    }
}

            return $log -join "`n"
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = $result
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($DebugMode) {
            Write-Host "DEBUG [$srv]: $errorMsg" -ForegroundColor Red
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "❌ Failed"
            Output = $errorMsg
        }
    }
}