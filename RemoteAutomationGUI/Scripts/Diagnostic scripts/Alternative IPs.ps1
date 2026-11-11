param(
    [string[]]$Server
)

foreach ($srv in $Server) {
    try {
        $raw = Invoke-Command -ComputerName $srv -ScriptBlock {
            ipconfig /all
        } -ErrorAction Stop

        $adapters = @()
        $currentAdapter = $null

        foreach ($line in $raw) {
            if ($line -match "adapter\s(.+):") {
                $currentAdapter = @{
                    Name = $Matches[1].Trim()
                    IPv4 = @()
                    Mask = @()
                }
                $adapters += $currentAdapter
                continue
            }

            if ($line -match "IPv4 Address.*?:\s*(\d+\.\d+\.\d+\.\d+)") {
                $currentAdapter.IPv4 += $Matches[1]
            }

            if ($line -match "Subnet Mask.*?:\s*(\d+\.\d+\.\d+\.\d+)") {
                $currentAdapter.Mask += $Matches[1]
            }

            if ($line -match "Default Gateway.*?:\s*(\d+\.\d+\.\d+\.\d+)") {
                if (-not $currentAdapter.ContainsKey("Gateway")) {
                    $currentAdapter.Gateway = $Matches[1]
                }
            }

            if ($line -match "DNS Servers.*?:\s*(\d+\.\d+\.\d+\.\d+)") {
                $currentAdapter.DNS = @($Matches[1])
            }
            elseif ($currentAdapter.DNS -and $line -match "^\s+(\d+\.\d+\.\d+\.\d+)") {
                $currentAdapter.DNS += $Matches[1]
            }
        }

        $outputs = @()
        foreach ($adapter in $adapters) {
            for ($i=0; $i -lt $adapter.IPv4.Count; $i++) {
                $type = if ($i -eq 0) { "Primary" } else { "Alternate" }
                $mask = if ($adapter.Mask.Count -gt $i) { $adapter.Mask[$i] } else { "" }
                $gateway = $adapter.Gateway
                $dns = if ($adapter.DNS) { $adapter.DNS -join ", " } else { "" }
                $outputs += "$type ($($adapter.Name)): $($adapter.IPv4[$i]) | Mask: $mask | GW: $gateway | DNS: $dns"
            }
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"       # <- this triggers green in most grids
            Output = ($outputs -join "  ;  ")
        }
    }
    catch {
        [PSCustomObject]@{
            Server = $srv
            Status = "❌ Failed"
            Output = $_.Exception.Message
        }
    }
}
