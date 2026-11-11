param(
    [string]$Server  # Server name from GUI checkbox
)

try {
    $ping = Test-Connection -ComputerName $Server -Count 2 -Quiet -ErrorAction Stop
    if (-not $ping) {
        [PSCustomObject]@{
            Server = $Server
            Status = "❌ No Response"
            Output = "Ping to ${Server}: No response"
        }
        return
    }

    $result = Invoke-Command -ComputerName $Server -ScriptBlock {
        $os = Get-WmiObject -Class Win32_OperatingSystem
$bootTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
$uptime = (Get-Date) - $bootTime
"Uptime: {0} days {1} hours {2} minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes

    } -ErrorAction Stop

    [PSCustomObject]@{
        Server = $Server
        Status = "✅ Reachable"
        Output = $result
    }
}
catch {
    [PSCustomObject]@{
        Server = $Server
        Status = "❌ Error"
        Output = $_.Exception.Message
    }
}
