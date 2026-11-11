param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {
    try {
        $ping = Test-Connection -ComputerName $srv -Count 2 -Quiet -ErrorAction Stop
        $status = if ($ping) { "✅ Reachable" } else { "❌ No Response" }

        [PSCustomObject]@{
            Server = $srv
            Status = $status
            Output = "Ping to ${srv}: $status"
        }
    }
    catch {
        [PSCustomObject]@{
            Server = $srv
            Status = "❌ Error"
            Output = $_.Exception.Message
        }
    }
}
