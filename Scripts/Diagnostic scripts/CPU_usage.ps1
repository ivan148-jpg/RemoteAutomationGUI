param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            # Grab current CPU load percentage
            $usage = Get-Counter '\Processor(_Total)\% Processor Time'
            $currentUsage = [math]::Round($usage.CounterSamples.CookedValue,2)

            # Return formatted string (only usage)
            "$currentUsage %"
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = $out
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
