param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            $cpu = Get-CimInstance Win32_Processor
            "$($cpu.Name) — $($cpu.NumberOfLogicalProcessors) cores"
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
