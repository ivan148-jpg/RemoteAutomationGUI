param(
    [string[]]$Server,          # one or more servers from your GUI
    [string]$ExtraParam = ""    # optional: unused here but kept for compatibility
)

foreach ($srv in $Server) {
    try {
        Invoke-Command -ComputerName $srv -ScriptBlock {
            Restart-Computer -Force
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Restarting"
            Output = "Restart command sent successfully"
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