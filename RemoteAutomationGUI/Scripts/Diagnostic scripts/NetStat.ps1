param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            netstat -an | Select-String "TCP"
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = ($out -join "`r`n")
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
