param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            Get-NetFirewallProfile | Select-Object Name, Enabled
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = ($out | ForEach-Object { "$($_.Name): $($_.Enabled)" }) -join "`r`n"
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
