param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,              # Selected servers from checkbox
    [string]$FilePath = "",         # Not used here, but kept for consistency
    [string]$ExtraParam = ""        # Optional: custom file path
)

$remotePath = if ($ExtraParam) { $ExtraParam } else { "C:\Kit\RemoteTest.txt" }

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            param($path)
            "Test file created on $env:COMPUTERNAME at $(Get-Date)" | Out-File -FilePath $path -Encoding UTF8
            "✅ File created at $path"
        } -ArgumentList $remotePath -ErrorAction Stop

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