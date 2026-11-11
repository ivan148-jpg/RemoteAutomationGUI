param(
    [string]$FilePath,          # Not used, but included for GUI compatibility
    [string[]]$Servers
)

foreach ($srv in $Servers) {
    try {
        $result = Test-WSMan -ComputerName $srv -ErrorAction Stop
        [PSCustomObject]@{
            Server = $srv
            Status = "✅ WinRM OK"
            Output = $result.ProductVersion
        }
    }
    catch {
        [PSCustomObject]@{
            Server = $srv
            Status = "❌ WinRM Failed"
            Output = $_.Exception.Message
        }
    }
}
