param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = "445"  # Default to SMB port
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            param($port)
            $client = New-Object System.Net.Sockets.TcpClient
            try {
                $client.Connect("localhost", [int]$port)
                "Port $port is open"
            } catch {
                "Port $port is closed"
            } finally {
                $client.Close()
            }
        } -ArgumentList $ExtraParam

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
