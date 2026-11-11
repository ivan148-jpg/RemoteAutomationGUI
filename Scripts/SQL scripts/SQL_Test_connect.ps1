param (
    [string]$Server,
    [string]$Database = "master",
    [string]$SqlUser = "",
    [string]$SqlPass = "",
    [switch]$UseSqlAuth
)

try {
    if ($UseSqlAuth) {
        $connectionString = "Server=$Server;Database=$Database;User ID=$SqlUser;Password=$SqlPass;Encrypt=True;TrustServerCertificate=True"
    } else {
        $connectionString = "Server=$Server;Database=$Database;Trusted_Connection=True;TrustServerCertificate=True"
    }

    $query = "SELECT GETDATE() AS ServerTime"
    $result = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query -ErrorAction Stop

    [PSCustomObject]@{
        Server = $Server
        Status = "✅ Connected"
        Output = $result.ServerTime
    }
}
catch {
    [PSCustomObject]@{
        Server = $Server
        Status = "❌ Failed"
        Output = $_.Exception.Message
    }
}