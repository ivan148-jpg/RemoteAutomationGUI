param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

foreach ($srv in $Server) {

    Start-Process "cmd.exe" -ArgumentList "/k ping -t $srv"

    [PSCustomObject]@{
        Server = $srv
        Status = "✅ Started"
        Output = "Continuous ping started for $srv"
    }
}