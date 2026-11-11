param(
    [Parameter(Mandatory = $true)]
    [string[]]$Server,
    [string]$FilePath = "",
    [string]$ExtraParam = ""
)

# Initialize an array to collect results
$finalResults = @()

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            $drives = @('C', 'D')
            $lowSpace = $false
            $results = foreach ($d in $drives) {
                $drive = Get-PSDrive -Name $d -ErrorAction SilentlyContinue
                if ($drive) {
                    $freeGB = [math]::Round($drive.Free / 1GB)
                    $totalGB = [math]::Round(($drive.Used + $drive.Free) / 1GB)
                    if ($freeGB -lt 10) { $lowSpace = $true }
                    "${d}: ${freeGB} GB free of ${totalGB} GB" + $(if ($freeGB -lt 10) { " ⚠️ LOW SPACE" } else { "" })
                } else {
                    "${d}: ❌ Not found"
                }
            }
            [PSCustomObject]@{
                Output = $results -join " | "
                LowSpace = $lowSpace
            }
        }

        $finalResults += [PSCustomObject]@{
            Server = $srv
            Status = if ($out.LowSpace) { "⚠️ Low space" } else { "✅ Success" }
            Output = $out.Output
        }
    }
    catch {
        $finalResults += [PSCustomObject]@{
            Server = $srv
            Status = "❌ Failed"
            Output = $_.Exception.Message
        }
    }
}

# DO NOT WRITE TO FILE — just return results
return $finalResults