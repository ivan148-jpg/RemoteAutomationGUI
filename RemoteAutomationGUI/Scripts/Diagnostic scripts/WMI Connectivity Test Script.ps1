param(
    [string]$FilePath,          # GUI compatibility
    [string[]]$Servers
)

foreach ($srv in $Servers) {
    try {
        $wmiStatus = Invoke-Command -ComputerName $srv -ScriptBlock {
            $result = @{
                WMIAvailable = $false
                NamespaceAccessible = $false
                Message = ""
            }

            try {
                # Basic WMI availability check
                $ping = Get-WmiObject -Namespace "root\cimv2" -Class Win32_OperatingSystem
                if ($ping) {
                    $result.WMIAvailable = $true
                    $result.NamespaceAccessible = $true
                    $result.Message = "WMI and root\cimv2 accessible"
                } else {
                    $result.Message = "WMI query returned null"
                }
            } catch {
                $result.Message = "WMI error: $($_.Exception.Message)"
            }

            return $result
        }

        [PSCustomObject]@{
            Server = $srv
            Status = if ($wmiStatus.WMIAvailable) { "✅ WMI Enabled" } else { "❌ WMI Unavailable" }
            Output = $wmiStatus.Message
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
