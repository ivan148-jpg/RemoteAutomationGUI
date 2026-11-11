param (
    [string[]]$Server,          # one or more servers from your GUI
    [string]$ExtraParam = ""    # optional
)

foreach ($srv in $Server) {
    try {
        $out = Invoke-Command -ComputerName $srv -ScriptBlock {
            # Detect OS version
            $os = (Get-WmiObject Win32_OperatingSystem).Version
            Write-Verbose "[$env:COMPUTERNAME] OS version detected: $os"

            if ($os -like "6.1*") {
                # Skip Windows 7
                Write-Verbose "[$env:COMPUTERNAME] Skipping BitLocker check (Windows 7 detected)"
                "Windows 7 detected → Skipped"
            }
            else {
                # Windows 8/2012 and newer
                Write-Verbose "[$env:COMPUTERNAME] Branch: Windows 8/2012+"
                $vols = Get-BitLockerVolume -MountPoint "C:","D:" |
                        Select-Object MountPoint, ProtectionStatus
                $vols | ForEach-Object {
                    Write-Verbose "[$env:COMPUTERNAME] Drive $($_.MountPoint) raw ProtectionStatus=$($_.ProtectionStatus)"
                    "$($_.MountPoint) → $($_.ProtectionStatus)"
                }
            }
        } -Verbose   # enable verbose output

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = ($out -join "; ")
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
