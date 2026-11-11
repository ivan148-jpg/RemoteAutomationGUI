param(
    [string]$FilePath,          # GUI compatibility
    [string[]]$Servers
)

foreach ($srv in $Servers) {
    try {
        $info = Invoke-Command -ComputerName $srv -ScriptBlock {
            $remotingUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

            $interactiveUser = try {
                $cs = Get-WmiObject -Class Win32_ComputerSystem
                $cs.UserName
            } catch {
                "Unavailable"
            }

            $osInfo = try {
                Get-WmiObject -Class Win32_OperatingSystem
            } catch {
                $null
            }

            $version   = if ($osInfo) { $osInfo.Caption } else { "Unknown OS" }
            $lastBoot = [System.Management.ManagementDateTimeConverter]::ToDateTime($osInfo.LastBootUpTime)
            $uptimeTS  = (Get-Date) - $lastBoot

            $uptime = "{0}d {1}h {2}m" -f $uptimeTS.Days, $uptimeTS.Hours, $uptimeTS.Minutes

            return @{
                RemotingUser    = $remotingUser
                InteractiveUser = $interactiveUser
                OSVersion       = $version
                LastBoot        = $lastBoot
                Uptime          = $uptime
            }
        }

        [PSCustomObject]@{
            Server = $srv
            Status = "✅ Success"
            Output = "Remote: $($info.RemotingUser) | Logged-in: $($info.InteractiveUser) | OS: $($info.OSVersion) | Boot: $($info.LastBoot) | Uptime: $($info.Uptime)"
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
